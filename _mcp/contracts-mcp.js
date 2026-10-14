const fs = require("node:fs");
const http = require("node:http");
const path = require("node:path");
const { randomUUID } = require("node:crypto");
const { McpServer } = require("@modelcontextprotocol/sdk/server/mcp.js");
const { StreamableHTTPServerTransport } = require("@modelcontextprotocol/sdk/server/streamableHttp.js");
const { textSchema, contractFields, contractSchema, readContracts, writeContracts } = require("./contract-store");

const port = Number(process.env.MCP_PORT) || 3001;
const sessions = new Map();

function withErrors(handler) {
  return async (...args) => {
    try {
      const value = await handler(...args);
      return { content: [{ type: "text", text: value }] };
    } catch (error) {
      return {
        isError: true,
        content: [{ type: "text", text: error.message }]
      };
    }
  };
}

function createMcpServer() {
  const server = new McpServer({ name: "zava-contracts-mcp", version: "1.0.0" });

  server.registerTool(
    "get_all_contracts",
    { title: "Get All Contracts", description: "Return all Zava contracts." },
    withErrors(() => JSON.stringify(readContracts(), null, 2))
  );

  server.registerTool(
    "get_payment_policy",
    {
      title: "Get Payment Policy",
      description: "Return the payment policy for a payment term.",
      inputSchema: { paymentTerm: contractFields.paymentTerm }
    },
    withErrors(({ paymentTerm }) => {
      const policies = JSON.parse(fs.readFileSync(path.join(__dirname, "payment-policy.json"), "utf8"));
      return JSON.stringify(policies[paymentTerm], null, 2);
    })
  );

  server.registerTool(
    "create_contract",
    {
      title: "Create Contract",
      description: "Create a contract. Provide a unique text ID or omit it to generate a UUID.",
      inputSchema: { id: textSchema.optional(), ...contractFields }
    },
    withErrors((input) => {
      const contracts = readContracts();
      const contract = contractSchema.parse({ ...input, id: input.id ?? randomUUID() });
      if (contracts.some((existing) => existing.id === contract.id)) {
        throw new Error(`Contract ${contract.id} already exists.`);
      }
      contracts.push(contract);
      writeContracts(contracts);
      return JSON.stringify(contract, null, 2);
    })
  );

  server.registerTool(
    "update_contract",
    {
      title: "Update Contract",
      description: "Update provided fields of a contract by its immutable ID. approvalStatus replaces the full approval list.",
      inputSchema: {
        id: textSchema,
        ...Object.fromEntries(Object.entries(contractFields).map(([key, schema]) => [key, schema.optional()]))
      }
    },
    withErrors(({ id, ...updates }) => {
      const providedUpdates = Object.fromEntries(Object.entries(updates).filter(([, value]) => value !== undefined));
      if (Object.keys(providedUpdates).length === 0) {
        throw new Error("Provide at least one field to update.");
      }
      const contracts = readContracts();
      const index = contracts.findIndex((contract) => contract.id === id);
      if (index === -1) {
        throw new Error(`Contract ${id} was not found.`);
      }
      const contract = contractSchema.parse({ ...contracts[index], ...providedUpdates });
      contracts[index] = contract;
      writeContracts(contracts);
      return JSON.stringify(contract, null, 2);
    })
  );

  return server;
}

const httpServer = http.createServer((request, response) => {
  if (new URL(request.url, "http://localhost").pathname !== "/mcp") {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
    return;
  }

  handleMcpRequest(request, response).catch((error) => {
    process.stderr.write(`${error.message}\n`);
    if (!response.headersSent) {
      response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("MCP request failed.");
    }
  });
});

async function handleMcpRequest(request, response) {
  const sessionIdHeader = request.headers["mcp-session-id"];
  const sessionId = typeof sessionIdHeader === "string" ? sessionIdHeader : undefined;
  let session = sessionId ? sessions.get(sessionId) : undefined;
  const isNewSessionRequest = !sessionId && request.method === "POST";

  if (isNewSessionRequest) {
    const id = randomUUID();
    const server = createMcpServer();
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: () => id });
    session = { server, transport };
    sessions.set(id, session);
    transport.onclose = () => sessions.delete(id);
    await server.connect(transport);
  } else if (!session) {
    response.writeHead(sessionId ? 404 : 400, { "Content-Type": "application/json" });
    response.end(JSON.stringify({
      jsonrpc: "2.0",
      error: {
        code: sessionId ? -32001 : -32000,
        message: sessionId ? "Session not found" : "Mcp-Session-Id header is required."
      },
      id: null
    }));
    return;
  }

  try {
    await session.transport.handleRequest(request, response);
  } finally {
    if (isNewSessionRequest && !session.transport.sessionId) {
      await session.transport.close();
      await session.server.close();
    } else if (request.method === "DELETE" && sessionId) {
      sessions.delete(sessionId);
      await session.transport.close();
      await session.server.close();
    }
  }
}

httpServer.listen(port, "127.0.0.1", () => {
  process.stderr.write(`Zava contracts MCP running at http://127.0.0.1:${port}/mcp\n`);
});