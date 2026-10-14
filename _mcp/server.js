const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const { randomUUID } = require("node:crypto");
const { z } = require("zod");
const { textSchema, contractFields, contractSchema, readContracts, writeContracts } = require("./contract-store");

const indexPath = path.join(__dirname, "index.html");
const port = Number(process.env.PORT) || 3000;
const createSchema = z.object({ id: textSchema.optional(), ...contractFields }).strict();
const updateSchema = z.object(contractFields).partial().strict().refine((value) => Object.keys(value).length > 0, "Provide at least one field to update.");

function sendJson(response, status, value) {
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
  response.end(JSON.stringify(value));
}

async function readBody(request) {
  let body = "";
  for await (const chunk of request) {
    body += chunk;
    if (Buffer.byteLength(body) > 65536) {
      throw Object.assign(new Error("Request body is too large."), { status: 413 });
    }
  }
  try {
    return JSON.parse(body);
  } catch {
    throw Object.assign(new Error("Request body must be valid JSON."), { status: 400 });
  }
}

async function handleContracts(request, response, pathname) {
  try {
    const collection = pathname === "/api/contracts" || pathname === "/contracts.json";
    const id = collection ? null : decodeURIComponent(pathname.slice("/api/contracts/".length));
    if (request.method === "GET") {
      const contracts = readContracts();
      const result = collection ? contracts : contracts.find((contract) => contract.id === id);
      sendJson(response, result ? 200 : 404, result || { error: "Contract not found." });
      return;
    }
    if (collection && request.method === "POST" && pathname === "/api/contracts") {
      const input = createSchema.parse(await readBody(request));
      const contract = contractSchema.parse({ ...input, id: input.id ?? randomUUID() });
      const contracts = readContracts();
      if (contracts.some((existing) => existing.id === contract.id)) {
        sendJson(response, 409, { error: "A contract with this ID already exists." });
        return;
      }
      contracts.push(contract);
      writeContracts(contracts);
      sendJson(response, 201, contract);
      return;
    }
    if (!collection && ["PATCH", "DELETE"].includes(request.method)) {
      const updates = request.method === "PATCH" ? updateSchema.parse(await readBody(request)) : null;
      const contracts = readContracts();
      const index = contracts.findIndex((contract) => contract.id === id);
      if (index === -1) {
        sendJson(response, 404, { error: "Contract not found." });
        return;
      }
      if (request.method === "DELETE") {
        contracts.splice(index, 1);
        writeContracts(contracts);
        response.writeHead(204);
        response.end();
      } else {
        contracts[index] = contractSchema.parse({ ...contracts[index], ...updates });
        writeContracts(contracts);
        sendJson(response, 200, contracts[index]);
      }
      return;
    }
    response.setHeader("Allow", collection ? "GET, POST" : "GET, PATCH, DELETE");
    sendJson(response, 405, { error: "Method not allowed." });
  } catch (error) {
    const status = error instanceof z.ZodError || error instanceof URIError ? 400 : error.status || 500;
    const message = error instanceof z.ZodError ? error.issues.map((issue) => `${issue.path.join(".") || "Contract"}: ${issue.message}`).join("; ") : status === 500 ? "Unable to access contracts." : error.message;
    sendJson(response, status, { error: message });
  }
}

const server = http.createServer((request, response) => {
  const requestUrl = new URL(request.url, "http://localhost");
  if (requestUrl.pathname === "/contracts.json" || requestUrl.pathname === "/api/contracts" || requestUrl.pathname.startsWith("/api/contracts/")) {
    handleContracts(request, response, requestUrl.pathname);
    return;
  }

  if (requestUrl.pathname === "/lucide.js") {
    fs.readFile(path.join(__dirname, "node_modules/lucide/dist/umd/lucide.min.js"), (error, script) => {
      response.writeHead(error ? 500 : 200, { "Content-Type": "application/javascript; charset=utf-8" });
      response.end(error ? "" : script);
    });
    return;
  }

  if (requestUrl.pathname !== "/" && requestUrl.pathname !== "/index.html") {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
    return;
  }

  fs.readFile(indexPath, (error, html) => {
    if (error) {
      response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Unable to load Zava Contract Management.");
      return;
    }

    response.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
    response.end(html);
  });
});

server.listen(port, () => {
  console.log(`Zava Contract Management running at http://localhost:${port}`);
});
