# Simple Testing MCP Server

A simple local testing project with an MCP server and a web UI for contract management. Intended for development and experimentation, not production use.

## Setup

Install Node.js and npm, then install dependencies:

```sh
npm install
```

## Run Locally

Start each service manually in 2 separate terminals.

Start the web UI:

```sh
npm run ui
```

Open http://localhost:3000 in your browser.

In a separate terminal, start the MCP server:

```sh
npm run mcp
```

Connect your MCP client using Streamable HTTP at http://127.0.0.1:3001/mcp.

The UI port can be changed with `PORT`, and the MCP port with `MCP_PORT`.