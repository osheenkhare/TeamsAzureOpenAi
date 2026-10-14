# Ignite Lab Teams bot

TypeScript Teams bot that streams Azure OpenAI responses and calls tools from the local MCP server.

## Private drafts

In a channel, send `@bot /draft <instructions>` to request a private draft. Instructions are optional; `/draft` alone asks for a response to the discussion. The request is public, but the generated draft is sent only to you in an Adaptive Card. Enter revision instructions and select **Revise** to update that same private card. Select **Publish** to post the latest draft publicly in the same conversation, without buttons or inputs. Publishing makes no LLM call. Normal mentions are unchanged.

Drafts use the short instructions in `src/azureOpenAi/agentDraft.md`, retain the original discussion context for revisions, and do not call MCP tools. Drafts are held in memory and expire after 24 hours. Restarting the bot invalidates pending drafts. Adaptive Cards support only a subset of Markdown, so tables may appear as plain text in the preview.

## Run

Requires Node.js 22, an Azure Bot registration, an Azure OpenAI deployment, and the MCP server at `http://127.0.0.1:3001/mcp` running before the bot starts.

Keep `.env` in the project root (it is ignored by Git). It needs `CLIENT_ID`, `TENANT_ID`, `CLIENT_SECRET`, `AZURE_OPENAI_ENDPOINT` (OpenAI-compatible `/openai/v1` URL), `AZURE_OPENAI_API_KEY`, and `AZURE_OPENAI_DEPLOYMENT`. `PORT` defaults to `3978`. The bot reads `src/azureOpenAi/agentInstruction.md` and `src/azureOpenAi/agentDraft.md` at startup.

Run `npm install`, then `npm run build` and `npm start` from the project root. For development use `npm run dev`; it restarts the bot when TypeScript changes. Restart the bot after changing prompt files. The Teams messaging endpoint is `/api/messages` on the configured port.