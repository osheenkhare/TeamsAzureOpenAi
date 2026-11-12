import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import type { IHtmlWidgetPayload, IMessageActivity } from '@microsoft/teams.api';
const MCP_URL = 'http://127.0.0.1:3001/mcp';

export async function createMcpWidgetMessage(): Promise<Pick<IMessageActivity, 'type' | 'textFormat' | 'text'>> {
  const client = new Client({ name: 'TeamsWidget', version: '1.0.0' }, {
    capabilities: {
      extensions: { 'io.modelcontextprotocol/ui': { mimeTypes: ['text/html;profile=mcp-app'] } },
    },
  });

  try {
    await client.connect(new StreamableHTTPClientTransport(new URL(MCP_URL)));

    const result = await client.callTool({ name: 'widget', arguments: {} });
    const resource = await client.readResource({ uri: 'ui://widget/active-contracts.html' });
    const content = resource.contents[0];
    if (result.isError || !content || !('text' in content)) throw new Error('MCP widget failed.');

    const widget: IHtmlWidgetPayload = {
      type: 'widget/mcp-ui',
      name: 'Zava mcp-html Widget',
      domain: 'https://zava',
      mcpUrl: MCP_URL,
      html: content.text,
      toolInput: {},
      toolOutput: result,
      securityPolicy: { connectDomains: [], resourceDomains: [], frameDomains: [], baseUriDomains: [] },
    };
    const json = JSON.stringify(widget).replace(/`/g, '\\u0060');
    return { type: 'message', textFormat: 'extendedmarkdown', text: '```html-widget\n' + json + '\n```' };
  } catch (error) {
    return {
      type: 'message',
      textFormat: 'plain',
      text: 'MCP server is not available, make sure Zava Contract Services are running.',
    };
  } finally {
    await client.close().catch(() => {});
  }
}