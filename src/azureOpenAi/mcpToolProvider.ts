import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import type OpenAI from 'openai';

export class McpToolProvider {
  private constructor(
    private readonly client: Client,
    readonly tools: OpenAI.Chat.ChatCompletionTool[],
  ) {}

  static async connect(url: string) {
    const client = new Client({ name: 'ZavaContractMcp', version: '1.0.0' });
    await client.connect(new StreamableHTTPClientTransport(new URL(url)));

    const tools: OpenAI.Chat.ChatCompletionTool[] = (await client.listTools()).tools.map((tool) => ({
      type: 'function',
      function: {
        name: tool.name,
        description: tool.description,
        parameters: tool.inputSchema,
      },
    }));

    return new McpToolProvider(client, tools);
  }

  async callTool(name: string, args: string): Promise<string> {
    const result = await this.client.callTool({
      name,
      arguments: JSON.parse(args) as Record<string, unknown>,
    });
    return JSON.stringify(result.content);
  }
}
