import { readFile } from 'node:fs/promises';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';
import OpenAI from 'openai';
import type { IAgentEvent } from './IAgentEvent.ts';
import type { IAgentMessage } from './IAgentMessage.ts';

const MCP_URL = 'http://127.0.0.1:3001/mcp';

type ToolCall = { id: string; name: string; arguments: string };

export class OpenAiAgenticOrchestrator {
  private constructor(
    private readonly openai: OpenAI,
    private readonly model: string,
    private readonly mcp: Client,
    private readonly tools: OpenAI.Chat.ChatCompletionTool[],
    private readonly basePrompt: string,
    private readonly draftPrompt: string,
  ) {}

  static async create(settings: { Endpoint: string; ApiKey: string; Deployment: string }) {
    const basePrompt = await readFile('src/azureOpenAi/agentInstruction.md', 'utf8');
    const draftPrompt = await readFile('src/azureOpenAi/agentDraft.md', 'utf8');
    const mcp = new Client({ name: 'ZavaContractMcp', version: '1.0.0' });
    await mcp.connect(new StreamableHTTPClientTransport(new URL(MCP_URL)));

    const tools: OpenAI.Chat.ChatCompletionTool[] = (await mcp.listTools()).tools.map((tool) => ({
      type: 'function',
      function: {
        name: tool.name,
        description: tool.description,
        parameters: tool.inputSchema,
      },
    }));

    return new OpenAiAgenticOrchestrator(
      new OpenAI({ apiKey: settings.ApiKey, baseURL: settings.Endpoint }),
      settings.Deployment,
      mcp,
      tools,
      basePrompt,
      draftPrompt,
    );
  }

  async *getUpdates(conversation: IAgentMessage[], mode: 'default' | 'draft' = 'default', streaming = false): AsyncGenerator<IAgentEvent> {
    const messages: OpenAI.Chat.ChatCompletionMessageParam[] = [
      { role: 'system', content: mode === 'draft' ? `${this.basePrompt}\n\n${this.draftPrompt}` : this.basePrompt },
      ...conversation,
    ];

    yield { text: 'Thinking...', informative: true };

    while (true) {
      const pending = new Map<number, ToolCall>();
      const request = {
        model: this.model,
        messages,
        tools: mode === 'draft' ? undefined : this.tools,
      };

      if (streaming) {
        const response = await this.openai.chat.completions.create({ ...request, stream: true });
        for await (const chunk of response) {
          const delta = chunk.choices[0]?.delta;
          if (delta?.content) {
            yield { text: delta.content };
          }

          for (const call of delta?.tool_calls ?? []) {
            const current = pending.get(call.index) ?? { id: '', name: '', arguments: '' };
            current.id += call.id ?? '';
            current.name += call.function?.name ?? '';
            current.arguments += call.function?.arguments ?? '';
            pending.set(call.index, current);
          }
        }
      } else {
        const response = await this.openai.chat.completions.create({ ...request, stream: false });
        const message = response.choices[0]?.message;
        if (message?.content) yield { text: message.content };
        for (const [index, call] of (message?.tool_calls ?? []).entries()) {
          if (call.type === 'function') {
            pending.set(index, { id: call.id, name: call.function.name, arguments: call.function.arguments });
          }
        }
      }

      if (pending.size === 0) break;

      const calls = [...pending.entries()].sort(([left], [right]) => left - right).map(([, call]) => call);
      messages.push({
        role: 'assistant',
        tool_calls: calls.map((call) => ({
          id: call.id,
          type: 'function',
          function: { name: call.name, arguments: call.arguments },
        })),
      });

      for (const call of calls) {
        yield { text: `Calling ${call.name}`, informative: true };
        const result = await this.mcp.callTool({
          name: call.name,
          arguments: JSON.parse(call.arguments) as Record<string, unknown>,
        });
        messages.push({ role: 'tool', tool_call_id: call.id, content: JSON.stringify(result.content) });
      }
    }
  }
}