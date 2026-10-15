import { readFile } from 'node:fs/promises';
import OpenAI from 'openai';
import type { McpToolProvider } from './mcpToolProvider.ts';
import type { IAgentEvent } from './IAgentEvent.ts';
import type { IAgentMessage } from './IAgentMessage.ts';

const AGENT_INSTRUCTION_PATH = 'src/azureOpenAi/agentInstruction.md';

type ToolCall = { id: string; name: string; arguments: string };
type Message = OpenAI.Chat.ChatCompletionMessageParam;

export class OpenAiAgenticOrchestrator {
  private constructor(
    private readonly openai: OpenAI,
    private readonly model: string,
    private readonly basePrompt: string,
    private readonly mcp?: McpToolProvider,
  ) {}

  static async create(
    settings: { Endpoint: string; ApiKey: string; Deployment: string },
    mcp?: McpToolProvider,
  ) {
    const basePrompt = await readFile(AGENT_INSTRUCTION_PATH, 'utf8');

    return new OpenAiAgenticOrchestrator(
      new OpenAI({ apiKey: settings.ApiKey, baseURL: settings.Endpoint }),
      settings.Deployment,
      basePrompt,
      mcp,
    );
  }

  // OpenAI rejects an empty tools array, so omit it when MCP is disabled.
  private get tools() {
    return this.mcp?.tools.length ? this.mcp.tools : undefined;
  }

  async *getUpdates(conversation: IAgentMessage[], streaming = false): AsyncGenerator<IAgentEvent> {
    const messages: Message[] = [{ role: 'system', content: this.basePrompt }, ...conversation];

    yield { text: 'Thinking...', informative: true };

    while (true) {
      const calls = streaming ? yield* this.streamCompletion(messages) : yield* this.complete(messages);
      if (calls.length === 0) return;

      messages.push(toAssistantMessage(calls));

      for (const call of calls) {
        yield { text: `Calling ${call.name}`, informative: true };
        messages.push(await this.runTool(call));
      }
    }
  }

  private async *streamCompletion(messages: Message[]): AsyncGenerator<IAgentEvent, ToolCall[]> {
    const response = await this.openai.chat.completions.create({
      model: this.model,
      messages,
      tools: this.tools,
      stream: true,
    });

    // Tool calls arrive in fragments keyed by index and must be stitched together.
    const pending = new Map<number, ToolCall>();
    for await (const chunk of response) {
      const delta = chunk.choices[0]?.delta;
      if (delta?.content) yield { text: delta.content };

      for (const call of delta?.tool_calls ?? []) {
        const current = pending.get(call.index) ?? { id: '', name: '', arguments: '' };
        current.id += call.id ?? '';
        current.name += call.function?.name ?? '';
        current.arguments += call.function?.arguments ?? '';
        pending.set(call.index, current);
      }
    }

    return [...pending.entries()].sort(([a], [b]) => a - b).map(([, call]) => call);
  }

  private async *complete(messages: Message[]): AsyncGenerator<IAgentEvent, ToolCall[]> {
    const response = await this.openai.chat.completions.create({
      model: this.model,
      messages,
      tools: this.tools,
      stream: false,
    });

    const message = response.choices[0]?.message;
    if (message?.content) yield { text: message.content };

    return (message?.tool_calls ?? []).flatMap((call) =>
      call.type === 'function'
        ? [{ id: call.id, name: call.function.name, arguments: call.function.arguments }]
        : [],
    );
  }

  private async runTool(call: ToolCall): Promise<Message> {
    const content = this.mcp
      ? await this.mcp.callTool(call.name, call.arguments)
      : `Tool ${call.name} is not available.`;
    return { role: 'tool', tool_call_id: call.id, content };
  }
}

function toAssistantMessage(calls: ToolCall[]): Message {
  return {
    role: 'assistant',
    tool_calls: calls.map((call) => ({
      id: call.id,
      type: 'function',
      function: { name: call.name, arguments: call.arguments },
    })),
  };
}