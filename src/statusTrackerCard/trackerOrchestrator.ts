import OpenAI from 'openai';
import type { IAgentMessage } from '../azureOpenAi/IAgentMessage.ts';

const ACTION_ITEM_COUNT = 5;

const TRACKER_PROMPT = `You will be given a conversation thread from a Teams channel. Each user message is a JSON object with the author, timestamp and text of one message in the thread.
Based on that conversation, respond with exactly ${ACTION_ITEM_COUNT} action items.
Each action item must be a short, concrete, single-sentence task.
Respond only with JSON in this shape: {"actionItems": ["...", "...", "...", "...", "..."]}`;

export class TrackerOrchestrator {
  private readonly openai: OpenAI;
  private readonly model: string;

  constructor(settings: { Endpoint: string; ApiKey: string; Deployment: string }) {
    this.openai = new OpenAI({ apiKey: settings.ApiKey, baseURL: settings.Endpoint });
    this.model = settings.Deployment;
  }

  async getActionItems(conversation: IAgentMessage[]): Promise<string[]> {
    const response = await this.openai.chat.completions.create({
      model: this.model,
      messages: [{ role: 'system', content: TRACKER_PROMPT }, ...conversation],
      response_format: { type: 'json_object' },
    });

    const content = response.choices[0]?.message?.content ?? '{}';
    const items: unknown = (JSON.parse(content) as { actionItems?: unknown }).actionItems;
    if (!Array.isArray(items)) return [];

    return items
      .filter((item): item is string => typeof item === 'string')
      .map((item) => item.trim())
      .filter((item) => item.length > 0)
      .slice(0, ACTION_ITEM_COUNT);
  }
}
