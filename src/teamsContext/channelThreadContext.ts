import type { App } from '@microsoft/teams.apps';
import type { IAgentMessage } from '../azureOpenAi/IAgentMessage.ts';
import type { IChannelThreadActivity } from './IChannelThreadActivity.ts';
import type { IThreadMessage } from './IThreadMessage.ts';

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const GRAPH_API_BASE_URL = 'https://graph.microsoft.com/v1.0';

export function createChannelUserMessage(activity: IChannelThreadActivity): IAgentMessage {
  return {
    role: 'user',
    content: JSON.stringify({
      author: { id: activity.from.aadObjectId ?? activity.from.id, displayName: activity.from.name },
      createdDateTime: activity.timestamp,
      text: activity.text ?? '',
    }),
  };
}

export async function getChannelThreadContext(activity: IChannelThreadActivity, tokenProvider: App['tokenProvider']): Promise<IAgentMessage[] | undefined> {
  const teamId = activity.team?.aadGroupId;
  const channelId = activity.channel?.id;
  const rootId = activity.conversation.id.match(/;messageid=([^;]+)/)?.[1];
  if (!teamId || !channelId || !rootId || rootId === activity.id) return undefined;

  const token = await tokenProvider.getAppToken(GRAPH_SCOPE);
  if (!token) throw new Error('Could not acquire an app token for Microsoft Graph');

  const baseUrl = `${GRAPH_API_BASE_URL}/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages/${encodeURIComponent(rootId)}`;
  const get = async <T>(url: string): Promise<T> => {
    const response = await fetch(url, { headers: { Authorization: `Bearer ${token.toString()}` } });
    if (!response.ok) throw new Error(`Graph thread request failed: ${response.status} ${await response.text()}`);
    return response.json() as Promise<T>;
  };

  const root = await get<IThreadMessage>(baseUrl);
  const replies: IThreadMessage[] = [];
  let nextUrl: string | undefined = `${baseUrl}/replies`;
  while (nextUrl) {
    const page: { value: IThreadMessage[]; '@odata.nextLink'?: string } = await get(nextUrl);
    replies.push(...page.value);
    nextUrl = page['@odata.nextLink'];
  }

  const currentTime = activity.timestamp ? Date.parse(activity.timestamp.toString()) : Infinity;
  return [root, ...replies]
    .filter((message) => message.id !== activity.id && Date.parse(message.createdDateTime) < currentTime)
    .sort((left, right) => Date.parse(left.createdDateTime) - Date.parse(right.createdDateTime))
    .map((message) => ({
      role: message.from?.application ? 'assistant' as const : 'user' as const,
      content: JSON.stringify({
        author: message.from?.application ?? message.from?.user ?? null,
        createdDateTime: message.createdDateTime,
        text: message.body?.content ?? '',
      }),
    }));
}