import { MessageActivityInput, stripMentionsText, type ConversationType, type IMessageActivity } from '@microsoft/teams.api';
import type { App, IActivityContext } from '@microsoft/teams.apps';
import type { OpenAiAgenticOrchestrator } from './azureOpenAi/azureOpenAiOrchestrator.ts';
import type { IAgentMessage } from './azureOpenAi/IAgentMessage.ts';
import { createChannelUserMessage, getChannelThreadContext } from './teamsContext/channelThreadContext.ts';
import { createStatusTrackerCard } from './statusTrackerCard/statusTrackerCard.ts';
import type { TrackerOrchestrator } from './statusTrackerCard/trackerOrchestrator.ts';

// Context type the SDK passes to app.message handlers.
type MessageContext = Parameters<Parameters<App['message']>[1]>[0];

// Activity/conversation values typed against the Teams SDK.
const MESSAGE_ACTIVITY: IMessageActivity['type'] = 'message';
const PERSONAL_CONVERSATION: ConversationType = 'personal';
const CHANNEL_CONVERSATION: ConversationType = 'channel';

// Reaction added to a mention to acknowledge it while the agent works.
const SPARKLES_REACTION = '2728_sparkles';

// Handles 1:1 chats, streaming the agent's reply as it is generated.
export function createMessageHandler(agent: OpenAiAgenticOrchestrator) {
  return async ({ activity, stream }: MessageContext) => {
    if (activity.conversation?.conversationType !== PERSONAL_CONVERSATION) return;

    // 1:1 chats carry no history; only the current message is sent to the agent.
    const conversation: IAgentMessage[] = [{ role: 'user', content: activity.text ?? '' }];
    let streamingStarted = false;

    // Show status updates until real content starts, then stream the reply text.
    for await (const update of agent.getUpdates(conversation, true)) {
      if (update.informative) {
        if (!streamingStarted) stream.update(update.text);
      } else if (update.text.length > 0) {
        streamingStarted = true;
        stream.emit(update.text);
      }
    }

    await stream.close();
  };
}

// Handles @mentions in channel posts and thread replies.
export function createMentionHandler(app: App, agent: OpenAiAgenticOrchestrator, tracker: TrackerOrchestrator) {
  return async ({ activity, api, send }: IActivityContext) => {
    if (activity.type !== MESSAGE_ACTIVITY) return;
    if (activity.conversation?.conversationType !== CHANNEL_CONVERSATION) return;

    // Channel conversation ids end with ';messageid=<rootId>'; without it there is no thread to reply into.
    const threadRootId = activity.conversation.id.match(/;messageid=([^;]+)/)?.[1];
    if (!threadRootId) return;

    // Best-effort acknowledgement; a failed reaction should not block the reply.
    await api.conversations
      .addReaction(activity.conversation.id, activity.id, SPARKLES_REACTION)
      .catch(() => {});

    // Earlier thread messages (empty for a top-level post or if Graph fails), then the new message.
    const conversation: IAgentMessage[] =
      (await getChannelThreadContext(activity, app.tokenProvider).catch(() => undefined)) ?? [];

    if (stripMentionsText(activity)?.trim().toLowerCase() === 'card') {
      const actionItems = await tracker.getActionItems(conversation);
      const tasks = actionItems.map((text) => ({ text, done: false }));
      await send(new MessageActivityInput().addCard('adaptive', createStatusTrackerCard(tasks)));
      return;
    }

    conversation.push(createChannelUserMessage(activity));

    // Channels don't support streaming, so collect the full reply and send it once.
    let response = '';
    for await (const update of agent.getUpdates(conversation)) {
      if (!update.informative) response += update.text;
    }

    if (response) await send(response);
  };
}