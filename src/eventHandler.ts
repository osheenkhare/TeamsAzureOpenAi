import { MessageActivityInput, stripMentionsText, type ConversationType, type IMessageActivity } from '@microsoft/teams.api';
import type { App, IActivityContext } from '@microsoft/teams.apps';
import type { OpenAiAgenticOrchestrator } from './azureOpenAi/azureOpenAiOrchestrator.ts';
import type { IAgentMessage } from './azureOpenAi/IAgentMessage.ts';
import { createChannelUserMessage, getChannelThreadContext } from './teamsContext/channelThreadContext.ts';
import { createStatusTrackerCard } from './statusTrackerCard/statusTrackerCard.ts';
import type { TrackerOrchestrator } from './statusTrackerCard/trackerOrchestrator.ts';
import { createMcpWidgetMessage } from './widget/widgetHtmlMcp.ts';

// Context type the SDK passes to app.message handlers.
type MessageContext = Parameters<Parameters<App['message']>[1]>[0];

// Activity/conversation values typed against the Teams SDK.
const MESSAGE_ACTIVITY: IMessageActivity['type'] = 'message';
const PERSONAL_CONVERSATION: ConversationType = 'personal';
const CHANNEL_CONVERSATION: ConversationType = 'channel';

// Reaction added to a mention to acknowledge it while the agent works.
const SPARKLES_REACTION = '2728_sparkles';

export function createInstallHandler() {
  return async ({ activity, send }: IActivityContext) => {
    if (activity.conversation?.conversationType !== PERSONAL_CONVERSATION) return;

    await send(
      '\u{1F44B} **Welcome to Zava Deal Desk!**\n\n' +
      'Your partner for customer contract reviews. I can help you:\n\n' +
      '- \u{1F4DD} **Review** customer contract requests and summarize redlines.\n' +
      '- \u{1F50D} **Identify** missing information and potential bottlenecks.\n' +
      '- \u{2705} **Clarify** approval requirements and next steps.\n\n' +
      '\u{1F4AC} Share a **contract question** or the **details of a deal** whenever you are ready.'
    );
  };
}

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
  return async ({ activity, api, send, stream }: IActivityContext) => {
    if (activity.type !== MESSAGE_ACTIVITY) return;
    if (activity.conversation?.conversationType !== CHANNEL_CONVERSATION) return;

    // Channel conversation ids end with ';messageid=<rootId>'; without it there is no thread to reply into.
    const threadRootId = activity.conversation.id.match(/;messageid=([^;]+)/)?.[1];
    if (!threadRootId) return;    

    // Best-effort acknowledgement; a failed reaction should not block the reply.
    await api.conversations
      .addReaction(activity.conversation.id, activity.id, SPARKLES_REACTION)
      .catch(() => {});

    // Render the widget if the user explicitly requests it.
    if (stripMentionsText(activity)?.trim().toLowerCase() === 'widget') {
      await send(await createMcpWidgetMessage());
      return;
    }

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

    let streamingStarted = false;
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