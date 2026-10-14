import type { App, IActivityContext } from '@microsoft/teams.apps';
import type { OpenAiAgenticOrchestrator } from './azureOpenAi/azureOpenAiOrchestrator.ts';
import type { IAgentMessage } from './azureOpenAi/IAgentMessage.ts';
import type { createDraftFlow } from './targatedMessage/draftMessage.ts';
import { createChannelUserMessage, getChannelThreadContext } from './teamsContext/channelThreadContext.ts';
import type { IChannelThreadActivity } from './teamsContext/IChannelThreadActivity.ts';

type MessageContext = Parameters<Parameters<App['message']>[1]>[0];

export function createMessageHandler(agent: OpenAiAgenticOrchestrator) {
  return async ({ activity, stream }: MessageContext) => {
    if (activity.conversation?.conversationType !== 'personal') return;

    let streamingStarted = false;
    for await (const update of agent.getUpdates([{ role: 'user', content: activity.text ?? '' }], 'default', true)) {
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

export function createMentionHandler(app: App, agent: OpenAiAgenticOrchestrator, draftFlow: ReturnType<typeof createDraftFlow>) {
  return async ({ activity, api, send }: IActivityContext) => {
    if (activity.type !== 'message' || activity.conversation?.conversationType !== 'channel') return;

    const command = activity.stripMentionsText().text?.trim() ?? '';
    const draftCommand = command.match(/^\/draft(?:\s+([\s\S]*))?$/i);

    if (draftCommand) {
      await api.conversations.addReaction(activity.conversation.id, activity.id, '2728_sparkles').catch(() => {});
      const instructions = draftCommand[1]?.trim() || 'Draft a response to this discussion.';
      const response = await getChannelResponse(app, agent, activity, instructions);
      if (response.text) await draftFlow.sendDraft({ activity, send }, response.text, response.conversation);
      return;
    }

    const threadRootId = activity.conversation.id.match(/;messageid=([^;]+)/)?.[1];
    if (!threadRootId || threadRootId === activity.id) return;

    await api.conversations.addReaction(activity.conversation.id, activity.id, '2728_sparkles').catch(() => {});
    const response = await getChannelResponse(app, agent, activity);
    if (response.text) await send(response.text);
  };
}

async function getChannelResponse(app: App, agent: OpenAiAgenticOrchestrator, activity: IChannelThreadActivity, draftInstructions?: string) {
  const conversation: IAgentMessage[] = await getChannelThreadContext(activity, app.tokenProvider)
    .catch(() => undefined) ?? [];
  conversation.push(createChannelUserMessage(draftInstructions === undefined
    ? activity
    : { ...activity, text: draftInstructions }));

   console.log("------")
   console.log("------")
   console.log("------")
  for (const message of conversation) {
    console.log(message);
    console.log("------")
  }
  let response = '';
  for await (const update of agent.getUpdates(conversation, draftInstructions === undefined ? 'default' : 'draft')) {
    if (!update.informative) response += update.text;
  }
  return { text: response, conversation };
}