import { randomUUID } from 'node:crypto';
import type { App, IActivityContext } from '@microsoft/teams.apps';
import { MessageActivityInput } from '@microsoft/teams.api';
import type { AdaptiveCardActionCardResponse, AdaptiveCardActionMessageResponse, IAdaptiveCardActionInvokeActivity } from '@microsoft/teams.api';
import { draftCard, processingCard } from './adaptiveCardTemplate.ts';
import type { OpenAiAgenticOrchestrator } from '../azureOpenAi/azureOpenAiOrchestrator.ts';
import type { IAgentMessage } from '../azureOpenAi/IAgentMessage.ts';

type PublishContext = IActivityContext<IAdaptiveCardActionInvokeActivity>;

type Draft = {
  text: string;
  card: AdaptiveCardActionCardResponse['value'];
  conversation: IAgentMessage[];
  messageId?: string;
  ownerId: string;
  tenantId: string | undefined;
  conversationId: string;
  expiresAt: number;
  status: 'pending' | 'revising' | 'publishing' | 'published' | 'failed';
  publishedMessageId?: string;
};

const DRAFT_LIFETIME_MS = 24 * 60 * 60 * 1000;

function actionMessage(value: string): AdaptiveCardActionMessageResponse {
  return { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value };
}

export function createDraftFlow(app: App, agent: OpenAiAgenticOrchestrator) {
  const drafts = new Map<string, Draft>();

  function getOwnedDraft(activity: PublishContext['activity']) {
    const draftId = activity.value.action.data?.draftId;
    if (typeof draftId !== 'string') return undefined;
    const draft = drafts.get(draftId);
    if (!draft || (activity.from.aadObjectId ?? activity.from.id) !== draft.ownerId ||
        activity.channelData?.tenant?.id !== draft.tenantId ||
        activity.conversation.id.split(';')[0] !== draft.conversationId.split(';')[0]) return undefined;
    return { draft, draftId };
  }

  async function reviseDraft(draft: Draft, draftId: string, instructions: string, { api, log }: PublishContext) {
    const conversation: IAgentMessage[] = [
      ...draft.conversation,
      { role: 'assistant', content: draft.text },
      { role: 'user', content: instructions },
    ];
    let card = draft.card;
    try {
      let text = '';
      for await (const update of agent.getUpdates(conversation, 'draft')) {
        if (!update.informative) text += update.text;
      }
      if (!text.trim()) throw new Error('Draft revision was empty');
      card = draftCard(text, draftId);
      await api.conversations.updateTargetedActivity(draft.conversationId, draft.messageId!, {
        type: 'message', attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: card }],
      });
      draft.text = text;
      draft.card = card;
      draft.conversation = conversation;
    } catch (error) {
      log.error('Draft revision failed', error);
      await api.conversations.updateTargetedActivity(draft.conversationId, draft.messageId!, {
        type: 'message',
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: { ...draft.card, body: [...draft.card.body ?? [], { type: 'TextBlock', text: 'Could not revise the draft. Please try again.', wrap: true }] },
        }],
      }).catch((updateError) => log.error('Could not update draft card', updateError));
    } finally {
      draft.status = 'pending';
    }
  }

  return {
    async sendDraft({ activity, send }: Pick<IActivityContext, 'activity' | 'send'>, text: string, conversation: IAgentMessage[]) {
      const draftId = randomUUID();
      const card = draftCard(text, draftId);
      const draft: Draft = {
        text,
        card,
        conversation,
        ownerId: activity.from.aadObjectId ?? activity.from.id,
        tenantId: activity.channelData?.tenant?.id,
        conversationId: activity.conversation.id,
        expiresAt: Date.now() + DRAFT_LIFETIME_MS,
        status: 'pending',
      };
      drafts.set(draftId, draft);
      try {
        const sent = await send(new MessageActivityInput()
          .addCard('adaptive', card)
          .withRecipient(activity.from, true));
        draft.messageId = sent.id;
      } catch (error) {
        drafts.delete(draftId);
        throw error;
      }
    },

    async revise(context: PublishContext) {
      const { activity } = context;
      if (activity.value?.trigger !== 'manual') return actionMessage('Select Revise to update this draft.');
      const owned = getOwnedDraft(activity);
      if (!owned) return actionMessage('This draft is unavailable. Request a new draft.');
      const { draft, draftId } = owned;
      const instructions = activity.value.action.data?.revisionInstructions;
      if (typeof instructions !== 'string' || !instructions.trim()) return actionMessage('Enter revision instructions.');
      if (draft.status !== 'pending' || !draft.messageId) return actionMessage('This draft is busy or already published.');
      draft.status = 'revising';
      void reviseDraft(draft, draftId, instructions.trim(), context);
      return {
        statusCode: 200,
        type: 'application/vnd.microsoft.card.adaptive',
        value: processingCard(),
      } satisfies AdaptiveCardActionCardResponse;
    },

    async publish({ activity }: PublishContext) {
      if (activity.value?.trigger !== 'manual') return actionMessage('Select Publish to share this draft.');
      const owned = getOwnedDraft(activity);
      if (!owned) return actionMessage('This draft is unavailable. Request a new draft.');
      const { draft, draftId } = owned;
      if (draft.status === 'revising') return actionMessage('Wait for the draft revision to finish.');
      if (draft.status === 'publishing') return actionMessage('This draft is being published.');
      if (draft.status === 'failed') return actionMessage('Publication could not be confirmed. Check the thread before requesting a new draft.');
      if (draft.status === 'pending') {
        draft.status = 'publishing';
        try {
          const sent = await app.send(draft.conversationId, new MessageActivityInput()
            .addCard('adaptive', draftCard(draft.text, draftId, true)));
          draft.publishedMessageId = sent.id;
          draft.status = 'published';
        } catch {
          draft.status = 'failed';
          return actionMessage('Publication could not be confirmed. Check the thread before requesting a new draft.');
        }
      }
      return {
        statusCode: 200,
        type: 'application/vnd.microsoft.card.adaptive',
        value: draftCard(draft.text, draftId, true),
      } satisfies AdaptiveCardActionCardResponse;
    },
  };
}