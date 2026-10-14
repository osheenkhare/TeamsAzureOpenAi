import type { AdaptiveCardActionCardResponse } from '@microsoft/teams.api';

export function processingCard(): AdaptiveCardActionCardResponse['value'] {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [{ type: 'TextBlock', text: 'Agent is processing...', wrap: true }],
    actions: [],
  };
}

export function draftCard(text: string, draftId: string, published = false): AdaptiveCardActionCardResponse['value'] {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: published ? 'Published' : 'Draft response', weight: 'Lighter', wrap: true },
      { type: 'TextBlock', text, wrap: true },
      ...published ? [] : [{
        type: 'Input.Text' as const,
        id: 'revisionInstructions',
        label: 'Revision instructions',
        placeholder: 'What should change?',
        isMultiline: true,
        isRequired: true,
        errorMessage: 'Enter revision instructions.',
      }],
    ],
    actions: published ? [] : [
      { type: 'Action.Execute', title: 'Publish', verb: 'publishDraft', data: { action: 'publishDraft', draftId }, associatedInputs: 'none' },
      { type: 'Action.Execute', title: 'Revise', verb: 'reviseDraft', data: { action: 'reviseDraft', draftId }, associatedInputs: 'auto' },
    ],
  };
}