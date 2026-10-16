import { MessageActivityInput, type Activity } from '@microsoft/teams.api';
import type { IActivityContext } from '@microsoft/teams.apps';

const ACTION_RECORDED_TEXT = 'Your Action has been recorded';

// Deferred so the invoke response (card update) goes out before the confirmation.
export function sendActionRecordedMessage({ activity, send, log }: Pick<IActivityContext<Activity>, 'activity' | 'send' | 'log'>): void {
  setImmediate(() => {
    send(new MessageActivityInput(ACTION_RECORDED_TEXT).withRecipient(activity.from, true))
      .catch((error) => log.error('Failed to send targeted confirmation', error));
  });
}
