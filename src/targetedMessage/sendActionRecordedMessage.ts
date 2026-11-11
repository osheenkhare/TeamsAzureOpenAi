import { MessageActivityInput, type Activity } from '@microsoft/teams.api';
import type { IActivityContext } from '@microsoft/teams.apps';

// Deferred so the invoke response (card update) goes out before the confirmation.
export function sendActionRecordedMessage({ activity, send, log }: Pick<IActivityContext<Activity>, 'activity' | 'send' | 'log'>, taskText: string): void {
  const text = `You Marked ${taskText} Done on ${new Date().toUTCString()}`;
  setImmediate(() => {
    send(new MessageActivityInput(text).withRecipient(activity.from, true))
      .catch((error) => log.error('Failed to send targeted confirmation', error));
  });
}
