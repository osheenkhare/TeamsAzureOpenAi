import { MessageActivityInput, type AdaptiveCardActionCardResponse, type IAdaptiveCardActionInvokeActivity } from '@microsoft/teams.api';
import type { IActivityContext } from '@microsoft/teams.apps';
import { sendActionRecordedMessage } from '../targetedMessage/sendActionRecordedMessage.ts';

type AdaptiveCard = AdaptiveCardActionCardResponse['value'];

export const MARK_TASK_DONE_ACTION = 'markTaskDone';

export interface ITrackerTask {
  text: string;
  done: boolean;
}

// The card is stateless on the server, so the task list rides along in every action's data.
export function createStatusTrackerCard(tasks: ITrackerTask[]): AdaptiveCard {
  const doneCount = tasks.filter((task) => task.done).length;

  return {
    type: 'AdaptiveCard',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: 'Status Tracker', size: 'Large', weight: 'Bolder', wrap: true },
      { type: 'TextBlock', text: `${doneCount} of ${tasks.length} done`, isSubtle: true, spacing: 'None', wrap: true },
      ...tasks.map((task, index) => ({
        type: 'ColumnSet' as const,
        separator: true,
        columns: [
          {
            type: 'Column' as const,
            width: 'stretch',
            verticalContentAlignment: 'Center' as const,
            items: [{
              type: 'TextBlock' as const,
              text: task.done ? `✓ ${task.text}` : task.text,
              color: task.done ? 'Good' as const : 'Default' as const,
              isSubtle: task.done,
              wrap: true,
            }],
          },
          {
            type: 'Column' as const,
            width: 'auto',
            verticalContentAlignment: 'Center' as const,
            items: task.done
              ? [{ type: 'TextBlock' as const, text: 'Done', color: 'Good' as const, weight: 'Bolder' as const }]
              : [{
                  type: 'ActionSet' as const,
                  actions: [{
                    type: 'Action.Execute' as const,
                    title: 'Mark Done',
                    data: { action: MARK_TASK_DONE_ACTION, index, tasks },
                  }],
                }],
          },
        ],
      })),
    ],
  };
}

function readTasks(data: Record<string, unknown>): ITrackerTask[] {
  if (!Array.isArray(data.tasks)) return [];
  return data.tasks.map((task: Partial<ITrackerTask>) => ({ text: String(task?.text ?? ''), done: task?.done === true }));
}

export async function handleMarkTaskDone({ activity, api, send, log }: IActivityContext<IAdaptiveCardActionInvokeActivity>): Promise<AdaptiveCardActionCardResponse> {
  // replyToId identifies the original bot message, not this invoke activity.
  const messageId = activity.replyToId;
  if (!messageId) throw new Error('Cannot update the shared tracker: missing original message ID.');

  const { data } = activity.value.action;
  const tasks = readTasks(data);
  const task = tasks[Number(data.index)];
  if (task) task.done = true;

  const card = createStatusTrackerCard(tasks);
  // Update the shared channel message, not just the clicking user's card view.
  // POC: action data can be stale, so concurrent clicks may overwrite progress.
  await api.conversations.activities(activity.conversation.id).update(
    messageId,
    new MessageActivityInput().addCard('adaptive', card),
  );

  // Comment out to skip the targeted confirmation while testing.
  // if (task) sendActionRecordedMessage({ activity, send, log }, task.text);

  return { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive', value: card };
}
