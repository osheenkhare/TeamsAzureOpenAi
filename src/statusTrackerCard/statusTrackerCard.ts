import type { AdaptiveCardActionCardResponse, IAdaptiveCardActionInvokeActivity } from '@microsoft/teams.api';
import type { IActivityContext } from '@microsoft/teams.apps';
import { sendActionRecordedMessage } from '../targatedMessage/sendActionRecordedMessage.ts';

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

export function handleMarkTaskDone({ activity, send, log }: IActivityContext<IAdaptiveCardActionInvokeActivity>): AdaptiveCardActionCardResponse {
  const { data } = activity.value.action;
  const tasks = readTasks(data);
  const task = tasks[Number(data.index)];
  if (task) task.done = true;

  // Comment out to skip the targeted confirmation while testing.
  sendActionRecordedMessage({ activity, send, log });

  return { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive', value: createStatusTrackerCard(tasks) };
}
