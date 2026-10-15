import { App } from '@microsoft/teams.apps';
import { ConsoleLogger } from '@microsoft/teams.common';
import { createMentionHandler, createMessageHandler } from './eventHandler.ts';
import { OpenAiAgenticOrchestrator } from './azureOpenAi/azureOpenAiOrchestrator.ts';
import type { McpToolProvider } from './azureOpenAi/mcpToolProvider.ts';
import { MARK_TASK_DONE_ACTION, handleMarkTaskDone } from './statusTrackerCard/statusTrackerCard.ts';
import { TrackerOrchestrator } from './statusTrackerCard/trackerOrchestrator.ts';

function requiredEnv(name: string): string {
  const value = process.env[name];
  if (!value) return '';
  return value;
}

let mcp: McpToolProvider | undefined;
const Endpoint = requiredEnv('AZURE_OPENAI_ENDPOINT');
const ApiKey = requiredEnv('AZURE_OPENAI_API_KEY');
const Deployment = requiredEnv('AZURE_OPENAI_DEPLOYMENT');

const { McpToolProvider: Mcp } = await import('./azureOpenAi/mcpToolProvider.ts');
mcp = await Mcp.connect('http://127.0.0.1:3001/mcp');


const agent = await OpenAiAgenticOrchestrator.create({ Endpoint, ApiKey, Deployment }, mcp);

const tracker = new TrackerOrchestrator({ Endpoint, ApiKey, Deployment });

const app = new App({
  clientId: requiredEnv('CLIENT_ID'),
  clientSecret: requiredEnv('CLIENT_SECRET'),
  tenantId: requiredEnv('TENANT_ID'),
  dangerouslyAllowUnauthenticatedRequests: requiredEnv('LOCAL_DEBUG') === 'true',
  logger: new ConsoleLogger('TeamsApp'),
});

const port = requiredEnv('PORT');


app.on('mention', createMentionHandler(app, agent, tracker));
app.on('message', createMessageHandler(agent));
app.on(`card.action.${MARK_TASK_DONE_ACTION}`, handleMarkTaskDone);

await app.start(Number(port || 3978));