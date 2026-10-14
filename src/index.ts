import { App } from '@microsoft/teams.apps';
import { ConsoleLogger } from '@microsoft/teams.common';
import { createMentionHandler, createMessageHandler } from './messageHandler.ts';
import { OpenAiAgenticOrchestrator } from './azureOpenAi/azureOpenAiOrchestrator.ts';
import { createDraftFlow } from './targatedMessage/draftMessage.ts';

function requiredEnv(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`Missing required environment variable: ${name}`);
  return value;
}

const agent = await OpenAiAgenticOrchestrator.create({
  Endpoint: requiredEnv('AZURE_OPENAI_ENDPOINT'),
  ApiKey: requiredEnv('AZURE_OPENAI_API_KEY'),
  Deployment: requiredEnv('AZURE_OPENAI_DEPLOYMENT'),
});

const app = new App({
  clientId: requiredEnv('CLIENT_ID'),
  clientSecret: requiredEnv('CLIENT_SECRET'),
  tenantId: requiredEnv('TENANT_ID'),
  logger: new ConsoleLogger('AzureOpenAi'),
});

const draftFlow = createDraftFlow(app, agent);
app.on('mention', createMentionHandler(app, agent, draftFlow));
app.on('card.action.publishDraft', draftFlow.publish);
app.on('card.action.reviseDraft', draftFlow.revise);
app.on('message', createMessageHandler(agent));

await app.start(Number(process.env.PORT || 3978));