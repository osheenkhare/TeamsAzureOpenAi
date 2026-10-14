export interface IAgentMessage {
  role: 'user' | 'assistant';
  content: string;
}