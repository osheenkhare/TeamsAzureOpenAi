export interface IChannelThreadActivity {
  id: string;
  text?: string;
  timestamp?: Date;
  from: { id: string; aadObjectId?: string; name?: string };
  conversation: { id: string };
  team?: { aadGroupId?: string };
  channel?: { id?: string };
}