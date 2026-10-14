export interface IThreadMessage {
  id: string;
  replyToId: string | null;
  createdDateTime: string;
  lastModifiedDateTime: string;
  from?: {
    user?: { id?: string; displayName?: string };
    application?: { id?: string; displayName?: string };
  };
  body?: { contentType?: string; content?: string };
}