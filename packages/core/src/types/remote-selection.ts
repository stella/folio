/** Serializable collaborator selection projected into ProseMirror positions. */
export type RemoteSelection = {
  anchor: number;
  clientId: number;
  color: string;
  head: number;
  name: string;
};
