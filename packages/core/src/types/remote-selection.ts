/** Serializable collaborator selection projected into ProseMirror positions. */
export type RemoteSelection = {
  anchor: number;
  /** The collaborator's identity: one selection per peer. */
  peer: string;
  color: string;
  head: number;
  name: string;
};
