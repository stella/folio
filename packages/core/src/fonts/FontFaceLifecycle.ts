import { panic } from "better-result";

import { Subscribable } from "../managers/Subscribable";

type FontFaceLifecycleState<Face> =
  | { type: "active"; generation: number; faces: readonly Face[] }
  | { type: "disposed" };

export class FontFaceLifecycle<Face> extends Subscribable<number> {
  private readonly remove: (faces: readonly Face[]) => void;
  private state: FontFaceLifecycleState<Face> = {
    type: "active",
    generation: 0,
    faces: [],
  };

  constructor(remove: (faces: readonly Face[]) => void) {
    super(0);
    this.remove = remove;
  }

  beginReplacement(): (faces: readonly Face[]) => void {
    const previous = this.requireActive();
    const generation = previous.generation + 1;
    this.state = { type: "active", generation, faces: previous.faces };
    return (faces) => {
      const current = this.state;
      if (current.type === "disposed" || current.generation !== generation) {
        const retained = new Set(
          current.type === "active" ? current.faces : [],
        );
        this.remove(faces.filter((face) => !retained.has(face)));
        return;
      }
      const retained = new Set(faces);
      this.state = { type: "active", generation, faces: [...retained] };
      this.remove(current.faces.filter((face) => !retained.has(face)));
      if (current.faces.length > 0 || retained.size > 0) {
        this.setSnapshot(this.getSnapshot() + 1);
      }
    };
  }

  clear(): void {
    this.beginReplacement()([]);
  }

  destroy(): void {
    const current = this.state;
    if (current.type === "disposed") return;
    this.state = { type: "disposed" };
    this.remove(current.faces);
  }

  private requireActive() {
    const current = this.state;
    if (current.type === "disposed") {
      panic(
        "Cannot replace fonts after the font lifecycle has been destroyed.",
      );
    }
    return current;
  }
}
