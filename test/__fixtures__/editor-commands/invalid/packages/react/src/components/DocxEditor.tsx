import { registerEditorCommandOwner as register } from "./executeEditorCommand";
import * as commandOwners from "./executeEditorCommand.ts";
const handleFormat = (view, action) => {
  const commandState = view.state;
  toggleBold(commandState, view.dispatch);
  setAlignment(action.value)(view.state, (tr) => view.dispatch(tr));
};
const unrelated = (view) => toggleBold(view.state, view.dispatch);
