import { executeEditorCommand, executeFirstEditorCommand } from "./executeEditorCommand";
const handleFormat = (view, command, commands) => {
  executeEditorCommand(view, command);
  executeFirstEditorCommand(view, commands);
};
