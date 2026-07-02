import { Dialog as DialogPrimitive } from "@base-ui/react/dialog";

import { cn } from "../../lib/utils";
import type {
  FolioDialogBackdropProps,
  FolioDialogCloseProps,
  FolioDialogPopupProps,
  FolioDialogPortalProps,
  FolioDialogRootProps,
  FolioDialogTitleProps,
} from "../folio-ui";

/**
 * Built-in, dependency-light Dialog parts used when a consumer does not inject
 * its own. Each part wraps the matching `@base-ui/react` primitive — the same
 * accessible primitive the app design system wraps — with minimal styling.
 * Folio's chrome positions the popup itself via `className`, so the default
 * popup is a bare portalled surface (no viewport grid); consumers inject a
 * polished Dialog through `DocxEditor`'s `components` prop.
 */

function DefaultDialogRoot(props: FolioDialogRootProps) {
  return <DialogPrimitive.Root {...props} />;
}

function DefaultDialogPortal({ children, ...props }: FolioDialogPortalProps) {
  // `folio-root` re-establishes the editor root inside the body portal so design
  // tokens and the standalone stylesheet's scoped utilities apply to the dialog
  // (and `.dark .folio-root` themes it). `display: contents` keeps the wrapper
  // out of layout, so the fixed-position backdrop/popup are unaffected.
  return (
    <DialogPrimitive.Portal {...props}>
      <div className="folio-root" style={{ display: "contents" }}>
        {children}
      </div>
    </DialogPrimitive.Portal>
  );
}

function DefaultDialogBackdrop({ className, ...props }: FolioDialogBackdropProps) {
  return (
    <DialogPrimitive.Backdrop
      className={cn("folio-default-dialog-backdrop", className)}
      {...props}
    />
  );
}

function DefaultDialogPopup({ className, ...props }: FolioDialogPopupProps) {
  return (
    <DialogPrimitive.Popup className={cn("folio-default-dialog-popup", className)} {...props} />
  );
}

function DefaultDialogTitle({ className, ...props }: FolioDialogTitleProps) {
  return (
    <DialogPrimitive.Title className={cn("folio-default-dialog-title", className)} {...props} />
  );
}

function DefaultDialogClose({ className, ...props }: FolioDialogCloseProps) {
  return (
    <DialogPrimitive.Close className={cn("folio-default-dialog-close", className)} {...props} />
  );
}

export const DefaultDialog = {
  Root: DefaultDialogRoot,
  Portal: DefaultDialogPortal,
  Backdrop: DefaultDialogBackdrop,
  Popup: DefaultDialogPopup,
  Title: DefaultDialogTitle,
  Close: DefaultDialogClose,
};
