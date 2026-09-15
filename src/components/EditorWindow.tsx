import { useEffect, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

interface Props {
  title: string;
  onClose: () => void;
  onBlocked: () => void;
  children: ReactNode;
}

/**
 * Renders children into a separate browser window.
 *
 * A portal is used rather than a second app instance so both windows share one React tree and one
 * copy of the data. The alternative — loading the app twice and syncing through storage events —
 * would mean two sources of truth for the same catalogue, which is the failure mode this whole
 * project exists to remove.
 *
 * Stylesheets have to be copied across because a new window starts with an empty document.
 */
export function EditorWindow({ title, onClose, onBlocked, children }: Props) {
  const [container, setContainer] = useState<HTMLElement | null>(null);

  useEffect(() => {
    const win = window.open(
      '',
      'aplus-editor',
      `width=${Math.min(1700, Math.round(screen.availWidth * 0.95))},height=${Math.min(
        1000,
        Math.round(screen.availHeight * 0.9),
      )},resizable=yes,scrollbars=yes`,
    );

    if (!win) {
      onBlocked();
      return;
    }

    win.document.title = title;
    win.document.body.innerHTML = '';

    for (const node of Array.from(document.querySelectorAll('style, link[rel="stylesheet"]'))) {
      win.document.head.appendChild(node.cloneNode(true));
    }

    const meta = win.document.createElement('meta');
    meta.name = 'viewport';
    meta.content = 'width=device-width, initial-scale=1';
    win.document.head.appendChild(meta);

    const mount = win.document.createElement('div');
    mount.className = 'editor-window';
    win.document.body.appendChild(mount);
    setContainer(mount);

    // Closing the window from its own title bar has to reach the parent, or the parent keeps
    // believing the editor is still open and the button never comes back.
    const handleUnload = () => onClose();
    win.addEventListener('beforeunload', handleUnload);

    // If the parent is closed or reloaded first, take the child with it rather than leaving an
    // orphan window whose React tree no longer exists.
    const closeChild = () => win.close();
    window.addEventListener('beforeunload', closeChild);

    return () => {
      win.removeEventListener('beforeunload', handleUnload);
      window.removeEventListener('beforeunload', closeChild);
      win.close();
    };
    // Opening the window is a one-time effect; re-running it would spawn duplicates.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (!container) return null;
  return createPortal(children, container);
}
