import { useEffect, useRef, useSyncExternalStore, type ReactElement } from 'react';
import type { PeerlessViewStore } from './peerlessViewStore';
import type { PeerlessVisualStore } from './usePeerlessPresentation';
import { commitComposition, peerlessInputReference } from './peerlessViewStore';
import type { PresentationCommandAck, PresentationCommandIntent } from '@betteroffice/pptx';
import type { PptxCommandStore } from '../commands/types';
import { formatPeerlessText } from './peerlessTextCommands';
import { usePeerlessTextSelection } from './usePeerlessTextSelection';

export type PeerlessInputShortcut = 'undo' | 'redo' | 'copy' | 'cut' | 'save';

export interface PeerlessTextInputProps {
  store: PeerlessViewStore;
  visual?: PeerlessVisualStore;
  viewport?: { left: number; top: number; scale: number };
  commands?: PptxCommandStore;
  onShortcut?: (shortcut: PeerlessInputShortcut) => void;
}

let nextInputGroup = 0;

export function PeerlessTextInput({ store, visual = store, viewport, onShortcut, commands }: PeerlessTextInputProps): ReactElement {
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const commandsRef = useRef(commands);
  commandsRef.current = commands;
  const shortcutRef = useRef(onShortcut);
  shortcutRef.current = onShortcut;
  const view = useSyncExternalStore(visual.subscribePresentation, visual.getSnapshot, visual.getSnapshot);
  const { move } = usePeerlessTextSelection(store);

  useEffect(() => {
    const input = inputRef.current;
    if (!input) return;
    let trailingComposition: string | null = null;
    let pastedText: string | null = null;
    let inputOperation: Promise<PresentationCommandAck> | undefined;
    let cancelledComposition = false;
    let composing = false;

    function dispatch(intent: PresentationCommandIntent): void {
      const operation = store.dispatch(intent);
      if (operation !== inputOperation) void operation.catch(() => {});
      inputOperation = operation;
    }

    function insert(text: string, time: number, paste = false): void {
      const at = peerlessInputReference(store, true);
      if (store.getSnapshot().selection?.kind === 'object' && at?.source !== 'afterSeq') return;
      if (!at || !text) return;
      const payload = { at, text, inputTime: time,
        undoGroup: paste ? `paste-${++nextInputGroup}` : 'typing' };
      dispatch({ kind: paste ? 'replaceText' : 'insertText', payload });
    }

    function clearInput(): void {
      if (!store.getSnapshot().composition) input!.value = '';
    }

    function beforeInput(event: InputEvent): void {
      const current = store.getSnapshot();
      const at = peerlessInputReference(store, true);
      if (current.editState.readOnly || current.selection?.kind === 'object' && at?.source !== 'afterSeq') {
        event.preventDefault(); return;
      }
      if (cancelledComposition || composing && !current.composition) {
        event.preventDefault();
        clearInput();
        return;
      }
      if (current.composition || event.isComposing || event.inputType === 'insertCompositionText') {
        if (event.data !== null && current.composition) store.updateComposition(event.data);
        return;
      }
      if (trailingComposition !== null) {
        const trailing = trailingComposition;
        trailingComposition = null;
        if ((event.inputType === 'insertText' || event.inputType === 'insertFromComposition') && event.data === trailing) {
          event.preventDefault();
          clearInput();
          return;
        }
      }
      if (!at) return;
      const payload = { at, inputTime: event.timeStamp, undoGroup: 'typing' };
      switch (event.inputType) {
        case 'insertText':
        case 'insertReplacementText':
        case 'insertFromComposition':
          event.preventDefault();
          insert(event.data ?? '', event.timeStamp);
          break;
        case 'insertFromPaste': {
          event.preventDefault();
          const text = event.dataTransfer?.getData('text/plain') ?? event.data ?? '';
          if (pastedText !== text) insert(text, event.timeStamp, true);
          pastedText = null;
          break;
        }
        case 'deleteContentBackward':
          event.preventDefault();
          dispatch({ kind: 'deleteBackward', payload: { ...payload, unit: 'codePoint' } });
          break;
        case 'deleteContentForward':
          event.preventDefault();
          dispatch({ kind: 'deleteForward', payload: { ...payload, unit: 'codePoint' } });
          break;
        case 'insertParagraph':
        case 'insertLineBreak':
          event.preventDefault();
          dispatch({ kind: 'paragraphBreak', payload });
          break;
        default:
          return;
      }
      clearInput();
    }

    function keyDown(event: KeyboardEvent): void {
      if (event.defaultPrevented) return;
      const modifier = event.ctrlKey || event.metaKey;
      const key = event.key.toLowerCase();
      if (modifier && !event.altKey && (key === 'z' || key === 'y' && !event.shiftKey)) {
        event.preventDefault();
        if (event.repeat || store.getSnapshot().editState.readOnly) return;
        cancelledComposition = Boolean(store.getSnapshot().composition);
        trailingComposition = null;
        pastedText = null;
        input!.value = '';
        shortcutRef.current?.(key === 'z' && !event.shiftKey ? 'undo' : 'redo');
        return;
      }
      if (event.isComposing || store.getSnapshot().composition || event.keyCode === 229) return;
      trailingComposition = null;
      pastedText = null;
      if (modifier && !event.altKey && !event.shiftKey && ['b', 'i', 'u', 's'].includes(key)) {
        event.preventDefault();
        if (event.repeat) return;
        if (key === 's') shortcutRef.current?.('save');
        else if (commandsRef.current) void commandsRef.current.execute(key === 'b' ? 'bold' : key === 'i' ? 'italic' : 'underline', null);
        else void formatPeerlessText(store, key === 'b' ? 'bold' : key === 'i' ? 'italic' : 'underline',
          event.timeStamp)?.catch(() => {});
        return;
      }
      if (modifier) {
        const shortcut = key === 'c' ? 'copy' : key === 'x' ? 'cut' : null;
        if (shortcut) {
          event.preventDefault();
          if (shortcut !== 'copy' && store.getSnapshot().editState.readOnly) return;
          shortcutRef.current?.(shortcut);
          return;
        }
      }
      if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) {
        move(event);
      }
    }

    function compositionStart(event: CompositionEvent): void {
      if (store.getSnapshot().editState.readOnly) return;
      trailingComposition = null;
      cancelledComposition = false;
      composing = true;
      store.beginComposition(`composition-${++nextInputGroup}`, event.timeStamp);
    }

    function compositionUpdate(event: CompositionEvent): void {
      if (!store.getSnapshot().editState.readOnly) store.updateComposition(event.data);
    }

    function compositionEnd(event: CompositionEvent): void {
      composing = false;
      if (cancelledComposition || !store.getSnapshot().composition) {
        trailingComposition = event.data;
        cancelledComposition = false;
        clearInput();
        return;
      }
      if (store.getSnapshot().editState.readOnly || !store.getSnapshot().composition) return;
      trailingComposition = event.data;
      void commitComposition(store, event.data, event.timeStamp).catch(() => {});
      clearInput();
    }

    function inputEvent(event: Event): void {
      if (store.getSnapshot().editState.readOnly) return;
      if (store.getSnapshot().composition) {
        const data = (event as InputEvent).data;
        store.updateComposition(data ?? input!.value);
      } else clearInput();
    }

    function paste(event: ClipboardEvent): void {
      if (store.getSnapshot().composition || store.getSnapshot().editState.readOnly) return;
      event.preventDefault();
      const text = event.clipboardData?.getData('text/plain') ?? '';
      pastedText = text;
      insert(text, event.timeStamp, true);
      clearInput();
    }

    function focus(): void { store.setFocused(true); }
    function blur(): void { store.setFocused(false); }
    input.addEventListener('beforeinput', beforeInput);
    input.addEventListener('keydown', keyDown);
    input.addEventListener('compositionstart', compositionStart);
    input.addEventListener('compositionupdate', compositionUpdate);
    input.addEventListener('compositionend', compositionEnd);
    input.addEventListener('input', inputEvent);
    input.addEventListener('paste', paste);
    input.addEventListener('focus', focus);
    input.addEventListener('blur', blur);
    return () => {
      input.removeEventListener('beforeinput', beforeInput);
      input.removeEventListener('keydown', keyDown);
      input.removeEventListener('compositionstart', compositionStart);
      input.removeEventListener('compositionupdate', compositionUpdate);
      input.removeEventListener('compositionend', compositionEnd);
      input.removeEventListener('input', inputEvent);
      input.removeEventListener('paste', paste);
      input.removeEventListener('focus', focus);
      input.removeEventListener('blur', blur);
    };
  }, [store, move]);

  const scale = viewport?.scale ?? 1;
  return <><textarea ref={inputRef} aria-label="Slide text" autoComplete="off" autoCorrect="off"
    autoCapitalize="off" spellCheck={false} tabIndex={0} readOnly={view.editState.readOnly}
    style={{ position: 'absolute', left: (viewport?.left ?? 0) + (view.caret?.x ?? 0) * scale,
      top: (viewport?.top ?? 0) + (view.caret?.y ?? 0) * scale, width: 1,
      height: Math.max(1, (view.caret?.height ?? 1) * scale), opacity: 0,
      padding: 0, border: 0, resize: 'none', overflow: 'hidden', pointerEvents: 'none' }} />
    {view.inputFailure && <div role="alert">{view.inputFailure.message}</div>}
  </>;
}
