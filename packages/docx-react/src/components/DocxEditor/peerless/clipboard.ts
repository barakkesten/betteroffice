import type { DocxParagraphIdentitySnapshot, DocxReadParagraphsRequest, DocxReadParagraphsResult } from '@betteroffice/docx/yrs';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';

interface SelectionRead {
  binding: PeerlessEditSessionBinding;
  read(request: DocxReadParagraphsRequest): Promise<DocxReadParagraphsResult>;
  identities?(): Promise<DocxParagraphIdentitySnapshot>;
  isCurrent(): boolean;
}

export async function readWorkerSelection(options: SelectionRead) {
  const { binding } = options;
  await binding.admission?.whenCompositionSettled?.();
  await binding.session.flush();
  if (!options.isCurrent()) throw new Error('Document was replaced');
  const state = binding.viewStore.getSnapshot();
  if (!state.workerSelection) return null;
  const selection = structuredClone(state.workerSelection);
  const { anchor, focus } = selection;
  if (anchor.story !== focus.story) throw new Error('Cross-story selection is unavailable');
  let paraIds = [anchor.paraId];
  if (anchor.paraId !== focus.paraId) {
    if (!options.identities) throw new Error('Worker paragraph identities are unavailable');
    const identities = await options.identities();
    const ids = identities.paragraphs.flatMap(({ session }) => session?.story === anchor.story ? [session.paraId] : []);
    const first = ids.indexOf(anchor.paraId);
    const last = ids.indexOf(focus.paraId);
    if (first < 0 || last < 0) throw new Error('Selection target is unavailable');
    paraIds = ids.slice(Math.min(first, last), Math.max(first, last) + 1);
  }
  const result = await options.read({ story: anchor.story, paraIds, view: 'accepted' });
  if (!options.isCurrent()) throw new Error('Document was replaced');
  if (state.version !== binding.viewStore.getSnapshot().version) throw new Error('Selection read was superseded');
  if (!result.ok) throw new Error('Selection text is unavailable');
  const forward = anchor.paraId === focus.paraId ? anchor.offset <= focus.offset : paraIds[0] === anchor.paraId;
  const start = forward ? anchor : focus;
  const end = forward ? focus : anchor;
  const texts = result.paragraphs.map((paragraph) => paragraph.text.slice(
    paragraph.paraId === start.paraId ? start.offset : 0,
    paragraph.paraId === end.paraId ? end.offset : undefined,
  ));
  const paragraphText = result.paragraphs.find(({ paraId }) => paraId === focus.paraId)?.text ?? '';
  return { selection, text: texts.join('\n'), version: state.version, info: {
    paraId: focus.paraId, selectedText: texts.join('\n'), paragraphText,
    before: paragraphText.slice(0, Math.min(anchor.offset, focus.offset)),
    after: paragraphText.slice(Math.max(anchor.offset, focus.offset)),
  } };
}

export async function copyWorkerSelection(options: SelectionRead & { write(text: string): Promise<void>; cut?: boolean }) {
  const { binding } = options;
  if (options.cut && binding.viewStore.getSnapshot().editState?.readOnly) return false;
  const captured = await readWorkerSelection(options);
  if (!captured?.text) return false;
  await options.write(captured.text);
  if (!options.isCurrent()) return false;
  if (!options.cut) return true;
  if (binding.viewStore.getSnapshot().editState?.readOnly) return false;
  if (binding.viewStore.getSnapshot().version !== captured.version &&
    (!captured.selection.anchor.sticky || !captured.selection.focus.sticky)) return false;
  const result = binding.commands.submit({ kind: 'replaceSelection', payload: { text: '', selection: captured.selection } },
    { source: 'paste', fence: true });
  if (result.status !== 'submitted') return false;
  const outcome = await result.ticket.ack;
  return outcome.status === 'applied' && outcome.committed !== false;
}
