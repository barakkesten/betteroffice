import { expect, mock, test } from 'bun:test';
import { copyWorkerSelection } from './clipboard';
import type { PeerlessEditSessionBinding } from '../hooks/usePeerlessEditSession';

function fixture(readOnly = false) {
  const selection = { anchor: { story: 'body', paraId: 'a', offset: 1 }, focus: { story: 'body', paraId: 'a', offset: 4 } };
  const state = { workerSelection: selection, editState: { readOnly }, version: 1 };
  const submit = mock((_intent: unknown) => ({ status: 'submitted', ticket: { ack: Promise.resolve({ status: 'applied' }) } }));
  const binding = { viewStore: { getSnapshot: () => state }, commands: { submit }, session: { flush: async () => {} } } as unknown as PeerlessEditSessionBinding;
  const read = mock(async () => ({ ok: true as const, version: 'v1', view: 'accepted' as const,
    paragraphs: [{ story: 'body', paraId: 'a', text: 'abcdef', atoms: [] }] }));
  return { binding, submit, read, state, selection };
}

test('cut deletes the captured range after clipboard completion even if the caret moved', async () => {
  const { binding, submit, read, state, selection } = fixture();
  const write = mock(async (text: string) => {
    expect(text).toBe('bcd');
    state.workerSelection = { anchor: { ...selection.anchor, offset: 5 }, focus: { ...selection.focus, offset: 5 } };
  });
  expect(await copyWorkerSelection({ binding, read, write, cut: true, isCurrent: () => true })).toBe(true);
  expect(submit.mock.calls[0]?.[0]).toEqual({ kind: 'replaceSelection', payload: { text: '', selection } });
});

test('failed clipboard writes and read-only cuts preserve the document', async () => {
  const { binding, submit, read } = fixture();
  await expect(copyWorkerSelection({ binding, read, cut: true, isCurrent: () => true,
    write: async () => { throw new Error('clipboard unavailable'); } })).rejects.toThrow('clipboard unavailable');
  expect(submit).not.toHaveBeenCalled();
  const locked = fixture(true);
  expect(await copyWorkerSelection({ ...locked, write: async () => {}, cut: true, isCurrent: () => true })).toBe(false);
  expect(locked.submit).not.toHaveBeenCalled();
});

test('replacement while copying cannot cut the new document', async () => {
  const { binding, submit, read } = fixture();
  let current = true;
  expect(await copyWorkerSelection({ binding, read, cut: true, isCurrent: () => current,
    write: async () => { current = false; } })).toBe(false);
  expect(submit).not.toHaveBeenCalled();
});
