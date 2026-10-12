import { expect, mock, test } from 'bun:test';
import type { DeckSnapshot, PresentationHandle } from '@betteroffice/pptx';
import { open } from './peerlessTestHarness';
import * as wasm from '../../../pptx/src/wasm/loader';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

type NotesAccess = { [K in 'setSlideNotes' | 'addComment' | 'replyToComment' | 'setCommentStatus' | 'setCommentPosition' | 'removeComment' | 'setCommentFlavor' | 'comments']:
  (...args: Parameters<PresentationHandle[K]>) => Promise<ReturnType<PresentationHandle[K]>> };

test('notes_comments_change', async () => {
  const changed = mock((_snapshot: DeckSnapshot) => {});
  const value = await open({ onChange: changed });
  const notes = value.view.getByTestId('pptx-notes-textarea');
  await act(async () => { fireEvent.change(notes, { target: { value: 'Speaker notes' } }); await value.api.flushPendingInput(); });
  const access = await value.api.handleAsync() as unknown as NotesAccess;
  expect((await (await value.api.handleAsync()).snapshot()).slides[0]!.notes).toBe('Speaker notes');
  await act(async () => {
    await access.setCommentFlavor('modern');
    const root = await access.addComment(value.client.state.activeSlideId!, { author: 'Writer', text: 'Discuss', created: '2026-01-01T00:00:00Z' });
    await access.replyToComment(root.commentId, { author: 'Reader', text: 'Agreed', created: '2026-01-01T00:00:01Z' });
    await access.setCommentPosition(root.commentId, { xEmu: 100, yEmu: 200 });
    await access.setCommentStatus(root.commentId, true);
  });
  const comments = await access.comments();
  expect(comments).toHaveLength(2);
  expect(comments.find((comment) => !comment.parentId)).toMatchObject({ resolved: true, xEmu: 100, yEmu: 200 });
  await waitFor(() => expect(changed.mock.calls[changed.mock.calls.length - 1]?.[0]).toMatchObject({ commentFlavor: 'modern', comments }));
  const saved = wasm.openPresentation(await value.api.saveAsync());
  try {
    expect(saved.snapshot().slides[0]!.notes).toBe('Speaker notes');
    const normalized = (values: typeof comments) => values.map(({ id, parentId, ...comment }) => ({
      ...comment, parent: parentId ? values.findIndex((value) => value.id === parentId) : null,
    }));
    expect(normalized(saved.comments())).toEqual(normalized(comments));
  } finally { saved.dispose(); }
  await act(async () => { value.view.rerender(value.ui(true)); await value.api.flushPendingInput(); });
  expect((value.view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement).disabled).toBe(true);
  await expect(access.setSlideNotes(value.client.state.activeSlideId!, 'Blocked')).rejects.toThrow();
});
