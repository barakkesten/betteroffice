import { faultControl, open, story, type } from './peerlessTestHarness';
import { expect, mock, test } from 'bun:test';
import type { DeckSnapshot } from '@betteroffice/pptx';
const { act, waitFor } = await import('@testing-library/react');

test('unknown_outcome', async () => {
  const error = mock(() => {});
  const changed = mock((_snapshot: DeckSnapshot) => {});
  const value = await open({ onError: error, onChange: changed });
  const target = await story(value.api);
  await act(async () => { await value.api.selectText({ slide: 1, shapeId: target.shapeId, storyId: target.storyId, start: 0, end: 0 }); });
  const input = value.view.getByLabelText('Slide text');
  await act(async () => { type(input, 'Covered'); await value.api.flushPendingInput(); });
  await waitFor(() => expect(changed).toHaveBeenCalled());
  const checkpoint = await value.api.recoverySave();
  faultControl.unknown();
  await act(async () => { type(input, 'Uncertain'); await value.api.flushPendingInput().catch(() => {}); });
  await waitFor(() => expect(error).toHaveBeenCalledTimes(1));
  const recovered = await value.api.recoverySave();
  expect(recovered.bytes).toEqual(checkpoint.bytes);
  expect(value.view.getByRole('button', { name: 'Download recovered presentation' })).toBeDefined();
  expect(value.stage.getAttribute('aria-readonly')).toBe('true');
  await expect(value.api.saveAsync()).rejects.toThrow();
  value.view.unmount();
  await expect(value.api.recoverySave()).rejects.toThrow();
});
