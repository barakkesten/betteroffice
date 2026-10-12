import { open, story } from './peerlessTestHarness';
import { expect, spyOn, test } from 'bun:test';
import type { PresentationHandle } from '@betteroffice/pptx';
const { act } = await import('@testing-library/react');
type ReviewAccess = { [K in 'propose' | 'previewProposal' | 'layoutProposalDiffSlide' | 'acceptProposal' | 'rejectProposal']:
  (...args: Parameters<PresentationHandle[K]>) => Promise<ReturnType<PresentationHandle[K]>> };

test('async_review_export', async () => {
  const value = await open();
  const target = await story(value.api);
  const access = await value.api.handleAsync();
  const review = access as unknown as ReviewAccess;
  let id = '';
  await act(async () => { id = (await review.propose('reviewer', 'Update', [{ type: 'replaceText', storyId: target.storyId, start: 0, end: 5, text: 'Updated' }])).id; });
  expect((await review.previewProposal(id)).proposal.id).toBe(id);
  expect((await review.layoutProposalDiffSlide(id, 0)).frame.primitives.length).toBeGreaterThan(0);
  await act(async () => { expect(await value.api.commands!.execute('proposalsPanel', { open: true })).toMatchObject({ ok: true }); });
  expect(value.view.getByTestId('pptx-proposals-panel')).toBeDefined();
  await act(async () => { expect(await value.api.commands!.execute('proposalAccept', { proposalId: id })).toMatchObject({ ok: true }); });
  expect((await story(value.api)).text.startsWith('Updated')).toBe(true);
  const markdown = await access.exportMarkdown();
  expect(markdown.ok).toBe(true);
  if (!markdown.ok) throw new Error(markdown.failure.message);
  expect(markdown.content.markdown).toContain('Updated');
  const structured = await access.exportStructured();
  expect(structured.ok).toBe(true);
  if (!structured.ok) throw new Error(structured.failure.message);
  expect(structured.content.slides).toHaveLength(2);
  const png = new Blob([new Uint8Array([137, 80, 78, 71])], { type: 'image/png' });
  const encode = spyOn(HTMLCanvasElement.prototype, 'toBlob').mockImplementation((callback) => callback(png));
  const createUrl = spyOn(URL, 'createObjectURL').mockReturnValue('blob:slide');
  const revokeUrl = spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
  const download = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
  try {
    await act(async () => { expect(await value.api.commands!.execute('exportPng', null)).toMatchObject({ ok: true }); });
    expect(encode).toHaveBeenCalledWith(expect.any(Function), 'image/png');
    expect(createUrl).toHaveBeenCalledWith(png);
    expect(download).toHaveBeenCalledTimes(1);
    expect(revokeUrl).toHaveBeenCalledWith('blob:slide');
  } finally { encode.mockRestore(); createUrl.mockRestore(); revokeUrl.mockRestore(); download.mockRestore(); }
  await act(async () => { expect(await value.api.commands!.execute('slideshow', null)).toMatchObject({ ok: true }); });
  expect(value.view.getByRole('dialog')).toBeDefined();
  await act(async () => { value.view.rerender(value.ui(true)); await value.api.flushPendingInput(); });
  await expect(review.acceptProposal(id)).rejects.toThrow();
});
