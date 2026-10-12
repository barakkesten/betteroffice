import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { act, render, waitFor } from '@testing-library/react';
import { open, story } from './worker/peerlessTestHarness';
import * as wasm from '../../pptx/src/wasm/loader';
import { workerEditorSessionOpener } from './worker/useEditableSessionPresentation';
import { peerlessSessionOpener } from './worker/usePeerlessPresentation';
import { PptxEditor } from './PptxEditor';
import type { PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';
import { resetPptxEngineWarningsForTests } from './engineWarnings';

test('default readiness uses one worker engine and no eager main-thread engine', async () => {
  const hybrid = spyOn(workerEditorSessionOpener, 'open');
  const engines = spyOn(wasm, 'openPresentation');
  const warn = spyOn(console, 'warn');
  try {
    const { api, view } = await open();
    expect(view.getByTestId('pptx-peerless-editor')).toBeTruthy();
    expect(hybrid).not.toHaveBeenCalled();
    expect(engines).toHaveBeenCalledTimes(1);
    expect(warn).not.toHaveBeenCalled();
    expect(() => Reflect.get(api, 'handle')).toThrow('handleAsync');
    expect(() => Reflect.get(api, 'save')()).toThrow('saveAsync');
    expect(await api.saveAsync()).toBeInstanceOf(Uint8Array);
    expect((await story(api)).paragraphs).not.toHaveLength(0);
  } finally { hybrid.mockRestore(); engines.mockRestore(); warn.mockRestore(); }
});

test('initialSlide clamps on open and later changes preserve edits and the session', async () => {
  const opener = spyOn(peerlessSessionOpener, 'open');
  try {
    const options = { initialSlide: Number.MAX_SAFE_INTEGER };
    const { api, client, view, ui } = await open(options);
    const starts = opener.mock.calls.length;
    expect(client.state.frame?.slideIndex).toBe(1);
    const before = await story(api);
    const result = await api.applyEdits({ expectVersion: await api.version(), steps: [{
      op: 'insertText', target: { kind: 'range', slideId: before.slideId, shapeId: before.shapeId, storyId: before.storyId, start: 0, end: 0 }, at: 'start', text: 'kept ',
    }] });
    expect(result.ok).toBe(true);
    const version = await api.version();
    options.initialSlide = 1;
    await act(async () => { view.rerender(ui()); });
    expect(opener).toHaveBeenCalledTimes(starts);
    expect(client.state.frame?.slideIndex).toBe(1);
    expect(await api.version()).toBe(version);
    expect(JSON.stringify(await story(api))).toContain('kept ');
  } finally { opener.mockRestore(); }
});

for (const capability of ['Worker', 'OffscreenCanvas', 'experimentalWorkerOpen'] as const) {
  test(`${capability === 'experimentalWorkerOpen' ? 'disabled' : 'missing'} ${capability} uses the in-thread async API and warns once`, async () => {
    const { view, bytes } = await open();
    view.unmount();
    resetPptxEngineWarningsForTests();
    const descriptor = Object.getOwnPropertyDescriptor(globalThis, capability);
    if (capability !== 'experimentalWorkerOpen') Object.defineProperty(globalThis, capability, { configurable: true, value: undefined });
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    const opener = spyOn(peerlessSessionOpener, 'open');
    opener.mockClear();
    const hybrid = spyOn(workerEditorSessionOpener, 'open');
    let api: PptxAsyncEditorApi | undefined;
    try {
      const font = await readFile(new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
      const props = { fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }], file: bytes,
        experimentalWorkerOpen: capability === 'experimentalWorkerOpen' ? false : undefined,
        onReady: (value: PptxAsyncEditorApi) => { api = value; } };
      const fallback = render(<PptxEditor {...props} />);
      await waitFor(() => expect(api).toBeDefined());
      expect(fallback.queryByTestId('pptx-peerless-editor')).toBeNull();
      expect(opener).not.toHaveBeenCalled();
      expect(hybrid).not.toHaveBeenCalled();
      expect('handle' in api!).toBe(false);
      expect('save' in api!).toBe(false);
      expect(await api!.saveAsync()).toBeInstanceOf(Uint8Array);
      expect(await api!.goToSlide(2)).toBe(true);
      await act(async () => { fallback.rerender(<PptxEditor {...props} readOnly />); });
      const access = await api!.handleAsync();
      await expect(access.deleteSlide('slide')).rejects.toMatchObject({ code: 'read-only' });
      fallback.unmount();
      api = undefined;
      render(<PptxEditor {...props} />);
      await waitFor(() => expect(api).toBeDefined());
      expect(warn).toHaveBeenCalledTimes(1);
      expect(warn.mock.calls[0]?.[0]).toContain(capability);
      expect(warn.mock.calls[0]?.[0]).toContain('in-thread');
    } finally {
      if (descriptor) Object.defineProperty(globalThis, capability, descriptor);
      else Reflect.deleteProperty(globalThis, capability);
      warn.mockRestore(); opener.mockRestore(); hybrid.mockRestore(); resetPptxEngineWarningsForTests();
    }
  });
}
