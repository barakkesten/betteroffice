import './peerlessTestCapabilities';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import type { PeerlessSession, PresentationCommandClient } from '@betteroffice/pptx';
import { StrictMode, useState } from 'react';
import { PEERLESS_FLAG, PptxEditor } from '../PptxEditor';
import type { PptxAsyncEditorProps, PptxEditorApi } from '../PptxEditor';
import type { PptxAsyncEditorApi } from './createPeerlessEditorApi';
import type { PptxWorkerEditorApi } from './createWorkerEditorApi';
import { PptxAsyncEditorUnavailableError } from '../index';
import { peerlessSessionOpener } from './usePeerlessPresentation';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
const file = new Uint8Array([1]);

class PendingWorker extends EventTarget {
  postMessage() {}
  terminate() {}
}

function setWorker(value: unknown): void {
  const original = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
  Object.defineProperty(globalThis, 'Worker', { configurable: true, writable: true, value });
  restorers.push(() => {
    if (original) Object.defineProperty(globalThis, 'Worker', original);
    else Reflect.deleteProperty(globalThis, 'Worker');
  });
}

beforeEach(() => { setWorker(PendingWorker); });
afterEach(() => {
  cleanup();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function pendingSession() {
  const owner: PeerlessSession = {
    ready: new Promise<PresentationCommandClient>(() => {}),
    failure: undefined,
    subscribeFailure: () => () => {},
    dispose: mock(() => Promise.resolve()),
  };
  const opener = spyOn(peerlessSessionOpener, 'open').mockReturnValue(owner);
  restorers.push(() => opener.mockRestore());
  return { owner, opener };
}

test('async opt-in preserves contextual editor API types', () => {
  const check = (flag: boolean) => [
    <PptxEditor fonts={[]} {...{ [PEERLESS_FLAG]: true as const }} onReady={(api) => {
      const value: PptxAsyncEditorApi = api;
      void value;
      // @ts-expect-error
      api.handle;
      // @ts-expect-error
      api.save();
    }} />,
    <PptxEditor fonts={[]} onReady={(api) => {
      const value: PptxEditorApi = api;
      void value;
    }} />,
    <PptxEditor fonts={[]} experimentalWorkerOpen onReady={(api) => {
      const value: PptxWorkerEditorApi = api;
      void value;
    }} />,
    // @ts-expect-error
    <PptxEditor fonts={[]} {...{ [PEERLESS_FLAG]: flag }} />,
    // @ts-expect-error
    <PptxEditor fonts={[]} experimentalWorkerOpen readOnly {...{ [PEERLESS_FLAG]: flag }} />,
    // @ts-expect-error
    <PptxEditor fonts={[]} {...{ [PEERLESS_FLAG]: true as const }} experimentalWorkerOpen={false} />,
  ];
  void check;
});

test('async API flag mounts the peer-less surface', async () => {
  const { opener } = pendingSession();
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }} />);
  await waitFor(() => expect(view.getByTestId('pptx-peerless-editor')).toBeTruthy());
  expect(opener).toHaveBeenCalledTimes(1);
});

test('async API without Worker refuses once without becoming ready', async () => {
  setWorker(undefined);
  const { opener } = pendingSession();
  const onError = mock((_error: Error) => {});
  const onReady = mock(() => {});
  const props = { fonts: [], file, [PEERLESS_FLAG]: true as const, onError, onReady };
  const view = render(<StrictMode><PptxEditor {...props} /></StrictMode>);
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  expect(onError.mock.calls[0][0]).toBeInstanceOf(PptxAsyncEditorUnavailableError);
  expect(view.getByRole('alert').textContent).toContain('Worker');
  view.rerender(<StrictMode><PptxEditor {...props} /></StrictMode>);
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onReady).not.toHaveBeenCalled();
  expect(opener).not.toHaveBeenCalled();
  expect(view.container.querySelector('textarea')).toBeNull();
  expect(view.container.querySelector('[role="application"]')).toBeNull();
});

test('async API without OffscreenCanvas refuses once without opening a session', async () => {
  const original = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
  Object.defineProperty(globalThis, 'OffscreenCanvas', { configurable: true, value: undefined });
  restorers.push(() => {
    if (original) Object.defineProperty(globalThis, 'OffscreenCanvas', original);
    else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
  });
  const { opener } = pendingSession();
  const onError = mock((_error: Error) => {});
  const onReady = mock(() => {});
  const props = { fonts: [], file, [PEERLESS_FLAG]: true as const, onError, onReady };
  const view = render(<StrictMode><PptxEditor {...props} /></StrictMode>);
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  expect(onError.mock.calls[0][0]).toBeInstanceOf(PptxAsyncEditorUnavailableError);
  expect(view.getByRole('alert').textContent).toContain('OffscreenCanvas');
  view.rerender(<StrictMode><PptxEditor {...props} /></StrictMode>);
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onReady).not.toHaveBeenCalled();
  expect(opener).not.toHaveBeenCalled();
  expect(view.container.querySelector('textarea')).toBeNull();
});

test('async API with worker opening disabled refuses once without becoming ready', async () => {
  const { opener } = pendingSession();
  const onError = mock((_error: Error) => {});
  const onReady = mock(() => {});
  const props = { fonts: [], file, [PEERLESS_FLAG]: true, experimentalWorkerOpen: false,
    onError, onReady } as unknown as PptxAsyncEditorProps;
  const view = render(<PptxEditor {...props} />);
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  expect(onError.mock.calls[0][0]).toBeInstanceOf(PptxAsyncEditorUnavailableError);
  expect(view.getByRole('alert').textContent).toContain('experimentalWorkerOpen');
  view.rerender(<PptxEditor {...props} />);
  expect(onError).toHaveBeenCalledTimes(1);
  expect(onReady).not.toHaveBeenCalled();
  expect(opener).not.toHaveBeenCalled();
  expect(view.container.querySelector('textarea')).toBeNull();
});

test('async API collaboration reaches the worker mount', async () => {
  const { opener } = pendingSession();
  const onError = mock((_error: Error) => {});
  const onReady = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }}
    collaboration={{ clientId: 1 }} onError={onError} onReady={onReady} />);
  await waitFor(() => expect(opener).toHaveBeenCalledTimes(1));
  expect(onError).not.toHaveBeenCalled();
  expect(view.queryByRole('alert')).toBeNull();
  expect(onReady).not.toHaveBeenCalled();
});

test('refusal_rerender_stable', async () => {
  setWorker(undefined);
  const { opener } = pendingSession();
  const onError = mock((_error: Error) => {});
  const onReady = mock(() => {});
  function Host() {
    const [error, setError] = useState<Error | null>(null);
    return <>
      <PptxEditor fonts={[]} file={file} {...{ [PEERLESS_FLAG]: true as const }}
        collaboration={{ clientId: 1 }} onReady={onReady} onError={(value) => {
          onError(value);
          expect(onError).toHaveBeenCalledTimes(1);
          setError(value);
        }} />
      <output>{error?.message}</output>
    </>;
  }
  const view = render(<Host />);
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  expect(onError.mock.calls[0][0]).toBeInstanceOf(PptxAsyncEditorUnavailableError);
  expect(view.container.querySelector('output')?.textContent).toBe(onError.mock.calls[0][0].message);
  view.rerender(<Host />);
  expect(onError).toHaveBeenCalledTimes(1);
  expect(view.getByRole('alert').textContent).toMatch(/Worker support/i);
  expect(onReady).not.toHaveBeenCalled();
  expect(opener).not.toHaveBeenCalled();
  expect(view.container.querySelector('textarea')).toBeNull();
  expect(view.container.querySelector('[role="application"]')).toBeNull();
});

test('refusal_reason_cycle', async () => {
  const { opener } = pendingSession();
  const onError = mock((_error: Error) => {});
  const props = { fonts: [], file, [PEERLESS_FLAG]: true as const, onError };
  const disabled = { ...props, experimentalWorkerOpen: false } as unknown as PptxAsyncEditorProps;
  const view = render(<PptxEditor {...disabled} />);
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  expect(onError.mock.calls[0][0]).toBeInstanceOf(PptxAsyncEditorUnavailableError);
  expect(opener).not.toHaveBeenCalled();
  view.rerender(<PptxEditor {...props} />);
  expect(view.queryByRole('alert')).toBeNull();
  await waitFor(() => expect(opener).toHaveBeenCalledTimes(1));
  view.rerender(<PptxEditor {...disabled} />);
  expect(onError).toHaveBeenCalledTimes(2);
  expect(view.getByRole('alert').textContent).toMatch(/experimentalWorkerOpen/i);
});

test('refusal_file_change', async () => {
  setWorker(undefined);
  const { opener } = pendingSession();
  const onError = mock((_error: Error) => {});
  const onReady = mock(() => {});
  const props = { fonts: [], file, [PEERLESS_FLAG]: true as const, onError, onReady };
  const view = render(<PptxEditor {...props} />);
  await waitFor(() => expect(onError).toHaveBeenCalledTimes(1));
  view.rerender(<PptxEditor {...props} file={new Uint8Array([2])} />);
  expect(onError).toHaveBeenCalledTimes(1);
  expect(view.getByRole('alert').textContent).toContain('Worker');
  expect(onReady).not.toHaveBeenCalled();
  expect(opener).not.toHaveBeenCalled();
  expect(view.container.querySelector('textarea')).toBeNull();
  expect(view.container.querySelector('[role="application"]')).toBeNull();
});

test('toggling the async API flag disposes the old session and mounts a fresh session', async () => {
  const { owner, opener } = pendingSession();
  const onReady = mock(() => {});
  const view = render(<PptxEditor fonts={[]} file={file} readOnly {...{ [PEERLESS_FLAG]: true as const }} onReady={onReady} />);
  await view.findByTestId('pptx-peerless-editor');
  view.rerender(<PptxEditor fonts={[]} file={file} readOnly {...{ [PEERLESS_FLAG]: false as const }} onReady={onReady} />);
  expect(owner.dispose).toHaveBeenCalledTimes(1);
  expect(view.queryByTestId('pptx-peerless-editor')).toBeNull();
  expect(onReady).not.toHaveBeenCalled();
  const next: PeerlessSession = { ...owner, dispose: mock(() => Promise.resolve()) };
  opener.mockReturnValue(next);
  view.rerender(<PptxEditor fonts={[]} file={file} readOnly {...{ [PEERLESS_FLAG]: true as const }} onReady={onReady} />);
  expect(await view.findByTestId('pptx-peerless-editor')).toBeTruthy();
  expect(opener).toHaveBeenCalledTimes(2);
  expect(owner.dispose).toHaveBeenCalledTimes(1);
  expect(next.dispose).not.toHaveBeenCalled();
  view.unmount();
  expect(next.dispose).toHaveBeenCalledTimes(1);
});
