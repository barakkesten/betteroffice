import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeEach, expect, mock, spyOn, test } from 'bun:test';
import type { PeerlessSession, PresentationCommandClient } from '@betteroffice/pptx';
import { PptxEditor } from '../PptxEditor';
import { peerlessSessionOpener } from './usePeerlessPresentation';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { cleanup, render, waitFor } = await import('@testing-library/react');
const restorers: (() => void)[] = [];
const file = new Uint8Array([1]);

beforeEach(() => {
  for (const capability of ['Worker', 'OffscreenCanvas'] as const) {
    const original = Object.getOwnPropertyDescriptor(globalThis, capability);
    Object.defineProperty(globalThis, capability, { configurable: true, value: class {} });
    restorers.push(() => {
      if (original) Object.defineProperty(globalThis, capability, original);
      else Reflect.deleteProperty(globalThis, capability);
    });
  }
});
afterEach(() => {
  cleanup();
  for (const restore of restorers.splice(0).reverse()) restore();
});
afterAll(async () => { if (ownsDom) await GlobalRegistrator.unregister(); });

function openingSession() {
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

test('default collaboration reaches worker startup', async () => {
  const { opener } = openingSession();
  const onError = mock((_error: Error) => {});
  const view = render(<PptxEditor fonts={[]} file={file} collaboration={{ clientId: 1 }} onError={onError} />);
  await waitFor(() => expect(opener).toHaveBeenCalledTimes(1));
  expect(opener).toHaveBeenCalledWith(file, expect.objectContaining({ collaboration: { clientId: 1 } }));
  expect(onError).not.toHaveBeenCalled();
  expect(view.queryByRole('alert')).toBeNull();
});

test('switching to legacy sync disposes a pending worker session', async () => {
  const { owner, opener } = openingSession();
  const view = render(<PptxEditor fonts={[]} file={file} readOnly />);
  await waitFor(() => expect(opener).toHaveBeenCalledTimes(1));
  view.rerender(<PptxEditor legacySyncApi experimentalWorkerOpen={false} fonts={[]} readOnly />);
  expect(owner.dispose).toHaveBeenCalledTimes(1);
  expect(view.queryByTestId('pptx-peerless-editor')).toBeNull();
  const next: PeerlessSession = { ...owner, dispose: mock(() => Promise.resolve()) };
  opener.mockReturnValue(next);
  view.rerender(<PptxEditor fonts={[]} file={file} readOnly />);
  await waitFor(() => expect(view.getByTestId('pptx-peerless-editor')).toBeTruthy());
  expect(opener).toHaveBeenCalledTimes(2);
  expect(next.dispose).not.toHaveBeenCalled();
  view.unmount();
  expect(next.dispose).toHaveBeenCalledTimes(1);
});
