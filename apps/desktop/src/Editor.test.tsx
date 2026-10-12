import { GlobalRegistrator } from "@happy-dom/global-registrator";
import type { GlobalWindow } from "happy-dom";
import { afterAll, afterEach, expect, mock, test } from "bun:test";
import { forwardRef, useEffect, useImperativeHandle } from "react";
import type { PptxAsyncEditorApi, PptxAsyncEditorProps } from "@betteroffice/pptx-react";
import type { EditorApi } from "./Editor";

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

interface DocxProps {
  onSave(buffer: ArrayBuffer): void;
  onError(error: Error): void;
  mode?: "editing" | "viewing";
  readOnly?: boolean;
  experimentalPeerless?: boolean;
  onChange?: () => void;
  onDocumentChange?: () => void;
}
let docxProps: DocxProps;
let save: (props: DocxProps) => Promise<ArrayBuffer | null> = async () =>
  null;
const presentationReplica = {
  clientId: 7,
  encodeStateVector: async () => Uint8Array.of(1),
  encodeStateAsUpdate: async () => Uint8Array.of(2),
  applyUpdate: async () => {},
  onUpdate: () => () => {},
};
const providerReplicas: unknown[] = [];
let presentationProps: PptxAsyncEditorProps;
const presentationOptions: PptxAsyncEditorProps['collaboration'][] = [];
let savePresentation = async () => Uint8Array.of(3);
mock.module("@betteroffice/pptx-react", () => ({
  PptxEditor: (props: PptxAsyncEditorProps) => {
    presentationProps = props;
    presentationOptions.push(props.collaboration);
    useEffect(() => props.onReady?.({ saveAsync: () => savePresentation(), focus() {} } as PptxAsyncEditorApi), [props.onReady]);
    useEffect(() => {
      props.collaboration?.onReplica?.(presentationReplica);
      return () => props.collaboration?.onReplica?.(null);
    }, [props.collaboration?.onReplica]);
    return null;
  },
}));
class Provider {
  constructor(replica: unknown) { providerReplicas.push(replica); }
  onStatus() { return () => {}; }
  onError() { return () => {}; }
  connect() {}
  destroy() {}
}
mock.module("@betteroffice/docx-react", () => ({
  DocxEditor: forwardRef<unknown, DocxProps>(function DocxEditor(props, ref) {
    docxProps = props;
    useImperativeHandle(ref, () => ({ save: () => save(props) }));
    return <span data-testid="docx-mode" data-mode={props.mode} data-readonly={String(props.readOnly ?? false)} />;
  }),
}));
for (const [specifier, exports] of Object.entries({
  "@betteroffice/docx-react/peerless-harness": { registerBrowserHarness: () => () => {} },
  "@betteroffice/docx-react/styles.css": {},
  "@betteroffice/xlsx-react": { XlsxEditor: () => null },
  "@betteroffice/docx/layout": { configureDefaultFonts: () => {} },
  "@betteroffice/docx/utils": { setGoogleFontsEnabled: () => {} },
  "@betteroffice/docx/collaboration": { CollaborationProvider: Provider },
  "@betteroffice/xlsx/collaboration": { CollaborationProvider: Provider },
  "@betteroffice/pptx/editor": { CollaborationProvider: Provider },
  "@betteroffice/fonts": {
    loadBundledFontBytes: () => {},
    resolveLastResortFace: () => {},
    resolveMetricCompatFace: () => {},
  },
}))
  mock.module(specifier, () => exports);

const { act, cleanup, fireEvent, render } = await import("@testing-library/react");
const { default: Editor } = await import("./Editor");

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test("peerless harness mode controls commit the editor read-only prop", async () => {
  const scope = globalThis as typeof globalThis & { __DOCX_PEERLESS_HARNESS__?: boolean };
  const previousFlag = scope.__DOCX_PEERLESS_HARNESS__;
  const previousUrl = location.href;
  scope.__DOCX_PEERLESS_HARNESS__ = true;
  const browser = window as unknown as GlobalWindow;
  browser.happyDOM.setURL("http://localhost/?format=docx&peerless=1");
  try {
    const view = render(<Editor
      file={{ bytes: new Uint8Array(), format: "docx" }} name="input.docx" room={null}
      onSave={() => {}} onChange={() => {}} onError={() => {}} onReady={() => {}}
      onOpen={async () => {}} onStatus={() => {}}
    />);
    const mode = view.getByTestId("docx-mode");
    expect(mode.dataset.readonly).toBe("false");
    fireEvent.click(view.getByRole("button", { name: "Viewing" }));
    expect(mode.dataset.mode).toBe("viewing");
    expect(mode.dataset.readonly).toBe("true");
    fireEvent.click(view.getByRole("button", { name: "Editing" }));
    expect(mode.dataset.mode).toBe("editing");
    expect(mode.dataset.readonly).toBe("false");
    view.unmount();
  } finally {
    browser.happyDOM.setURL(previousUrl);
    if (previousFlag === undefined) delete scope.__DOCX_PEERLESS_HARNESS__;
    else scope.__DOCX_PEERLESS_HARNESS__ = previousFlag;
  }
});

async function mount(format: "docx" | "pptx" = "docx", collaborative = false, onChange = () => {}) {
  const ready: { api?: EditorApi } = {};
  const errors: string[] = [];
  const saved: Uint8Array[] = [];
  await act(async () => {
    render(
      <Editor
        file={{ bytes: new Uint8Array(), format }}
        name="saved.docx"
        room={collaborative ? { id: "test", name: "Peer" } : null}
        onSave={(bytes) => saved.push(bytes)}
        onChange={onChange}
        onError={(error) => errors.push(error.message)}
        onReady={(api) => {
          if (api) ready.api = api;
        }}
        onOpen={async () => {}}
        onStatus={() => {}}
      />
    );
  });
  return { api: ready.api!, errors, saved };
}

test("peerless preview reports dirty state through worker document events", async () => {
  const previousUrl = location.href;
  const browser = window as unknown as GlobalWindow;
  browser.happyDOM.setURL("http://localhost/?format=docx&peerless=1");
  try {
    const changed = mock(() => {});
    await mount("docx", false, changed);
    expect(docxProps!.experimentalPeerless).toBe(true);
    expect(docxProps!.onChange).toBeUndefined();
    docxProps!.onDocumentChange?.();
    expect(changed).toHaveBeenCalledTimes(1);
  } finally {
    browser.happyDOM.setURL(previousUrl);
  }
});

test("a failed save rejects with the editor's error and a retry saves", async () => {
  const { api, errors, saved } = await mount();
  save = async ({ onError }) => {
    onError(new Error("The document changed while saving"));
    return null;
  };
  await expect(api.serialize()).rejects.toThrow(
    "The document changed while saving"
  );
  save = async ({ onSave }) => {
    onSave(Uint8Array.of(1).buffer);
    return Uint8Array.of(1).buffer;
  };
  expect(await api.serialize()).toEqual(Uint8Array.of(1));
  expect(errors).toEqual(["The document changed while saving"]);
  expect(saved).toEqual([]);
});

test("overlapping serializations run in turn and keep their own failure", async () => {
  const { api } = await mount();
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const calls: string[] = [];
  save = async ({ onError }) => {
    calls.push("first");
    save = async () => {
      calls.push("second");
      return Uint8Array.of(2).buffer;
    };
    await gate;
    onError(new Error("first failed"));
    return null;
  };
  const first = api.serialize().catch((cause: Error) => cause.message);
  const second = api.serialize();
  await new Promise((resolve) => setTimeout(resolve, 20));
  expect(calls).toEqual(["first"]);
  release();
  expect(await first).toBe("first failed");
  expect(await second).toEqual(Uint8Array.of(2));
  expect(calls).toEqual(["first", "second"]);
});

test("presentation serialization uses the async default and preserves save failures", async () => {
  const { api, saved } = await mount("pptx");
  expect(presentationProps.legacySyncApi).toBeUndefined();
  savePresentation = async () => { throw new Error("Worker stopped"); };
  await expect(api.serialize()).rejects.toThrow("Worker stopped");
  savePresentation = async () => Uint8Array.of(4);
  expect(await api.serialize()).toEqual(Uint8Array.of(4));
  expect(saved).toEqual([]);
});

test("presentation collaboration connects the async replica and presence provider", async () => {
  providerReplicas.length = 0;
  presentationOptions.length = 0;
  await mount("pptx", true);
  expect(providerReplicas).toEqual([presentationReplica]);
  expect(await presentationReplica.encodeStateVector()).toEqual(Uint8Array.of(1));
  expect(presentationProps.collaboration?.presence).toBeInstanceOf(Provider);
  const openingOptions = presentationOptions.filter((options) => options && !options.presence);
  expect(openingOptions.length).toBeGreaterThan(1);
  expect(new Set(openingOptions).size).toBe(1);
});
