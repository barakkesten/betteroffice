import {
  createWorkbookEditPeer, failWorkbookEditPeer, hydratePeer, openWorkbookSession,
  workbookSessionRecovery, WorkbookRecoveryReplayError,
  type OpenWorkbookSessionOptions, type WorkbookHandle, type WorkbookSession, type WorkbookEditPeer,
} from '@betteroffice/xlsx';
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { XlsxEditorCollaborationOptions, XlsxEditorProps } from '../XlsxEditor';
import { XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import type { XlsxCommandStore } from '../commands/types';
import type { WorkerInputCoordinator } from '../commands/workerInputCoordinator';
import {
  createWorkerEditorApi, type WorkerEditorApiBridge, type WorkerEditorSessionAccess,
  type XlsxWorkerEditorApi, XlsxWorkerEditorCollaborationError,
} from './createWorkerEditorApi';
import { createAsyncEditorApi, type XlsxAsyncEditorApi } from './createAsyncEditorApi';
import type { WorkerDraftController } from './WorkerDraftController';
import type { PeerlessWorkerInputCoordinator } from '../commands/workerInputCoordinator';

export const editableWorkbookSessionBackend = {
  open: openWorkbookSession,
  hydrate: hydratePeer,
  attach: createWorkbookEditPeer,
};

export type EditableSessionWorkbookProps = Omit<XlsxEditorProps, 'onReady' | 'collaboration'> & {
  onError?: (error: Error) => void;
  onReady?: (api: XlsxWorkerEditorApi) => void | (() => void);
};

export interface EditableWorkbookSessionOptions {
  changed(): void;
  onError(error: Error): void;
  onReady(): void | (() => void);
  isCurrent(): boolean;
}

function asError(value: unknown): Error {
  return value instanceof Error ? value : new Error(
    value && typeof value === 'object' && 'message' in value ? String(value.message) : String(value)
  );
}

export class EditableWorkbookSession implements WorkerEditorSessionAccess {
  alive = true;
  peer: WorkbookHandle | null = null;
  editPeer: WorkbookEditPeer | null = null;
  failure: Error | null = null;
  recovering = false;
  retiring = false;
  beforeRetire: (() => void) | null = null;
  private painted = false;
  private hydration: Promise<void> | null = null;
  private readonly hydrated: Promise<void>;
  private resolveHydrated!: () => void;
  private rejectHydrated!: (error: unknown) => void;
  private readonly hydrationWaiters = new Set<{ recovery: boolean; reject(error: unknown): void }>();
  private offFailure = () => {};
  private cleanup: void | (() => void) = undefined;
  private input: WorkerInputCoordinator | null = null;

  constructor(
    readonly session: WorkbookSession,
    readonly generation: number,
    private readonly options: EditableWorkbookSessionOptions
  ) {
    this.hydrated = new Promise((resolve, reject) => {
      this.resolveHydrated = resolve;
      this.rejectHydrated = reject;
    });
    void this.hydrated.catch(() => {});
    this.offFailure = session.onFailure((error) => this.fail(error));
    if (session.failure) this.fail(session.failure);
  }

  get current(): boolean { return this.alive && (this.retiring || this.options.isCurrent()); }

  get ready(): boolean {
    return this.current && (!this.failure || this.recovering) && this.peer !== null &&
      (this.editPeer?.state === 'ready' || this.recovering && this.editPeer !== null);
  }

  whenHydrated(): Promise<void> {
    if (!this.current) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.recovering && this.peer && this.editPeer) return Promise.resolve();
    if (this.failure) return Promise.reject(this.failure);
    return this.hydrated;
  }

  connectInput(input: WorkerInputCoordinator | null): void {
    if (!this.current || this.retiring || this.input === input) return;
    this.input = input;
    if (this.failure) input?.fail(this.failure);
  }

  firstPaint(): void {
    if (!this.current || this.failure || this.painted) return;
    this.painted = true;
    try {
      const cleanup = this.options.onReady();
      if (!this.current && typeof cleanup === 'function') cleanup();
      else this.cleanup = cleanup;
    } catch (error) { this.fail(error); }
    if (this.current && !this.failure) void this.requestHydration('first-paint').catch(() => {});
  }

  requestHydration(reason: string): Promise<void> {
    if (!this.current) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.failure && reason !== 'recovery') return Promise.reject(this.failure);
    if (this.editPeer) return Promise.resolve();
    if (this.hydration) return this.waitForHydration(this.hydration, reason);
    const pending = (async () => {
      let peer: WorkbookHandle | null = null;
      try {
        peer = await editableWorkbookSessionBackend.hydrate(this.session);
        if (!this.current) throw new XlsxCommandAdmissionError('document-replaced');
        const edits = editableWorkbookSessionBackend.attach({
          session: this.session, peer, onError: (error) => this.fail(error),
        });
        if (!this.current) {
          edits.dispose();
          throw new XlsxCommandAdmissionError('document-replaced');
        }
        this.peer = peer;
        this.editPeer = edits;
        if (this.failure) failWorkbookEditPeer(edits, this.failure);
        peer = null;
        if (!this.failure) this.resolveHydrated();
        this.options.changed();
      } catch (error) {
        peer?.dispose();
        if (this.current) this.fail(error);
        throw error;
      }
    })();
    this.hydration = pending;
    void pending.catch(() => {
      if (this.hydration === pending) this.hydration = null;
    });
    return this.waitForHydration(pending, reason);
  }

  fail(value: unknown): void {
    if (!this.current || this.failure) return;
    this.failure = asError(value);
    if (this.editPeer) failWorkbookEditPeer(this.editPeer, this.failure);
    this.rejectHydrated(this.failure);
    for (const waiter of this.hydrationWaiters) {
      if (waiter.recovery) continue;
      this.hydrationWaiters.delete(waiter);
      waiter.reject(this.failure);
    }
    this.input?.fail(this.failure);
    this.options.changed();
    try { this.options.onError(this.failure); } catch {}
  }

  reportRefusal(value: unknown): void {
    try { this.options.onError(asError(value)); } catch {}
  }

  dispose(): void {
    if (!this.alive || this.retiring) return;
    this.retiring = true;
    try { this.beforeRetire?.(); } catch (error) { this.reportRefusal(error); }
    if (this.input?.pending || this.input?.draft) {
      void this.input.drain().catch((error) => {
        this.reportRefusal(error);
      }).finally(() => this.destroy());
    } else this.destroy();
  }

  private destroy(): void {
    if (!this.alive) return;
    this.alive = false;
    const error = new XlsxCommandAdmissionError('document-replaced');
    this.rejectHydrated(error);
    for (const waiter of this.hydrationWaiters) waiter.reject(error);
    this.hydrationWaiters.clear();
    this.offFailure();
    this.input?.reset();
    this.input = null;
    try { if (typeof this.cleanup === 'function') this.cleanup(); }
    catch (error) {
      try { this.options.onError(asError(error)); } catch {}
    }
    finally {
      try { this.editPeer?.dispose(); }
      finally {
        try { this.peer?.dispose(); }
        finally {
          this.editPeer = null;
          this.peer = null;
          void this.session.dispose().catch(() => {});
        }
      }
    }
  }

  private waitForHydration(pending: Promise<void>, reason: string): Promise<void> {
    if (!this.current) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.failure && reason !== 'recovery') return Promise.reject(this.failure);
    return new Promise<void>((resolve, reject) => {
      const waiter = { recovery: reason === 'recovery', reject };
      this.hydrationWaiters.add(waiter);
      void pending.then(() => {
        this.hydrationWaiters.delete(waiter);
        resolve();
      }, (error) => {
        this.hydrationWaiters.delete(waiter);
        reject(error);
      });
    });
  }
}

export class XlsxWorkerSessionFailedError extends Error {
  readonly code = 'worker-session-failed';
  constructor(readonly cause: unknown) {
    super(`Workbook worker session failed: ${asError(cause).message}`);
    this.name = 'XlsxWorkerSessionFailedError';
  }
}

interface CheckpointScheduler {
  now(): number;
  schedule(callback: () => void, delay: number): () => void;
}

export class PeerlessWorkbookSession {
  alive = true;
  retiring = false;
  failure: XlsxWorkerSessionFailedError | null = null;
  beforeRetire: (() => void) | null = null;
  readonly session: WorkbookSession;
  private painted = false;
  private cleanup: void | (() => void) = undefined;
  private readonly listeners = new Set<() => void>();
  private readonly pending = new Set<{ reject(error: unknown): void }>();
  private readonly readiness: Promise<void>;
  private resolveReady!: () => void;
  private rejectReady!: (error: unknown) => void;
  private offFailure = () => {};
  private offAck = () => {};
  private offDraft = () => {};
  private offInput = () => {};
  private recoveryInput: Pick<PeerlessWorkerInputCoordinator, 'draft' | 'pending'> | null = null;
  private recoveryDrafts: Pick<WorkerDraftController, 'compositionId' | 'subscribe'> | null = null;
  private checkpointTimer: (() => void) | undefined;
  private readonly checkpointScheduler: CheckpointScheduler;
  private checkpointRevision = 0;
  private checkpointBusy = false;
  private disposal: Promise<void> = Promise.resolve();
  private readonly disposeSession = () => this.dispose();

  constructor(
    private readonly workerSession: WorkbookSession,
    readonly generation: number,
    private readonly options: EditableWorkbookSessionOptions & { checkpointScheduler?: CheckpointScheduler }
  ) {
    this.checkpointScheduler = options.checkpointScheduler ?? {
      now: () => performance.now(),
      schedule(callback, delay) {
        const timer = setTimeout(callback, delay);
        return () => clearTimeout(timer);
      },
    };
    this.readiness = new Promise((resolve, reject) => {
      this.resolveReady = resolve;
      this.rejectReady = reject;
    });
    void this.readiness.catch(() => {});
    const methods = (target: object) => {
      const wrappers = new WeakMap<(...args: unknown[]) => Promise<unknown>, (...args: unknown[]) => Promise<unknown>>();
      return (key: PropertyKey) => {
        const method = Reflect.get(target, key, target);
        if (typeof method !== 'function') return method;
        let wrapper = wrappers.get(method);
        if (!wrapper) {
          wrapper = (...args: unknown[]) => this.track(() => Reflect.apply(method, target, args),
            key === 'submit' || key === 'command' || key === 'applyEdits');
          wrappers.set(method, wrapper);
        }
        return wrapper;
      };
    };
    const callMethod = methods(workerSession.call);
    const sessionMethod = methods(workerSession);
    const calls = new Proxy(workerSession.call, { get: (_, key) => callMethod(key) });
    this.session = new Proxy(workerSession, {
      get: (target, key) => {
        if (key === 'call') return calls;
        if (key === 'dispose') return this.disposeSession;
        if (key === 'submit' || key === 'save' || key === 'whenPresented') return sessionMethod(key);
        return Reflect.get(target, key, target);
      },
    });
    this.offFailure = workerSession.onFailure((error) => this.fail(error));
    this.offAck = workerSession.onAck(() => { this.notify(); this.scheduleCheckpoint(true); });
    if (workerSession.failure) this.fail(workerSession.failure);
  }

  get current(): boolean { return this.alive && this.options.isCurrent(); }
  get commandsAccepted(): boolean {
    const identity = this.workerSession.commandIdentity;
    return this.current && !this.retiring && !this.failure && this.workerSession.state.stage === 'ready' &&
      Boolean(identity.session && identity.incarnation);
  }
  get ready(): boolean { return this.commandsAccepted && this.painted; }
  get readOnly(): boolean { return !this.ready || this.workerSession.editState.readOnly === true; }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => { this.listeners.delete(listener); };
  }

  whenReady(): Promise<void> {
    if (!this.current || this.retiring) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.failure) return Promise.reject(this.failure);
    return this.readiness;
  }

  connectRecoveryInput(
    input: Pick<PeerlessWorkerInputCoordinator, 'draft' | 'pending'> & Partial<Pick<PeerlessWorkerInputCoordinator, 'subscribe'>>,
    drafts: Pick<WorkerDraftController, 'compositionId' | 'subscribe'>
  ): void {
    this.offDraft();
    this.offInput();
    this.recoveryInput = input;
    this.recoveryDrafts = drafts;
    this.offDraft = drafts.subscribe(() => this.scheduleCheckpoint());
    this.offInput = input.subscribe?.(() => this.scheduleCheckpoint()) ?? (() => {});
    this.scheduleCheckpoint();
  }

  recoverySave(): Promise<{ bytes: Uint8Array; recovery: true }> {
    const assertCurrent = () => {
      if (!this.current || this.retiring) throw new XlsxCommandAdmissionError('document-replaced');
    };
    return this.track(async () => {
      assertCurrent();
      if (this.workerSession.replica) throw new Error('Recovery is unavailable for collaborative workbooks');
      if (!this.failure) throw new Error('Recovery requires a failed workbook edit peer');
      try {
        const recovery = workbookSessionRecovery.get(this.workerSession);
        if (!recovery) throw new Error('Workbook recovery state is unavailable');
        const bytes = await recovery.save(assertCurrent);
        assertCurrent();
        return { bytes, recovery: true as const };
      } catch (error) {
        assertCurrent();
        if (error instanceof XlsxCommandAdmissionError) throw error;
        this.failure = error instanceof XlsxWorkerSessionFailedError ? error : new XlsxWorkerSessionFailedError(error);
        this.changed();
        try { this.options.onError(this.failure); } catch {}
        throw this.failure;
      }
    }, false, true);
  }

  firstPaint(): void {
    if (!this.commandsAccepted || this.painted) return;
    this.painted = true;
    try {
      const cleanup = this.options.onReady();
      if (!this.current && typeof cleanup === 'function') cleanup();
      else this.cleanup = cleanup;
    } catch (error) { this.fail(error); }
    if (this.ready) this.resolveReady();
    this.changed();
    this.scheduleCheckpoint();
  }

  fail(cause: unknown): void {
    if (!this.current || this.failure) return;
    this.failure = cause instanceof XlsxWorkerSessionFailedError ? cause : new XlsxWorkerSessionFailedError(cause);
    this.checkpointTimer?.();
    this.checkpointTimer = undefined;
    this.checkpointRevision += 1;
    this.rejectReady(this.failure);
    this.rejectPending(this.failure);
    this.changed();
    try { this.options.onError(this.failure); } catch {}
  }

  dispose(): Promise<void> {
    if (!this.alive || this.retiring) return this.disposal;
    this.retiring = true;
    try { this.beforeRetire?.(); }
    catch (error) { try { this.options.onError(asError(error)); } catch {} }
    this.alive = false;
    const error = new XlsxCommandAdmissionError('document-replaced');
    this.rejectReady(error);
    this.rejectPending(error);
    this.offFailure();
    this.offAck();
    this.offDraft();
    this.offInput();
    this.checkpointTimer?.();
    this.checkpointTimer = undefined;
    this.checkpointRevision += 1;
    this.recoveryInput = null;
    this.recoveryDrafts = null;
    this.changed();
    this.listeners.clear();
    try { if (typeof this.cleanup === 'function') this.cleanup(); }
    catch (error) { try { this.options.onError(asError(error)); } catch {} }
    finally {
      this.disposal = this.workerSession.dispose();
      void this.disposal.catch(() => {});
    }
    return this.disposal;
  }

  private changed(): void {
    this.options.changed();
    this.notify();
  }

  private notify(): void {
    for (const listener of this.listeners) { try { listener(); } catch {} }
  }

  private rejectPending(error: unknown): void {
    for (const request of this.pending) request.reject(error);
    this.pending.clear();
  }

  private scheduleCheckpoint(acknowledged = false): void {
    const revision = ++this.checkpointRevision;
    this.checkpointTimer?.();
    this.checkpointTimer = undefined;
    const recovery = workbookSessionRecovery.get(this.workerSession);
    const highWater = (recovery?.retainedCommands ?? 0) >= 192;
    const allowed = () => this.ready && Boolean(this.recoveryInput) &&
      !this.recoveryInput?.draft && !this.recoveryDrafts?.compositionId &&
      (highWater || revision === this.checkpointRevision && !this.recoveryInput?.pending);
    if (!recovery || recovery.retainedCommands < 64 || this.checkpointBusy || !allowed()) return;
    const capture = () => {
      this.checkpointTimer = undefined;
      if (!allowed()) return;
      this.checkpointBusy = true;
      const throughSeq = recovery.throughSeq;
      void recovery.capture(allowed).then((captured) => {
        this.checkpointBusy = false;
        if (captured && recovery.throughSeq > throughSeq || revision !== this.checkpointRevision) this.scheduleCheckpoint();
      }, (error) => {
        this.checkpointBusy = false;
        if (this.current && !this.retiring) this.fail(error);
      });
    };
    if (highWater) {
      if (acknowledged || !this.recoveryInput?.pending) capture();
      return;
    }
    const started = this.checkpointScheduler.now();
    const idleMs = recovery.checkpointIdleMs ?? 150;
    const idle = () => {
      this.checkpointTimer = undefined;
      if (!allowed()) return;
      const remaining = idleMs - (this.checkpointScheduler.now() - started);
      if (remaining > 0) this.checkpointTimer = this.checkpointScheduler.schedule(idle, remaining);
      else capture();
    };
    this.checkpointTimer = this.checkpointScheduler.schedule(idle, idleMs);
  }

  private track<T>(operation: () => Promise<T>, requiresReady = false, recovery = false): Promise<T> {
    if (!this.current || this.retiring) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (this.failure && !recovery) return Promise.reject(this.failure);
    if (requiresReady && !this.ready) return Promise.reject(new XlsxCommandAdmissionError('editor-unavailable'));
    return new Promise<T>((resolve, reject) => {
      const request = { reject };
      this.pending.add(request);
      try {
        void operation().then(resolve, (error) => {
          if (error instanceof WorkbookRecoveryReplayError) this.fail(error);
          reject(recovery ? error : this.failure ?? error);
        }).finally(() => this.pending.delete(request));
      } catch (error) {
        this.pending.delete(request);
        if (error instanceof WorkbookRecoveryReplayError) this.fail(error);
        reject(this.failure ?? error);
      }
    });
  }
}

export type PeerlessSessionWorkbookProps = Omit<EditableSessionWorkbookProps, 'onReady'> & {
  collaboration?: XlsxEditorCollaborationOptions;
  onReady?: (api: XlsxAsyncEditorApi) => void | (() => void);
};

interface SessionWorkbookOptions {
  backend: 'legacy' | 'peerless';
  initialViewport?: () => OpenWorkbookSessionOptions['initialViewport'];
}

interface SessionWorkbookResult<R> {
  run: R | null;
  error: Error | null;
  loading: boolean;
  reportError(value: unknown): void;
}

export function useEditableSessionWorkbook(
  props: PeerlessSessionWorkbookProps, commands: XlsxCommandStore, bridge: WorkerEditorApiBridge,
  options: SessionWorkbookOptions & { backend: 'peerless' }
): SessionWorkbookResult<PeerlessWorkbookSession>;
export function useEditableSessionWorkbook(
  props: EditableSessionWorkbookProps, commands: XlsxCommandStore, bridge: WorkerEditorApiBridge,
  options?: SessionWorkbookOptions & { backend: 'legacy' }
): SessionWorkbookResult<EditableWorkbookSession>;
export function useEditableSessionWorkbook(
  props: EditableSessionWorkbookProps | PeerlessSessionWorkbookProps, commands: XlsxCommandStore,
  bridge: WorkerEditorApiBridge, options: SessionWorkbookOptions
): SessionWorkbookResult<EditableWorkbookSession | PeerlessWorkbookSession>;

export function useEditableSessionWorkbook(
  props: EditableSessionWorkbookProps | PeerlessSessionWorkbookProps, commands: XlsxCommandStore,
  bridge: WorkerEditorApiBridge, options?: SessionWorkbookOptions
): SessionWorkbookResult<EditableWorkbookSession | PeerlessWorkbookSession> {
  const backend = useRef(options?.backend ?? 'legacy').current;
  return backend === 'peerless'
    ? usePeerlessSessionWorkbook(props as PeerlessSessionWorkbookProps, commands, bridge, options ?? { backend })
    : useLegacySessionWorkbook(props as EditableSessionWorkbookProps, commands, bridge);
}

function useLegacySessionWorkbook(
  props: EditableSessionWorkbookProps, commands: XlsxCommandStore, bridge: WorkerEditorApiBridge
): {
  run: EditableWorkbookSession | null;
  error: Error | null;
  loading: boolean;
  reportError(value: unknown): void;
} {
  const latest = useRef({ props, bridge });
  latest.current = { props, bridge };
  const generation = useRef(0);
  const [run, setRun] = useState<EditableWorkbookSession | null>(null);
  const [, setRevision] = useState(0);
  const [error, setError] = useState<Error | null>(null);
  const [loading, setLoading] = useState(false);
  const collaboration = (props as EditableSessionWorkbookProps & Pick<XlsxEditorProps, 'collaboration'>).collaboration;
  const reportError = useCallback((value: unknown) => {
    const error = asError(value);
    setError(error);
    try { latest.current.props.onError?.(error); } catch {}
  }, []);

  useEffect(() => {
    const token = ++generation.current;
    const controller = new AbortController();
    let disposed = false;
    let opened: EditableWorkbookSession | undefined;
    const current = () => !disposed && token === generation.current && latest.current.props.file === props.file;
    const changed = () => { if (current()) setRevision((revision) => revision + 1); };
    setRun(null);
    setError(null);
    setLoading(Boolean(props.file));
    if (collaboration) {
      reportError(new XlsxWorkerEditorCollaborationError());
      setLoading(false);
    } else if (props.file) void (async () => {
      try {
        const session = await editableWorkbookSessionBackend.open(props.file!, {
          signal: controller.signal, retainPeerHydration: true,
        });
        if (!current()) { void session.dispose().catch(() => {}); return; }
        opened = new EditableWorkbookSession(session, token, {
          changed, isCurrent: current,
          onError: (error) => {
            if (current()) reportError(error);
            else { try { latest.current.props.onError?.(error); } catch {} }
          },
          onReady: () => current() ? latest.current.props.onReady?.(api) : undefined,
        });
        const api = createWorkerEditorApi(opened, commands, () => latest.current.bridge);
        setRun(opened);
        setLoading(false);
      } catch (error) {
        if (!current()) return;
        if (opened) opened.fail(error);
        else reportError(error);
      } finally {
        if (current()) setLoading(false);
      }
    })();
    return () => {
      disposed = true;
      generation.current += 1;
      controller.abort();
      opened?.dispose();
    };
  }, [props.file, collaboration, commands, reportError]);

  useLayoutEffect(() => {
    if (run?.current) run.connectInput(bridge.coordinator());
  });
  useLayoutEffect(() => () => run?.dispose(), [run, props.file, collaboration]);

  return { run: run?.current ? run : null, error, loading, reportError };
}

function usePeerlessSessionWorkbook(
  props: PeerlessSessionWorkbookProps, commands: XlsxCommandStore,
  bridge: WorkerEditorApiBridge, options: SessionWorkbookOptions
): SessionWorkbookResult<PeerlessWorkbookSession> {
  const latest = useRef({ props, bridge, options });
  latest.current = { props, bridge, options };
  const generation = useRef(0);
  const [run, setRun] = useState<PeerlessWorkbookSession | null>(null);
  const [, setRevision] = useState(0);
  const [error, setError] = useState<Error | null>(null);
  const [loading, setLoading] = useState(false);
  const collaboration = props.collaboration;
  const collaborationKey = collaboration !== undefined;
  const collaborationClientId = collaboration?.clientId;
  const collaborationInitialUpdate = collaboration?.initialUpdate;
  const reportError = useCallback((value: unknown) => {
    const error = asError(value);
    setError(error);
    try { latest.current.props.onError?.(error); } catch {}
  }, []);

  useEffect(() => {
    const token = ++generation.current;
    const controller = new AbortController();
    let disposed = false;
    let opened: PeerlessWorkbookSession | undefined;
    const current = () => !disposed && token === generation.current && latest.current.props.file === props.file;
    const changed = () => { if (current()) setRevision((revision) => revision + 1); };
    setRun(null);
    setError(null);
    setLoading(Boolean(props.file));
    if (props.file) void (async () => {
      try {
        const session = await editableWorkbookSessionBackend.open(props.file!, {
          signal: controller.signal,
          includeEditState: true,
          ...(collaboration ? { collaborative: true,
            clientId: collaborationClientId, initialUpdate: collaborationInitialUpdate?.slice() } : {}),
          ...(latest.current.options.initialViewport
            ? { initialViewport: latest.current.options.initialViewport() } : {}),
        });
        if (!current()) { void session.dispose().catch(() => {}); return; }
        const lifecycle = {
          changed, isCurrent: current,
          onError: (error: Error) => {
            if (current()) reportError(error);
            else { try { latest.current.props.onError?.(error); } catch {} }
          },
        };
        const peerless: PeerlessWorkbookSession = new PeerlessWorkbookSession(session, token, {
          ...lifecycle, onReady: () => current() ? latest.current.props.onReady?.(api) : undefined,
        });
        const api = createAsyncEditorApi(peerless, commands, () => latest.current.bridge);
        opened = peerless;
        setRun(opened);
        setLoading(false);
      } catch (error) {
        if (!current()) return;
        if (opened) opened.fail(error);
        else reportError(new XlsxWorkerSessionFailedError(error));
      } finally {
        if (current()) setLoading(false);
      }
    })();
    return () => {
      disposed = true;
      generation.current += 1;
      controller.abort();
      opened?.dispose();
    };
  }, [props.file, collaborationKey, collaborationClientId, collaborationInitialUpdate, commands, reportError]);

  useLayoutEffect(() => () => { run?.dispose(); }, [run, props.file, collaborationKey, collaborationClientId, collaborationInitialUpdate]);

  const onReplica = collaboration?.onReplica;
  const replicaReady = run instanceof PeerlessWorkbookSession && run.ready;
  useEffect(() => {
    if (!replicaReady || !(run instanceof PeerlessWorkbookSession) || !onReplica) return;
    onReplica(run.session.replica ?? null);
    return () => onReplica(null);
  }, [run, replicaReady, onReplica]);

  return { run: run?.current ? run : null, error, loading, reportError };
}
