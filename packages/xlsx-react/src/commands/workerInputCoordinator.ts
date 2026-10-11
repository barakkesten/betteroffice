import { XlsxCommandAdmissionError } from './createXlsxCommandStore';
import type { InputCoordinatorHooks, InputDraft, InputSeal } from './inputCoordinator';
import type {
  CommandStatus, WorkbookAck, WorkbookCommand, WorkbookCommandFrameRequest,
  WorkbookCommandSelection, WorkbookSession, XlsxRangeTarget,
} from '@betteroffice/xlsx';

export type WorkerInputKind =
  | 'commit'
  | 'input'
  | 'clipboard'
  | 'chart'
  | 'command'
  | 'host'
  | 'plugin';

export interface WorkerInputTarget {
  readonly sheet: number;
  readonly target: string;
}

export interface WorkerInputIntent<T = unknown> extends WorkerInputTarget {
  readonly id: number;
  readonly generation: number;
  readonly kind: WorkerInputKind;
  readonly input: T | undefined | Promise<T | undefined>;
  readonly draft: InputDraft | null;
}

export interface WorkerInputOperationOptions<T = unknown> {
  readonly kind?: Exclude<WorkerInputKind, 'commit'>;
  readonly target?: WorkerInputTarget;
  readonly input?: T;
  readonly recover?: boolean;
  readonly barrier?: boolean;
  readonly prepare?: () => void | Promise<void>;
}

export interface WorkerInputLease {
  (): void;
  check(): void;
  refuse(error: unknown, barrier?: boolean): void;
}

export class WorkerInputRefusal extends XlsxCommandAdmissionError {
  constructor(message: string, readonly barrier = false) {
    super('input-failed');
    this.message = message;
  }
}

export function inputRefusal(error: unknown): WorkerInputRefusal | null {
  if (error instanceof WorkerInputRefusal) return error;
  if (error instanceof Error && (error.name === 'NotAllowedError' || error.name === 'NotFoundError' ||
    error instanceof RangeError || /32,?767|out of range|invalid (sheet|cell|input|formula)|(text|formula).*limit/i.test(error.message))) {
    return new WorkerInputRefusal(error.message);
  }
  return null;
}

export type WorkerInputOperation<T, R> = (
  intent: WorkerInputIntent<T>,
  markApplied: WorkerInputLease
) => R | Promise<R>;

export interface WorkerInputCoordinatorHooks
  extends Pick<InputCoordinatorHooks, 'generation' | 'seal' | 'sync'> {
  write(draft: InputDraft, markApplied: WorkerInputLease): boolean;
  capture(): WorkerInputTarget;
  isReady(): boolean;
  isAcknowledged?(): boolean;
  /** Waits passively for the edit peer. */
  whenReady(): Promise<void>;
  resolveDraft?(draft: InputDraft): Promise<InputDraft>;
  sameDraft?(a: InputDraft, b: InputDraft): boolean;
  /** Resolves after the pending text has crossed a browser paint boundary. */
  preview(draft: InputDraft): Promise<void>;
  requestHydration(reason: string): void | Promise<void>;
  flushEdits(): Promise<void>;
  acknowledgeEdits?(): Promise<void>;
  onError?(error: unknown): void;
  onRefusal?(error: unknown, draft?: InputDraft): void;
}

export class WorkerInputNotReadyError extends Error {
  constructor() {
    super('The edit peer and accepted input must be flushed before a synchronous mutation');
    this.name = 'WorkerInputNotReadyError';
  }
}

export interface WorkerInputCoordinator {
  readonly draft: InputDraft | null;
  readonly committed: readonly InputDraft[];
  readonly rejected: readonly InputDraft[];
  readonly unapplied: readonly WorkerInputIntent[];
  readonly pending: boolean;
  readonly error: unknown;
  setDraft(draft: InputDraft | null): void;
  /** Accepts a commit; its write waits for readiness and paint. */
  submit(draft: InputDraft): boolean;
  submitAsync(draft: InputDraft): Promise<void>;
  settle(): boolean;
  settleAsync(): Promise<void>;
  input(
    write: WorkerInputOperation<unknown, boolean | void>,
    options?: WorkerInputOperationOptions
  ): Promise<void>;
  /** Captures clipboard access in the gesture; false refuses its write. */
  clipboard<T, R>(
    capture: () => T | Promise<T>,
    operation: (input: T, intent: WorkerInputIntent<T>, markApplied: WorkerInputLease) => R | Promise<R>,
    target?: WorkerInputTarget,
    prepare?: (input: T) => void | Promise<void>,
    recover?: boolean
  ): Promise<R>;
  /** Async mutators call markApplied before awaiting acknowledgement of an applied edit. */
  runAfterPendingInput<T, R>(
    operation: WorkerInputOperation<T, R>,
    options?: WorkerInputOperationOptions<T>
  ): Promise<R>;
  runSync<T, R>(
    operation: (intent: WorkerInputIntent<T>, markApplied: WorkerInputLease) => R,
    options?: WorkerInputOperationOptions<T>
  ): R;
  requestHydration(reason: string): Promise<void>;
  flush(): Promise<void>;
  fail(error: unknown): void;
  recover(): Promise<void>;
  drain(): Promise<void>;
  retire(): void;
  reset(): void;
}

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let rejectPromise!: (error: unknown) => void;
  let rejection: { error: unknown } | null = null;
  const cancellations = new Set<(error: unknown) => void>();
  const promise = new Promise<T>((done, failed) => {
    resolve = done;
    rejectPromise = failed;
  });
  void promise.catch(() => {});
  return {
    promise, resolve,
    reject(error: unknown) {
      rejection = { error };
      rejectPromise(error);
      for (const cancel of cancellations) cancel(error);
      cancellations.clear();
    },
    wait<R>(value: R | Promise<R>): Promise<R> {
      if (rejection) {
        void Promise.resolve(value).catch(() => {});
        return Promise.reject(rejection.error);
      }
      return new Promise<R>((done, failed) => {
        const cancel = (error: unknown) => { cancellations.delete(cancel); failed(error); };
        cancellations.add(cancel);
        void Promise.resolve(value).then((result) => {
          cancellations.delete(cancel);
          done(result);
        }, cancel);
      });
    },
  };
}

function sameCell(a: InputDraft, b: InputDraft): boolean {
  return (
    a.generation === b.generation &&
    a.source === b.source &&
    a.sheet === b.sheet &&
    a.row === b.row &&
    a.col === b.col
  );
}

function isPromise<T>(value: T | Promise<T>): value is Promise<T> {
  return typeof (value as Promise<T> | null)?.then === 'function';
}

interface Entry {
  intent: WorkerInputIntent;
  state: 'queued' | 'running' | 'failed';
  applied: boolean;
  draftWrite: boolean;
  recover: boolean;
  refused: boolean;
  prepared?: Promise<void>;
  barrier: boolean;
  barrierAfter: number;
  error?: unknown;
  run(markApplied: WorkerInputLease): unknown | Promise<unknown>;
  resolve(value: unknown): void;
  reject(error: unknown): void;
}

/** One FIFO for worker input, with retained unapplied work after failure. */
export function createWorkerInputCoordinator(
  hooks: WorkerInputCoordinatorHooks
): WorkerInputCoordinator {
  let draft: InputDraft | null = null;
  let rejected: readonly InputDraft[] = [];
  let entries: Entry[] = [];
  let snapshots = new WeakMap<InputDraft, InputDraft>();
  let written = new WeakSet<InputDraft>();
  let preparedDrafts = new WeakMap<InputDraft, Promise<InputDraft>>();
  let preparation: Promise<void> = Promise.resolve();
  let generation = hooks.generation();
  let epoch = 0;
  let cycle = deferred<never>();
  let replaced = deferred<never>();
  let active: typeof cycle | null = null;
  let work: Promise<void> | null = null;
  let recovery: Promise<void> | null = null;
  let failure: { error: unknown } | null = null;
  let reported = false;
  let inputFailed = false;
  let syncRunning = false;
  let nextId = 1;
  let refused: { id: number; error: unknown }[] = [];
  let acknowledgedRefusal = 0;

  const snapshot = (value: InputDraft): InputDraft => {
    let saved = snapshots.get(value);
    if (!saved) {
      saved = Object.freeze({ ...value });
      snapshots.set(value, saved);
      snapshots.set(saved, saved);
    }
    return saved;
  };

  const reset = () => {
    const error = new XlsxCommandAdmissionError('document-replaced');
    cycle.reject(error);
    replaced.reject(error);
    for (const entry of entries) entry.reject(error);
    cycle = deferred<never>();
    replaced = deferred<never>();
    epoch += 1;
    active = null;
    work = null;
    recovery = null;
    entries = [];
    refused = [];
    acknowledgedRefusal = 0;
    rejected = [];
    draft = null;
    snapshots = new WeakMap();
    written = new WeakSet();
    preparedDrafts = new WeakMap();
    preparation = Promise.resolve();
    failure = null;
    reported = false;
    inputFailed = false;
    generation = hooks.generation();
  };

  const current = () => {
    if (generation !== hooks.generation()) reset();
  };

  const check = (intent: WorkerInputIntent, lease: typeof cycle) => {
    if (lease !== cycle || intent.generation !== hooks.generation()) {
      throw new XlsxCommandAdmissionError('document-replaced');
    }
    if (failure) throw failure.error;
  };

  const wait = <T>(promise: T | Promise<T>, lease: typeof cycle): Promise<T> =>
    lease.wait(promise);

  const settleRejected = (value: InputDraft) => {
    rejected = rejected.filter((entry) => !sameCell(entry, value));
  };

  const coordinatorError = (): { error: unknown } | null => failure;

  const fail = (error: unknown) => {
    if (failure) return;
    failure = { error };
    cycle.reject(error);
    for (const entry of entries) entry.reject(error);
    for (const entry of entries) {
      if (!entry.applied && !entry.recover) {
        hooks.onRefusal?.(error);
        if (entry.intent.draft && !written.has(entry.intent.draft)) {
          entry.intent = Object.freeze({ ...entry.intent, kind: 'commit' });
          entry.recover = true;
          entry.run = (applied) => write(entry.intent.draft, entry.intent, cycle, applied);
        }
      }
    }
    entries = entries.filter((entry) => !entry.applied && entry.recover);
    if (!reported) {
      reported = true;
      try {
        hooks.onError?.(error);
      } catch {}
    }
  };

  const prepareDraft = (value: InputDraft): Promise<InputDraft> => {
    const existing = preparedDrafts.get(value);
    if (existing) return existing;
    const prepared = Promise.resolve().then(async () => {
      const resolved = hooks.resolveDraft ? await hooks.resolveDraft(value) : value;
      preparedDrafts.set(resolved, prepared);
      await hooks.preview(resolved);
      return resolved;
    });
    preparedDrafts.set(value, prepared);
    void prepared.catch(() => {});
    return prepared;
  };

  const write = async (value: InputDraft | null, intent: WorkerInputIntent, lease: typeof cycle, markApplied: WorkerInputLease) => {
    if (!value || written.has(value)) return;
    check(intent, lease);
    if (value.generation !== intent.generation) {
      throw new XlsxCommandAdmissionError('document-replaced');
    }
    const resolved = await wait(prepareDraft(value), lease);
    check(intent, lease);
    if (!hooks.isReady()) throw new WorkerInputNotReadyError();
    if (!hooks.write(resolved, markApplied)) {
      inputFailed = true;
      const correction = entries.some(
        (entry) =>
          entry.state === 'queued' &&
          entry.intent.kind === 'commit' &&
          entry.intent.draft !== null &&
          sameCell(entry.intent.draft, value)
      );
      if (!correction) {
        rejected = [...rejected.filter((entry) => !sameCell(entry, value)), value];
        if (draft === null) draft = value;
      }
      throw new XlsxCommandAdmissionError('input-failed');
    }
    written.add(value);
    written.add(resolved);
    const committed = entries.find((entry) => entry.intent.id === intent.id);
    if (committed?.intent.kind === 'commit') committed.applied = true;
    settleRejected(value);
    entries = entries.filter((entry) => {
      if (entry.state !== 'failed' || !entry.intent.draft || !sameCell(entry.intent.draft, value)) {
        return true;
      }
      written.add(entry.intent.draft);
      if (entry.intent.kind === 'commit') entry.applied = true;
      return entry.intent.kind !== 'commit';
    });
    if (draft && snapshot(draft) === value) draft = null;
  };

  const pump = (recovering = false): Promise<void> => {
    if (failure || (recovery && !recovering) || syncRunning) return Promise.resolve();
    if (active === cycle && work) return work;
    if (!entries.some((entry) => entry.state === 'queued')) return Promise.resolve();
    const lease = cycle;
    active = lease;
    const draining = async () => {
      while (lease === cycle && !failure) {
        let entry = entries[0];
        if (!entry) break;
        if (entry.state === 'failed' && entry.draftWrite && entry.intent.kind !== 'clipboard' &&
          entry.intent.draft && !written.has(entry.intent.draft)) {
          const blockedDraft = entry.intent.draft;
          const correction = entries[1];
          if (correction?.state === 'queued' && correction.intent.kind === 'commit' &&
            correction.intent.draft && sameCell(correction.intent.draft, blockedDraft)) {
            entries = [correction, ...entries.filter((candidate) => candidate !== correction)];
            entry = correction;
          }
        }
        if (entry.state !== 'queued') {
          for (const blocked of entries) {
            if (blocked.state === 'queued') blocked.reject(entry.error ?? new XlsxCommandAdmissionError('input-failed'));
          }
          break;
        }
        entry.state = 'running';
        try {
          check(entry.intent, lease);
          if (entry.prepared) await wait(entry.prepared, lease);
          await wait(hooks.whenReady(), lease);
          check(entry.intent, lease);
          if (!hooks.isReady()) throw new WorkerInputNotReadyError();
          if (!recovering && hooks.acknowledgeEdits) await wait(hooks.acknowledgeEdits(), lease);
          const markApplied: WorkerInputLease = Object.assign(() => {
            check(entry.intent, lease);
            entry.applied = true;
          }, {
            check: () => check(entry.intent, lease),
            refuse(error: unknown, barrier = false) {
              check(entry.intent, lease);
              entry.refused = true;
              if (barrier) refused.push({ id: entry.intent.id, error });
              hooks.onRefusal?.(error, entry.intent.kind === 'commit' ? entry.intent.draft ?? undefined : undefined);
            },
          });
          if (entry.barrier) {
            const covering = refused.find((item) => item.id > entry.barrierAfter && item.id < entry.intent.id);
            if (covering) {
              acknowledgedRefusal = Math.max(acknowledgedRefusal, covering.id);
              throw new WorkerInputRefusal(covering.error instanceof Error ? covering.error.message : String(covering.error));
            }
          }
          const result = await wait(entry.run(markApplied), lease);
          if (!recovering && hooks.acknowledgeEdits) await wait(hooks.acknowledgeEdits(), lease);
          check(entry.intent, lease);
          if (!entry.refused) entry.applied = true;
          entry.resolve(result);
          entries = entries.filter((candidate) => candidate !== entry);
        } catch (error) {
          if (lease !== cycle) return;
          entry.reject(error);
          const local = inputRefusal(error);
          if (local) {
            if (local.barrier) refused.push({ id: entry.intent.id, error: local });
            hooks.onRefusal?.(local, entry.intent.kind === 'commit' ? entry.intent.draft ?? undefined : undefined);
            entries = entries.filter((candidate) => candidate !== entry);
            continue;
          }
          if (entry.applied) entries = entries.filter((candidate) => candidate !== entry);
          else {
            entry.state = 'failed';
            entry.error = error;
          }
          if (error instanceof XlsxCommandAdmissionError) {
            if (error.code === 'document-replaced') {
              hooks.onRefusal?.(error);
              entries = entries.filter((candidate) => candidate !== entry);
              continue;
            }
            if (error.code === 'input-failed') {
              inputFailed = true;
              hooks.onRefusal?.(error);
            }
          } else {
            fail(error);
            return;
          }
        }
      }
      if (lease === cycle) inputFailed = entries.some((entry) => entry.state === 'failed');
    };
    const running = draining().finally(() => {
      if (active === lease) {
        active = null;
        work = null;
        if (entries[0]?.state === 'queued') void pump();
      }
    });
    work = running;
    return running;
  };

  const intent = <T>(
    kind: WorkerInputKind,
    input: WorkerInputIntent<T>['input'],
    sealed: InputDraft | null,
    target = hooks.capture()
  ): WorkerInputIntent<T> =>
    Object.freeze({
      id: nextId++,
      generation,
      kind,
      sheet: target.sheet,
      target: target.target,
      input,
      draft: sealed,
    });

  const record = <R>(
    saved: WorkerInputIntent,
    run: (entry: Entry, lease: typeof cycle, markApplied: WorkerInputLease) => R | Promise<R>
  ) => {
    const result = deferred<R>();
    const entry: Entry = {
      intent: saved,
      state: 'queued',
      applied: false,
      draftWrite: false,
      recover: true,
      refused: false,
      barrier: false,
      barrierAfter: acknowledgedRefusal,
      run: (markApplied) => run(entry, cycle, markApplied),
      resolve: (value) => result.resolve(value as R),
      reject: result.reject,
    };
    entries.push(entry);
    return { entry, result };
  };

  const accept = <R>(
    saved: WorkerInputIntent,
    run: (entry: Entry, lease: typeof cycle, markApplied: WorkerInputLease) => R | Promise<R>,
    draftWrite = false,
    options: WorkerInputOperationOptions = {}
  ): Promise<R> => {
    const { entry, result } = record(saved, run);
    entry.draftWrite = draftWrite;
    entry.recover = options.recover ?? true;
    entry.barrier = options.barrier ?? false;
    if (options.prepare) {
      entry.prepared = preparation.then(options.prepare);
      preparation = entry.prepared.catch(() => {});
    }
    if (failure) entry.reject(failure.error);
    const blocked = entries[0];
    if (blocked?.state === 'failed' && !(saved.kind === 'commit' && saved.draft &&
      blocked.draftWrite && blocked.intent.kind !== 'clipboard' && blocked.intent.draft &&
      !written.has(blocked.intent.draft) &&
      sameCell(saved.draft, blocked.intent.draft))) {
      entry.reject(blocked.error ?? new XlsxCommandAdmissionError('input-failed'));
    }
    void pump();
    return result.promise;
  };

  const submit = (finished: InputDraft): Promise<void> => {
    current();
    const saved = snapshot(finished);
    if (saved.generation !== generation) {
      return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    }
    if (draft === finished || (draft && snapshot(draft) === saved)) draft = null;
    settleRejected(saved);
    const accepted = intent('commit', saved, saved, {
      sheet: saved.sheet,
      target: `cell:${saved.row}:${saved.col}`,
    });
    return accept(accepted, (entry, lease, applied) => write(entry.intent.draft, entry.intent, lease, applied), true, {
      prepare: async () => {
        const resolved = await prepareDraft(saved);
        const entry = entries.find((item) => item.intent.id === accepted.id);
        if (entry) entry.intent = Object.freeze({ ...entry.intent, draft: resolved });
      },
    });
  };

  const runAfterPendingInput = <T, R>(
    operation: WorkerInputOperation<T, R>,
    options: WorkerInputOperationOptions<T> = {},
    captured?: { input: T | Promise<T> }
  ): Promise<R> => {
    current();
    if (rejected.length > 0) {
      return Promise.reject(new XlsxCommandAdmissionError('input-failed'));
    }
    let seal: InputSeal;
    try {
      seal = hooks.seal();
    } catch (error) {
      return Promise.reject(error);
    }
    if (seal.refused) return Promise.reject(new XlsxCommandAdmissionError(seal.refused));
    if (!seal.composition) hooks.sync();
    const original = draft;
    const saved = intent(
      options.kind ?? 'command',
      captured ? captured.input : structuredClone(options.input),
      original ? snapshot(original) : null,
      options.target
    );
    const acceptedEpoch = epoch;
    const composed = seal.composition?.then((ended) => {
      if (acceptedEpoch !== epoch || saved.generation !== hooks.generation()) {
        throw new XlsxCommandAdmissionError('document-replaced');
      }
      if (!ended) throw new XlsxCommandAdmissionError('input-failed');
      let composedDraft = saved.draft;
      if (saved.draft && draft && sameCell(saved.draft, draft) && (hooks.sameDraft?.(saved.draft, draft) ?? true)) {
        hooks.sync();
        composedDraft = draft ? snapshot(draft) : null;
      } else {
        const committed = entries.filter(
          (entry) =>
            entry.intent.id > saved.id &&
            entry.intent.kind === 'commit' &&
            saved.draft !== null &&
            entry.intent.draft !== null &&
            sameCell(saved.draft, entry.intent.draft) &&
            (hooks.sameDraft?.(saved.draft, entry.intent.draft) ?? true)
        );
        composedDraft = committed[committed.length - 1]?.intent.draft ?? saved.draft;
      }
      const entry = entries.find((candidate) => candidate.intent.id === saved.id);
      if (entry) entry.intent = Object.freeze({ ...entry.intent, draft: composedDraft });
      return composedDraft;
    });
    void composed?.catch(() => {});
    return accept(saved, async (entry, activeLease, markApplied) => {
      const sealed = composed ? await wait(composed, activeLease) : entry.intent.draft;
      check(entry.intent, activeLease);
      if (inputFailed || rejected.length > 0) {
        throw new XlsxCommandAdmissionError('input-failed');
      }
      if (sealed !== entry.intent.draft) {
        entry.intent = Object.freeze({ ...entry.intent, draft: sealed });
      }
      if (captured) {
        const input = await wait(captured.input, activeLease);
        check(entry.intent, activeLease);
        entry.intent = Object.freeze({ ...entry.intent, input });
      }
      await write(sealed, entry.intent, activeLease, Object.assign(() => {}, { check: markApplied.check, refuse: markApplied.refuse }));
      if (sealed && hooks.acknowledgeEdits) await wait(hooks.acknowledgeEdits(), activeLease);
      check(entry.intent, activeLease);
      const result = operation(entry.intent as WorkerInputIntent<T>, markApplied);
      if (!isPromise(result) && result !== false && !entry.refused) markApplied();
      return result;
    }, true, { ...options, prepare: async () => {
      if (saved.draft && !composed) {
        const resolved = await prepareDraft(saved.draft);
        const entry = entries.find((item) => item.intent.id === saved.id);
        if (entry) entry.intent = Object.freeze({ ...entry.intent, draft: resolved });
      }
      await options.prepare?.();
    } });
  };

  const requestHydration = async (reason: string) => {
    current();
    const acceptedEpoch = epoch;
    try {
      await replaced.wait(Promise.resolve(hooks.requestHydration(reason)));
      if (acceptedEpoch !== epoch || generation !== hooks.generation()) {
        throw new XlsxCommandAdmissionError('document-replaced');
      }
    } catch (error) {
      if (acceptedEpoch === epoch && generation === hooks.generation()) fail(error);
      throw error;
    }
  };

  const assertReady = () => {
    if (failure) throw failure.error;
    if (!hooks.isReady() || hooks.isAcknowledged?.() === false || entries.length > 0 || recovery || syncRunning || draft) {
      throw new WorkerInputNotReadyError();
    }
  };

  return {
    get draft() {
      return draft;
    },
    get committed() {
      return entries.flatMap((entry) => {
        const value = entry.intent.draft;
        return entry.draftWrite && !entry.applied && value && !written.has(value) ? [value] : [];
      });
    },
    get rejected() {
      return rejected;
    },
    get unapplied() {
      return entries.filter((entry) => !entry.applied).map((entry) => entry.intent);
    },
    get pending() {
      return entries.length > 0 || active !== null || recovery !== null;
    },
    get error() {
      return failure?.error ?? null;
    },
    setDraft(next) {
      current();
      draft = next;
    },
    submit(finished) {
      if (!finished.value.startsWith('=') && [...finished.value.replace(/^'/, '')].length > 32767) {
        hooks.onRefusal?.(new WorkerInputRefusal('Cell text exceeds 32,767 characters'), finished);
        return false;
      }
      void submit(finished).catch(() => {});
      return true;
    },
    submitAsync: submit,
    settle() {
      if (draft) void submit(draft).catch(() => {});
      return true;
    },
    settleAsync() {
      return draft ? submit(draft) : Promise.resolve();
    },
    input(operation, options = {}) {
      current();
      const saved = intent(
        options.kind ?? 'input',
        structuredClone(options.input),
        draft ? snapshot(draft) : null,
        options.target
      );
      return accept(saved, async (entry, lease, markApplied) => {
        const writing = operation(entry.intent, markApplied);
        if (!isPromise(writing) && writing !== false) markApplied();
        const result = await wait(writing, lease);
        if (result === false) throw new XlsxCommandAdmissionError('input-failed');
      });
    },
    clipboard(capture, operation, target, prepare, recover = true) {
      current();
      const origin = { ...(target ?? hooks.capture()) };
      const acceptedEpoch = epoch;
      let data: ReturnType<typeof capture>;
      try {
        const captured = capture();
        data = isPromise(captured)
          ? captured.then((value) => structuredClone(value))
          : structuredClone(captured);
      } catch (error) {
        return Promise.reject(error);
      }
      if (acceptedEpoch !== epoch || generation !== hooks.generation()) {
        if (isPromise(data)) void data.catch(() => {});
        return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
      }
      if (isPromise(data)) void data.catch(() => {});
      return runAfterPendingInput(
        async (intent, markApplied) => {
          markApplied.check();
          const writing = operation(intent.input as Awaited<typeof data>, intent, markApplied);
          if (!isPromise(writing) && writing !== false) markApplied();
          const result = await writing;
          if (result === false) throw new XlsxCommandAdmissionError('input-failed');
          return result;
        },
        { kind: 'clipboard', target: origin, recover, prepare: prepare ? async () => prepare(await data) : undefined },
        { input: data }
      );
    },
    runAfterPendingInput,
    runSync<T, R>(
      operation: (intent: WorkerInputIntent<T>, markApplied: WorkerInputLease) => R,
      options: WorkerInputOperationOptions<T> = {}
    ) {
      current();
      assertReady();
      const seal = hooks.seal();
      if (seal.refused) throw new XlsxCommandAdmissionError(seal.refused);
      if (seal.composition) throw new WorkerInputNotReadyError();
      hooks.sync();
      assertReady();
      const saved = intent<T>(
        options.kind ?? 'host',
        structuredClone(options.input),
        null,
        options.target
      );
      syncRunning = true;
      const { entry, result } = record(saved, (currentEntry, lease, markApplied) => {
        check(currentEntry.intent, lease);
        return operation(saved, markApplied);
      });
      entry.state = 'running';
      try {
        const lease = cycle;
        const value = operation(saved, Object.assign(() => {
          check(saved, lease);
          entry.applied = true;
        }, { check: () => check(saved, lease), refuse(error: unknown, barrier = false) {
          entry.refused = true;
          if (barrier) refused.push({ id: saved.id, error });
          hooks.onRefusal?.(error);
        } }));
        entry.applied = true;
        entries = entries.filter((candidate) => candidate !== entry);
        result.resolve(value);
        return value;
      } catch (error) {
        entry.state = 'failed';
        entry.error = error;
        entry.reject(error);
        if (entry.applied) entries = entries.filter((candidate) => candidate !== entry);
        const local = inputRefusal(error);
        if (local) {
          hooks.onRefusal?.(local);
          entries = entries.filter((candidate) => candidate !== entry);
        } else if (error instanceof XlsxCommandAdmissionError) {
          if (error.code === 'input-failed') inputFailed = true;
        } else {
          fail(error);
        }
        throw error;
      } finally {
        syncRunning = false;
        void pump();
      }
    },
    requestHydration,
    flush() {
      current();
      if (failure) return Promise.reject(failure.error);
      const lease = cycle;
      const barrierAfter = acknowledgedRefusal;
      const flushed = runAfterPendingInput(async (_, markApplied) => {
        markApplied();
        await hooks.flushEdits();
      }, { kind: 'host', recover: false, barrier: true });
      void flushed.catch(() => {});
      return Promise.all([requestHydration('flush'), flushed]).then(async () => {
        const throwFailure = () => { if (failure) throw failure.error; };
        let watermark: number;
        do {
          await wait(pump(), lease);
          current();
          if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
          throwFailure();
          const covering = refused.find((item) => item.id > barrierAfter);
          if (covering) {
            acknowledgedRefusal = Math.max(acknowledgedRefusal, covering.id);
            throw new WorkerInputRefusal(covering.error instanceof Error ? covering.error.message : String(covering.error));
          }
          if (entries.some((entry) => !entry.applied)) {
            throw new XlsxCommandAdmissionError('input-failed');
          }
          watermark = nextId;
          await wait(hooks.flushEdits(), lease);
          current();
          if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
          throwFailure();
        } while (nextId !== watermark || entries.length > 0);
      });
    },
    fail,
    recover() {
      current();
      if (recovery) return recovery;
      const lease = cycle;
      const recovered = async () => {
        await requestHydration('recovery');
        if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
        if (work) await work;
        if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
        cycle = deferred<never>();
        failure = null;
        inputFailed = false;
        rejected = [];
        preparedDrafts = new WeakMap();
        preparation = Promise.resolve();
        for (const entry of entries) { entry.state = 'queued'; entry.prepared = undefined; }
        do {
          await pump(true);
        } while (entries[0]?.state === 'queued' && !coordinatorError());
        const error = coordinatorError();
        if (error) throw error.error;
        if (entries.length > 0) throw new XlsxCommandAdmissionError('input-failed');
      };
      const result = Promise.resolve().then(recovered).finally(() => {
        if (recovery === result) {
          recovery = null;
          void pump();
        }
      });
      recovery = result;
      return result;
    },
    async drain() {
      if (draft) await submit(draft);
      await pump();
      if (entries.some((entry) => !entry.applied)) {
        const error = failure?.error ?? new XlsxCommandAdmissionError('input-failed');
        hooks.onRefusal?.(error);
        throw error;
      }
      await hooks.flushEdits();
    },
    retire() {
      const error = new XlsxCommandAdmissionError('document-replaced');
      for (const entry of entries) {
        if (entry.intent.kind !== 'clipboard') entry.reject(error);
      }
    },
    reset,
  };
}

export type PeerlessToolbarPayload = Extract<WorkbookCommand, { kind: 'applyOps' }>['payload'] extends infer P
  ? P extends { toolbar: infer T } ? T : never : never;

export type PeerlessWorkerCommand = WorkbookCommand;

export type PeerlessWorkerCommandInput = PeerlessWorkerCommand extends infer C
  ? C extends PeerlessWorkerCommand ? Pick<C, 'kind' | 'payload'> : never : never;

export interface PeerlessWorkerCommandSession extends Pick<WorkbookSession,
  'commandIdentity' | 'confirmedBasis' | 'throughSeq' | 'onAck' | 'save' | 'whenPresented'> {
  submit(command: PeerlessWorkerCommand): Promise<CommandStatus>;
  onFailure?: WorkbookSession['onFailure'];
  readonly failure?: unknown;
}

export interface PeerlessWorkerCommandTicket {
  readonly seq: number | null;
  readonly result: Promise<CommandStatus>;
}

export interface PeerlessWorkerCommandOptions {
  dispatch?<T>(write: () => T): T;
  acknowledge?<T>(publish: () => T): T;
  after?: PeerlessWorkerCommandTicket | number | null;
  compositionId?: string;
  frame?: WorkbookCommandFrameRequest;
  target?: XlsxRangeTarget;
  selection?: WorkbookCommandSelection;
  selectionRevision?: number;
}

export type PeerlessWorkerInputEvent = {
  type: 'dispatched' | 'accepted' | 'refused' | 'cancelled';
  seq: number | null;
  target: XlsxRangeTarget | null;
  draft: InputDraft | null;
  command?: PeerlessWorkerCommand;
  status?: CommandStatus;
  selectionRevision?: number;
} | {
  type: 'coverage';
  throughSeq: number;
  coversThroughSeq?: number;
  ack: WorkbookAck;
};

export interface PeerlessWorkerInputCoordinatorHooks
  extends Pick<InputCoordinatorHooks, 'generation' | 'seal' | 'sync'> {
  session(): PeerlessWorkerCommandSession | null;
  isReady(): boolean;
  compositionId?(): string | undefined;
  dependency?(draft: InputDraft): PeerlessWorkerCommandOptions['after'];
  target(sheet: number, row: number, col: number): XlsxRangeTarget;
  draftTarget?(draft: InputDraft): XlsxRangeTarget | undefined;
  frame?(): WorkbookCommandFrameRequest | undefined;
  selectionRevision?(): number;
  onEvent?(event: PeerlessWorkerInputEvent, publish?: () => void): void;
  onError?(error: unknown): void;
  maxOutstanding?: number;
  maxQueued?: number;
  readOnly?(): boolean;
}

export interface PeerlessWorkerInputCoordinator {
  readonly draft: InputDraft | null;
  readonly committed: readonly InputDraft[];
  readonly rejected: readonly InputDraft[];
  readonly pending: boolean;
  readonly error: unknown;
  readonly readOnly: boolean;
  readonly throughSeq: number;
  readonly submittedThroughSeq: number;
  setDraft(draft: InputDraft | null): void;
  discardRejected(draft: InputDraft): void;
  submit(draft: InputDraft, options?: PeerlessWorkerCommandOptions): PeerlessWorkerCommandTicket;
  command(input: PeerlessWorkerCommandInput, options?: PeerlessWorkerCommandOptions): PeerlessWorkerCommandTicket;
  clipboard<T>(
    capture: () => T | Promise<T>,
    command: (input: T) => PeerlessWorkerCommandInput,
    options?: PeerlessWorkerCommandOptions
  ): PeerlessWorkerCommandTicket;
  setReadOnly(readOnly: boolean): PeerlessWorkerCommandTicket;
  runAfterPendingInput<T>(operation: () => T | Promise<T>, fence?: boolean | 'posted'): Promise<T>;
  flush(): Promise<void>;
  whenApplied(seq?: number): Promise<void>;
  save(): Promise<Uint8Array>;
  whenPresented(seq?: number): Promise<void>;
  subscribe(listener: (event: PeerlessWorkerInputEvent) => void): () => void;
  fail(error: unknown): void;
  reset(): void;
  retire(): void;
}

interface PeerlessTicket extends PeerlessWorkerCommandTicket {
  acknowledge?: PeerlessWorkerCommandOptions['acknowledge'];
  lease?: object;
  outcome?: CommandStatus;
  failure?: { error: unknown };
}

interface PeerlessEntry {
  dispatch?: PeerlessWorkerCommandOptions['dispatch'];
  input: PeerlessWorkerCommandInput | null;
  draft: InputDraft | null;
  target: XlsxRangeTarget | null;
  after?: PeerlessWorkerCommandTicket | number;
  ticket: PeerlessTicket;
  result: ReturnType<typeof deferred<CommandStatus>>;
  frame?: WorkbookCommandFrameRequest;
  selection?: WorkbookCommandSelection;
  selectionRevision?: number;
  fence: boolean;
  prepare?: Promise<PeerlessWorkerCommandInput>;
}

interface PeerlessGate {
  run(): void | Promise<void>;
  reject(error: unknown): void;
}

function commandTarget(input: PeerlessWorkerCommandInput): XlsxRangeTarget | null {
  if ('target' in input.payload) return structuredClone(input.payload.target ?? null);
  if ('toolbar' in input.payload) return structuredClone(input.payload.toolbar.target ?? null);
  if ('selection' in input.payload) return structuredClone(input.payload.selection ?? null);
  return null;
}

function commandSelection(input: PeerlessWorkerCommandInput, selection?: WorkbookCommandSelection): PeerlessWorkerCommandInput {
  return selection && input.kind !== 'setReadOnly' && input.kind !== 'applyUpdate'
    ? { ...input, payload: { ...input.payload, selection } } as PeerlessWorkerCommandInput : input;
}

export function createPeerlessWorkerInputCoordinator(
  hooks: PeerlessWorkerInputCoordinatorHooks
): PeerlessWorkerInputCoordinator {
  const limit = (value: number | undefined, fallback: number) =>
    Number.isSafeInteger(value) && value! > 0 ? Math.min(value!, 128) : fallback;
  const maxOutstanding = limit(hooks.maxOutstanding, 64);
  const maxQueued = limit(hooks.maxQueued, 128);
  const listeners = new Set<(event: PeerlessWorkerInputEvent) => void>();
  const pendingGates = new Set<object>();
  let session: PeerlessWorkerCommandSession | null = null;
  let identity = '';
  let generation = hooks.generation();
  let cycle = deferred<never>();
  let queue: (PeerlessEntry | PeerlessGate)[] = [];
  let outstanding = new Map<number, PeerlessEntry>();
  const acknowledgements = new Map<number, PeerlessTicket>();
  let draft: InputDraft | null = null;
  let rejected: readonly InputDraft[] = [];
  let failure: { error: unknown } | null = null;
  let retired = false;
  let suspended = false;
  let nextSeq = 1;
  let throughSeq = 0;
  let boundary: PeerlessWorkerCommandTicket | undefined;
  let last: PeerlessWorkerCommandTicket | undefined;
  let compositions = new Map<string, PeerlessWorkerCommandTicket>();
  let submitted = new WeakMap<InputDraft, PeerlessWorkerCommandTicket>();
  let compositionIds = new WeakMap<Promise<boolean>, string>();
  let nextCompositionId = 1;
  let localReadOnly = hooks.readOnly?.() ?? false;
  let offAck = () => {};
  let offFailure = () => {};
  let pumping = false;
  let blocked: object | null = null;
  let insertion: number | null = null;
  let admittedReadOnly: boolean | null = null;
  let admittedAfter: { last: PeerlessWorkerCommandTicket | null;
    boundary: PeerlessWorkerCommandTicket | null } | undefined;

  const emit = (event: PeerlessWorkerInputEvent) => {
    let published = false;
    const publish = () => {
      if (published) return;
      published = true;
      for (const listener of listeners) { try { listener(event); } catch {} }
    };
    try { hooks.onEvent?.(event, publish); } catch {}
    publish();
  };
  const cancel = (entry: PeerlessEntry, error: unknown) => {
    if (entry.ticket.outcome || entry.ticket.failure) return;
    entry.ticket.failure = { error };
    if (boundary === entry.ticket) boundary = undefined;
    if (last === entry.ticket) last = undefined;
    entry.result.reject(error);
    emit({ type: 'cancelled', seq: entry.ticket.seq, target: entry.target, draft: entry.draft });
  };
  const invalidate = (error: unknown) => {
    offAck(); offFailure();
    cycle.reject(error);
    for (const entry of [...queue, ...outstanding.values()]) {
      if ('ticket' in entry) cancel(entry, error);
      else entry.reject(error);
    }
    queue = [];
    outstanding = new Map();
    acknowledgements.clear();
    blocked = null;
    cycle = deferred<never>();
    boundary = undefined;
    last = undefined;
    draft = null;
    rejected = [];
    compositions = new Map();
    compositionIds = new WeakMap();
    submitted = new WeakMap();
    pendingGates.clear();
  };
  const current = () => {
    if (retired) return;
    const next = hooks.session();
    const key = next ? `${next.commandIdentity.session}:${next.commandIdentity.incarnation}:${next.confirmedBasis.epoch}` : '';
    if (next === session && key === identity && generation === hooks.generation()) return;
    invalidate(new XlsxCommandAdmissionError('document-replaced'));
    session = next;
    identity = key;
    generation = hooks.generation();
    suspended = false;
    failure = null;
    localReadOnly = hooks.readOnly?.() ?? false;
    throughSeq = session?.throughSeq ?? 0;
    nextSeq = throughSeq + 1;
    const lease = cycle;
    offAck = session?.onAck((ack) => {
      current();
      if (lease !== cycle || retired || failure) return;
      if (ack.frame && (ack.frame.session !== session?.commandIdentity.session ||
        ack.frame.incarnation !== session?.commandIdentity.incarnation)) return;
      const committed = ack.statuses.filter((status) => status.committed && status.status === 'applied');
      const acknowledge = committed.length ? acknowledgements.get(committed[0].seq)?.acknowledge : undefined;
      for (const status of ack.statuses) receive(status);
      if (ack.throughSeq >= throughSeq) {
        throughSeq = ack.throughSeq;
        const publish = () => emit({ type: 'coverage', throughSeq, coversThroughSeq: ack.frame?.coversThroughSeq, ack });
        if (acknowledge && committed.every((status) => acknowledgements.get(status.seq)?.acknowledge === acknowledge)) {
          acknowledge(publish);
        } else publish();
        for (const seq of acknowledgements.keys()) if (seq <= throughSeq) acknowledgements.delete(seq);
      }
      pump();
    }) ?? (() => {});
    offFailure = session?.onFailure?.((error) => {
      current();
      if (lease === cycle) fail(error);
    }) ?? (() => {});
    if (session?.failure) fail(session.failure);
  };
  const fail = (error: unknown) => {
    if (failure || retired) return;
    failure = { error };
    localReadOnly = true;
    cycle.reject(error);
    for (const entry of [...queue, ...outstanding.values()]) {
      if ('ticket' in entry) cancel(entry, error);
      else entry.reject(error);
    }
    queue = [];
    outstanding.clear();
    acknowledgements.clear();
    pendingGates.clear();
    try { hooks.onError?.(error); } catch {}
  };
  const dependencyFailed = (entry: PeerlessEntry) => {
    const after = entry.after;
    return typeof after === 'object' && (after as PeerlessTicket).failure;
  };
  const receive = (status: CommandStatus) => {
    const entry = outstanding.get(status.seq);
    if (!entry) return;
    outstanding.delete(status.seq);
    entry.ticket.outcome = status;
    if (status.status === 'applied' || status.status === 'noop') {
      if (entry.draft) rejected = rejected.filter((value) => !sameCell(value, entry.draft!));
      entry.result.resolve(status);
      emit({ type: 'accepted', seq: status.seq, target: entry.target, draft: entry.draft, status });
    } else {
      const error = new PeerlessWorkerCommandRefusal(status, entry.target, entry.draft);
      entry.ticket.failure = { error };
      if (boundary === entry.ticket) boundary = undefined;
      if (last === entry.ticket) last = undefined;
      if (status.status === 'refused' && entry.draft) {
        rejected = [...rejected.filter((value) => !sameCell(value, entry.draft!)), entry.draft].slice(-maxQueued);
        if (!draft) draft = entry.draft;
      }
      entry.result.reject(error);
      emit({ type: status.status, seq: status.seq, target: entry.target, draft: entry.draft, status });
      for (const queued of queue) {
        if ('ticket' in queued && dependencyFailed(queued)) cancel(queued, error);
      }
    }
  };
  const check = () => {
    current();
    if (retired) throw new XlsxCommandAdmissionError('document-replaced');
    if (suspended) throw new XlsxCommandAdmissionError('document-replaced');
    if (failure) throw failure.error;
    if (!hooks.isReady() || !session?.commandIdentity.session || !session.commandIdentity.incarnation) {
      throw new XlsxCommandAdmissionError('editor-unavailable');
    }
  };
  const enqueue = (entry: PeerlessEntry | PeerlessGate) => {
    if (queue.length >= maxQueued) throw new WorkerInputRefusal('Worker input queue is full');
    if (insertion === null) queue.push(entry);
    else queue.splice(insertion++, 0, entry);
  };
  const pump = () => {
    current();
    if (pumping || blocked || failure || retired || suspended) return;
    pumping = true;
    try {
      while (queue.length && !failure && !retired) {
        const entry = queue[0];
        if ('ticket' in entry) {
          if (entry.ticket.failure) { queue.shift(); continue; }
          const failed = dependencyFailed(entry);
          if (failed) { queue.shift(); cancel(entry, failed.error); continue; }
          if (entry.prepare) {
            const lease = cycle;
            blocked = entry;
            void lease.wait(entry.prepare).then((input) => {
              if (lease !== cycle || blocked !== entry) return;
              try {
                entry.input = structuredClone(commandSelection(input, entry.selection));
                entry.target = commandTarget(entry.input);
                entry.prepare = undefined;
              } catch (error) { cancel(entry, error); }
              blocked = null;
              pump();
            }, (error) => {
              if (lease !== cycle || blocked !== entry) return;
              cancel(entry, error);
              blocked = null;
              pump();
            });
            break;
          }
          if (outstanding.size >= maxOutstanding) break;
          const dependency = typeof entry.after === 'object' ? entry.after as PeerlessTicket : undefined;
          if (dependency?.lease && dependency.lease !== cycle) {
            queue.shift(); cancel(entry, new XlsxCommandAdmissionError('document-replaced')); continue;
          }
          const afterSeq = dependency?.outcome ? undefined :
            typeof entry.after === 'number' ? entry.after : entry.after?.seq ?? undefined;
          if (typeof entry.after === 'number' && (!Number.isSafeInteger(entry.after) || entry.after <= 0 || entry.after >= nextSeq)) {
            queue.shift(); cancel(entry, new WorkerInputRefusal('Invalid command dependency')); continue;
          }
          if (entry.after && !dependency?.outcome && afterSeq === undefined) break;
          queue.shift();
          const dispatch = () => {
            const seq = nextSeq++;
            Object.defineProperty(entry.ticket, 'seq', { value: seq, configurable: true });
            const command = { ...structuredClone(entry.input!), ...session!.commandIdentity, seq,
              basis: { ...session!.confirmedBasis }, ...(afterSeq ? { afterSeq } : {}),
              ...(entry.frame ? { frame: entry.frame } : {}) } as PeerlessWorkerCommand;
            outstanding.set(seq, entry);
            if (entry.ticket.acknowledge) acknowledgements.set(seq, entry.ticket);
            const lease = cycle;
            emit({ type: 'dispatched', seq, target: entry.target, draft: entry.draft, command,
              selectionRevision: entry.selectionRevision });
            if (lease !== cycle || failure || retired) return;
            try {
              void session!.submit(command).then((status) => {
                current();
                if (lease !== cycle || failure || retired) return;
                receive(status);
                pump();
              }, (error) => {
                current();
                if (lease === cycle) fail(error);
              });
            } catch (error) { fail(error); }
          };
          try { if (entry.dispatch) entry.dispatch(dispatch); else dispatch(); }
          catch (error) { cancel(entry, error); }
          if (failure || retired) break;
        } else {
          queue.shift();
          insertion = 0;
          try {
            const result = entry.run();
            if (isPromise(result)) {
              blocked = entry;
              const lease = cycle;
              void lease.wait(result).then(() => {
                if (lease !== cycle || blocked !== entry) return;
                blocked = null; pump();
              }, (error) => {
                entry.reject(error);
                if (lease !== cycle || blocked !== entry) return;
                blocked = null; pump();
              });
              break;
            }
          } catch (error) { entry.reject(error); }
          finally { insertion = null; }
        }
      }
    } finally { pumping = false; }
  };
  const accept = (
    input: PeerlessWorkerCommandInput | null, value: InputDraft | null,
    options: PeerlessWorkerCommandOptions = {}, fence = false,
    prepare?: Promise<PeerlessWorkerCommandInput>
  ): PeerlessWorkerCommandTicket => {
    const result = deferred<CommandStatus>();
    const ticket: PeerlessTicket = { seq: null, result: result.promise, lease: cycle, acknowledge: options.acknowledge };
    let target: XlsxRangeTarget | null = null;
    try {
      target = input ? commandTarget(input) : null;
      check();
      ticket.lease = cycle;
      if ((admittedReadOnly ?? localReadOnly) && (prepare || input &&
        input.kind !== 'select' && input.kind !== 'setActiveSheet' && input.kind !== 'setReadOnly' && input.kind !== 'applyUpdate' &&
        !(input.kind === 'applyOps' && 'toolbar' in input.payload && input.payload.toolbar.id === 'refreshProposals'))) {
        throw new XlsxCommandAdmissionError('input-failed');
      }
      const entry: PeerlessEntry = {
        input: input ? structuredClone(commandSelection(input, options.selection)) : null,
        draft: value ? Object.freeze({ ...value }) : null,
        target,
        after: options.after === null ? undefined : options.after ??
          (admittedAfter !== undefined ? (fence ? admittedAfter.last : admittedAfter.boundary) ?? undefined : fence ? last : boundary),
        ticket, result, fence, prepare, dispatch: options.dispatch,
        frame: structuredClone('frame' in options ? options.frame : hooks.frame?.()),
        selection: structuredClone(options.selection),
        selectionRevision: 'selectionRevision' in options ? options.selectionRevision : hooks.selectionRevision?.(),
      };
      const previous = queue[queue.length - 1];
      if (!fence && value && previous && 'ticket' in previous && !previous.fence &&
        !previous.ticket.failure && !previous.prepare && previous.draft &&
        sameCell(previous.draft, value) && previous.after === entry.after &&
        JSON.stringify(previous.target) === JSON.stringify(entry.target)) {
        previous.input = entry.input;
        previous.draft = entry.draft;
        previous.frame = entry.frame;
        previous.selectionRevision = entry.selectionRevision;
        submitted.set(value, previous.ticket);
        return previous.ticket;
      }
      enqueue(entry);
      last = ticket;
      if (fence) boundary = ticket;
      pump();
    } catch (error) {
      ticket.failure = { error };
      result.reject(error);
      const saved = value ? Object.freeze({ ...value }) : null;
      if (saved) {
        rejected = [...rejected.filter((value) => !sameCell(value, saved)), saved].slice(-maxQueued);
        if (!draft) draft = saved;
      }
      emit({ type: 'refused', seq: null, target, draft: saved });
    }
    return ticket;
  };
  const submit = (value: InputDraft, options: PeerlessWorkerCommandOptions = {}) => {
    current();
    const existing = options.compositionId ? compositions.get(options.compositionId) : submitted.get(value);
    if (existing && (options.compositionId || !(existing as PeerlessTicket).failure)) return existing;
    let ticket: PeerlessWorkerCommandTicket;
    if (options.compositionId && value.value === '') {
      const outcome: CommandStatus = { seq: 0, status: 'noop', committed: false };
      ticket = { seq: null, result: Promise.resolve(outcome), outcome, lease: cycle } as PeerlessTicket;
    } else if (value.generation !== generation) {
      const result = deferred<CommandStatus>();
      const error = new XlsxCommandAdmissionError('document-replaced');
      result.reject(error);
      ticket = { seq: null, result: result.promise, failure: { error } } as PeerlessTicket;
    } else {
      try {
        const target = options.target ?? hooks.draftTarget?.(value) ?? hooks.target(value.sheet, value.row, value.col);
        ticket = accept({ kind: 'editCell', payload: {
          args: [value.sheet, value.row, value.col, value.value], target,
      } }, value, { ...options, after: options.after === undefined ? hooks.dependency?.(value) ?? null : options.after },
      Boolean(options.compositionId || options.selection));
      } catch (error) {
        const result = deferred<CommandStatus>();
        result.reject(error);
        ticket = { seq: null, result: result.promise, failure: { error } } as PeerlessTicket;
        const saved = Object.freeze({ ...value });
        rejected = [...rejected.filter((entry) => !sameCell(entry, saved)), saved].slice(-maxQueued);
        if (!draft) draft = saved;
        emit({ type: 'refused', seq: null, target: null, draft: saved });
      }
    }
    submitted.set(value, ticket);
    if (!(ticket as PeerlessTicket).failure && draft && sameCell(draft, value) && draft.value === value.value) draft = null;
    if (options.compositionId) {
      compositions.set(options.compositionId, ticket);
      if (compositions.size > maxQueued) compositions.delete(compositions.keys().next().value!);
    }
    return ticket;
  };
  const runAfterPendingInput = <T>(operation: () => T | Promise<T>, fence: boolean | 'posted' = false): Promise<T> => {
    const result = deferred<T>();
    try {
      check();
      if (pendingGates.size >= maxQueued) throw new WorkerInputRefusal('Worker input queue is full');
      pendingGates.add(result);
      void result.promise.then(() => pendingGates.delete(result), () => pendingGates.delete(result));
      const lease = cycle;
      const readOnly = localReadOnly;
      const after = { last: last ?? null, boundary: boundary ?? null };
      const seal = hooks.seal();
      if (seal.refused) throw new XlsxCommandAdmissionError(seal.refused);
      const composedTarget = seal.composition && draft ? { ...draft } : null;
      const compositionId = seal.composition ? hooks.compositionId?.() ?? compositionIds.get(seal.composition) ??
        `local-composition:${nextCompositionId++}` : undefined;
      if (seal.composition && compositionId) compositionIds.set(seal.composition, compositionId);
      let sealed: PeerlessWorkerCommandTicket | undefined;
      if (!seal.composition) { hooks.sync(); if (draft) sealed = submit(draft); }
      const ahead = [...outstanding.values(), ...queue.slice(0, insertion ?? queue.length)
        .filter((entry): entry is PeerlessEntry => 'ticket' in entry)]
        .map((entry) => entry.ticket.result);
      const invoke = () => {
        if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
        check();
        const previous = admittedReadOnly;
        const previousAfter = admittedAfter;
        admittedReadOnly = readOnly;
        admittedAfter = sealed ? { last: sealed, boundary: sealed } : after;
        try {
          const value = operation();
          void lease.wait(value).then(result.resolve, result.reject);
          return fence === true ? Promise.resolve(value).then(() => {}) : undefined;
        } finally { admittedReadOnly = previous; admittedAfter = previousAfter; }
      };
      const fencedInvoke = async () => {
        await lease.wait(Promise.all([
          ...ahead, ...[...outstanding.values()].map((entry) => entry.ticket.result),
        ]));
        if (rejected.length) throw new XlsxCommandAdmissionError('input-failed');
        await invoke();
      };
      enqueue({ reject: result.reject, run: () => {
        if (seal.composition || fence) return (async () => {
          if (seal.composition) {
            const ended = await lease.wait(seal.composition);
            if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
            pumping = true;
            insertion = 0;
            try {
              if (ended) {
                hooks.sync();
                const previous = admittedReadOnly;
                admittedReadOnly = readOnly;
                try {
                  if (draft) {
                    const finished = composedTarget ? { ...draft, generation: composedTarget.generation,
                      sheet: composedTarget.sheet, row: composedTarget.row, col: composedTarget.col,
                      source: composedTarget.source } : draft;
                    sealed = submit(finished, { compositionId });
                    ahead.push(sealed.result);
                  }
                }
                finally { admittedReadOnly = previous; }
              }
              if (fence) enqueue({ run: fencedInvoke, reject: result.reject });
              else invoke();
            } finally {
              insertion = null;
              pumping = false;
              blocked = null;
              pump();
            }
            return;
          }
          await fencedInvoke();
        })();
        return invoke();
      } });
      pump();
    } catch (error) { result.reject(error); }
    return result.promise;
  };
  current();
  return {
    get draft() { current(); return draft; },
    get committed() {
      current();
      return [...outstanding.values(), ...queue.filter((entry): entry is PeerlessEntry => 'ticket' in entry)]
        .flatMap((entry) => entry.draft && !entry.ticket.failure ? [entry.draft] : []);
    },
    get rejected() { current(); return rejected; },
    get pending() { current(); return queue.length > 0 || outstanding.size > 0 || blocked !== null; },
    get error() { current(); return failure?.error ?? null; },
    get readOnly() { current(); return localReadOnly || failure !== null || retired || suspended; },
    get throughSeq() { current(); return throughSeq; },
    get submittedThroughSeq() { current(); return nextSeq - 1; },
    setDraft(value) { current(); draft = value ? Object.freeze({ ...value }) : null; },
    discardRejected(value) {
      current();
      rejected = rejected.filter((entry) => !sameCell(entry, value));
      if (draft && sameCell(draft, value)) draft = null;
    },
    submit,
    command: (input, options) => accept(input, null, options, input.kind === 'setReadOnly'),
    clipboard(capture, command, options) {
      try {
        check();
        if (queue.length >= maxQueued) throw new WorkerInputRefusal('Worker input queue is full');
        const captured = { ...options, target: structuredClone(options?.target), selection: structuredClone(options?.selection),
          frame: structuredClone(options?.frame ?? hooks.frame?.()),
          selectionRevision: options?.selectionRevision ?? hooks.selectionRevision?.() };
        const lease = cycle;
        const data = capture();
        const target = captured.target;
        const prepare = lease.wait(data).then((value) => {
          check();
          if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
          const input = command(value);
          return target ? { ...input, payload: { ...input.payload, target } } as PeerlessWorkerCommandInput : input;
        });
        void prepare.catch(() => {});
        const result = deferred<CommandStatus>();
        const ticket: PeerlessTicket = { seq: null, result: result.promise, lease: cycle };
        const admitted = runAfterPendingInput(() => {
          const pasted = accept(null, null, captured, true, prepare);
          Object.defineProperty(ticket, 'seq', { get: () => pasted.seq });
          return pasted.result;
        });
        void admitted.then((status) => { ticket.outcome = status; result.resolve(status); }, (error) => {
          ticket.failure = { error }; result.reject(error);
        });
        return ticket;
      } catch (error) {
        const result = deferred<CommandStatus>();
        result.reject(error);
        return { seq: null, result: result.promise };
      }
    },
    setReadOnly(value) {
      current();
      const result = deferred<CommandStatus>();
      const ticket: PeerlessTicket = { seq: null, result: result.promise, lease: cycle };
      const admitted = runAfterPendingInput(() => {
        const mode = accept({ kind: 'setReadOnly', payload: { readOnly: value } }, null, { after: null }, true);
        Object.defineProperty(ticket, 'seq', { get: () => mode.seq });
        return mode.result;
      });
      void admitted.then((status) => { ticket.outcome = status; result.resolve(status); }, (error) => {
        if (!value && ticket.lease === cycle) localReadOnly = true;
        ticket.failure = { error }; result.reject(error);
      });
      localReadOnly = value;
      return ticket;
    },
    runAfterPendingInput,
    flush: () => runAfterPendingInput(() => {}, true),
    whenApplied(seq = nextSeq - 1) {
      try {
        check();
        if (!Number.isSafeInteger(seq) || seq < 0 || seq >= nextSeq) throw new WorkerInputRefusal('Invalid command sequence');
        const lease = cycle;
        return lease.wait(Promise.all([...outstanding.entries()].filter(([pending]) => pending <= seq)
          .map(([, entry]) => entry.ticket.result.catch((error) => {
            if (!(error instanceof PeerlessWorkerCommandRefusal)) throw error;
          })))).then(() => {
          if (lease !== cycle) throw new XlsxCommandAdmissionError('document-replaced');
          check();
        });
      } catch (error) { return Promise.reject(error); }
    },
    save: () => runAfterPendingInput(() => session!.save(), true),
    whenPresented: (seq) => runAfterPendingInput(() => session!.whenPresented(seq ?? nextSeq - 1), true),
    subscribe(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
    fail,
    reset() { current(); invalidate(new XlsxCommandAdmissionError('document-replaced')); suspended = true; },
    retire() { retired = true; invalidate(new XlsxCommandAdmissionError('document-replaced')); listeners.clear(); },
  };
}

export class PeerlessWorkerCommandRefusal extends WorkerInputRefusal {
  constructor(
    readonly status: CommandStatus, readonly target: XlsxRangeTarget | null,
    readonly draft: InputDraft | null
  ) {
    super(status.reason ?? 'Worker command refused');
    this.name = 'PeerlessWorkerCommandRefusal';
  }
}
