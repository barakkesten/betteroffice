import { useLayoutEffect, useRef } from 'react';
import { normalizeRange, StaleProposalError } from '@betteroffice/xlsx';
import type {
  EditResult,
  MergedRange,
  SelectionFormatting,
  WorkbookHandle,
  WorkbookEditPeer,
  WorkbookAck,
  WorkbookCommandSelection,
  WorkbookCommandFrameRequest,
  WorkbookEditState,
  WorkbookSession,
  XlsxRangeTarget,
} from '@betteroffice/xlsx';
import type { WorkbookFormatTokenResult } from '../../../xlsx/src/session/commandProtocol';
import type { MergeAction } from '../components/Toolbar';
import {
  XlsxCommandAdmissionError, type XlsxCommandBinding, type XlsxCommandController,
} from './createXlsxCommandStore';
import {
  commandReason,
  DEFAULT_FONT_FAMILIES,
  DEFAULT_FONT_SIZES,
  isCellCommand,
  type XlsxCommandEnvironment,
  type XlsxSelectionEnvironment,
} from './evaluate';
import { XLSX_COMMAND_DESCRIPTORS } from './descriptors';
import { inputRefusal, WorkerInputRefusal, type WorkerInputCoordinator } from './workerInputCoordinator';
import {
  PeerlessWorkerCommandRefusal, type PeerlessWorkerCommandInput, type PeerlessWorkerCommandSession,
  type PeerlessWorkerInputCoordinator,
} from './workerInputCoordinator';
import type { XlsxEditorBridge } from './useXlsxCommands';
import type {
  XlsxCommandArgs,
  XlsxCommandFailureCode,
  XlsxCommandId,
  XlsxCommandResult,
} from './types';

export interface WorkerXlsxEditorBridge extends Omit<XlsxEditorBridge, 'handle' | 'coordinator' | 'exportPng'> {
  peer(): WorkbookHandle | null;
  editPeer(): WorkbookEditPeer | null;
  coordinator: WorkerInputCoordinator;
  preview(): Promise<void>;
  recovering?(): boolean;
  refuse?(reason: { message: string }): void;
}

interface EngineRead {
  key: string;
  formatting: SelectionFormatting | null;
  merged: readonly MergedRange[];
  canUndo: boolean;
  canRedo: boolean;
}

const IMMEDIATE: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>([
  'zoom',
  'searchMenus',
  'proposalsPanel',
]);

function rangeA1(peer: WorkbookHandle, range: XlsxSelectionEnvironment): string {
  const from = peer.cell(range.sheet, range.top, range.left).a1;
  const to = peer.cell(range.sheet, range.bottom, range.right).a1;
  return `${from}:${to}`;
}

function readEngine(
  peer: WorkbookHandle,
  sheet: number,
  range: ReturnType<typeof normalizeRange>,
  key: string
): EngineRead {
  const read: EngineRead = { key, formatting: null, merged: [], canUndo: false, canRedo: false };
  try {
    const a1 = rangeA1(peer, { sheet, ...range, merged: [], formatting: null });
    read.formatting = peer.selectionFormatting(sheet, a1);
    read.merged = peer.mergedRanges(sheet, a1);
  } catch {}
  try {
    const history = peer.historyState();
    read.canUndo = history.canUndo;
    read.canRedo = history.canRedo;
  } catch {}
  return read;
}

function mergeOps(
  sheet: number,
  selection: XlsxSelectionEnvironment,
  action: MergeAction
): unknown[] {
  const range = (top: number, left: number, bottom: number, right: number) => ({
    start: { row: top, col: left },
    end: { row: bottom, col: right },
  });
  const ops: unknown[] = [];
  if (action === 'all') {
    ops.push({
      type: 'mergeCells',
      sheet,
      range: range(selection.top, selection.left, selection.bottom, selection.right),
    });
  } else if (action === 'horizontal') {
    for (let row = selection.top; row <= selection.bottom; row++) {
      ops.push({ type: 'mergeCells', sheet, range: range(row, selection.left, row, selection.right) });
    }
  } else if (action === 'vertical') {
    for (let col = selection.left; col <= selection.right; col++) {
      ops.push({ type: 'mergeCells', sheet, range: range(selection.top, col, selection.bottom, col) });
    }
  } else {
    for (const merged of selection.merged) {
      ops.push({
        type: 'unmergeCells',
        sheet,
        range: range(merged.start.row, merged.start.col, merged.end.row, merged.end.col),
      });
    }
  }
  return ops;
}

export function createWorkerXlsxCommandBinding(
  bridge: () => WorkerXlsxEditorBridge
): XlsxCommandBinding {
  let cache: EngineRead | null = null;
  let markApplied: (() => void) | null = null;

  const engine = (executing: boolean): EngineRead | null => {
    const current = bridge();
    const peer = current.peer();
    const view = current.view();
    if (!peer || current.status() !== 'ready') return null;
    const range = view.selection
      ? normalizeRange(view.selection)
      : { top: 0, left: 0, bottom: 0, right: 0 };
    const key = [
      current.generation(),
      current.mutation(),
      view.sheet,
      range.top,
      range.left,
      range.bottom,
      range.right,
    ].join(':');
    if (!executing && cache?.key === key) return cache;
    cache = readEngine(peer, view.sheet, range, key);
    return cache;
  };

  const environment = (executing: boolean): XlsxCommandEnvironment => {
    const current = bridge();
    const view = current.view();
    const read = engine(executing);
    let selection: XlsxCommandEnvironment['selection'] = null;
    if (view.chartSelected) selection = 'chart';
    else if (view.selection) {
      selection = {
        sheet: view.sheet,
        ...normalizeRange(view.selection),
        merged: read?.merged ?? [],
        formatting: read?.formatting ?? null,
      };
    }
    return {
      status: current.status(),
      readOnly: current.readOnly(),
      collaborative: current.collaborative(),
      selection,
      canUndo: read?.canUndo ?? false,
      canRedo: read?.canRedo ?? false,
      pendingInput: current.coordinator.draft !== null || current.coordinator.committed.length > 0,
      zoom: view.zoom,
      paintFormat: view.capturedFormat !== null,
      borderStyle: view.borderStyle ?? read?.formatting?.borderStyle ?? null,
      borderColor: view.borderColor ?? read?.formatting?.borderColor ?? null,
      pngExport: false,
      proposals: view.proposalsAvailable
        ? view.proposals.map((proposal) => ({ id: proposal.id, label: proposal.agentId }))
        : null,
      proposalsPanelOpen: view.proposalsPanelOpen,
      fontFamilies: DEFAULT_FONT_FAMILIES,
      fontSizes: DEFAULT_FONT_SIZES,
      translate: current.translate,
    };
  };

  const perform = <K extends XlsxCommandId>(
    id: K,
    args: XlsxCommandArgs[K],
    env: XlsxCommandEnvironment
  ): XlsxCommandResult | Promise<XlsxCommandResult> => {
    const current = bridge();
    const peer = current.peer();
    const fail = (code: XlsxCommandFailureCode): XlsxCommandResult => ({
      ok: false,
      failure: commandReason(code, env),
    });
    if (id === 'zoom') {
      const scale = (args as XlsxCommandArgs['zoom']).scale;
      if (env.zoom === scale) return { ok: true, status: 'noop' };
      current.setZoom(scale);
      return { ok: true, status: 'executed' };
    }
    const edits = current.editPeer();
    if (!peer || !edits) return fail('editor-unavailable');
    const done = (changed = true): XlsxCommandResult => ({
      ok: true,
      status: changed ? 'executed' : 'noop',
    });
    const attempt = (work: () => XlsxCommandResult): XlsxCommandResult => {
      try {
        return work();
      } catch (error) {
        if (inputRefusal(error)) current.refuse?.({ message: (error as Error).message });
        else current.fail(error);
        return fail('command-failed');
      }
    };
    const edit = (work: () => EditResult) =>
      attempt(() => {
        const result = work();
        markApplied?.();
        current.apply(result);
        return done(result.applied);
      });
    const bound = args as unknown as Record<string, never>;
    const selection = typeof env.selection === 'object' ? env.selection : null;
    const cells = () => {
      if (!selection) throw new Error('no cell selection');
      return { sheet: selection.sheet, range: rangeA1(peer, selection) };
    };
    const style = (patch: Parameters<WorkbookHandle['patchRangeStyle']>[2]) =>
      edit(() => {
        const { sheet, range } = cells();
        return edits.patchRangeStyle(sheet, range, patch);
      });
    const formatting = selection?.formatting ?? null;

    switch (id) {
      case 'bold':
        return style({ bold: formatting?.bold !== true });
      case 'italic':
        return style({ italic: formatting?.italic !== true });
      case 'strikethrough':
        return style({ strikethrough: formatting?.strikethrough !== true });
      case 'paintFormat':
        if (env.paintFormat) {
          current.setCapturedFormat(null, null);
          return done();
        }
        return attempt(() => {
          const { sheet, range } = cells();
          const source = `${sheet}:${selection!.top}:${selection!.left}:${selection!.bottom}:${selection!.right}`;
          current.setCapturedFormat(peer.captureFormat(sheet, range), source);
          return done();
        });
      case 'fontFamily':
        return style({ fontFamily: bound.family });
      case 'fontSize':
        return style({ fontSize: bound.points });
      case 'fontSizeStep': {
        const size = formatting?.fontSize ?? 10;
        const sizes = env.fontSizes;
        const next =
          bound.direction === 'increase'
            ? (sizes.find((entry) => entry > size) ?? size + 1)
            : ([...sizes].reverse().find((entry) => entry < size) ?? Math.max(1, size - 1));
        return style({ fontSize: next });
      }
      case 'textColor':
        return style({ textColor: bound.color });
      case 'fillColor':
        return style({ fillColor: bound.color });
      case 'numberFormat':
        return edit(() => {
          const { sheet, range } = cells();
          return bound.value === 'custom'
            ? edits.setNumberFormat(sheet, range, {
                type: 'custom',
                pattern: formatting?.numberFormatPattern ?? '0.00',
              })
            : edits.setNumberFormat(sheet, range, bound.value);
        });
      case 'decimalPlaces':
        return edit(() => {
          const { sheet, range } = cells();
          return edits.setNumberFormat(
            sheet,
            range,
            bound.direction === 'increase' ? 'increaseDecimal' : 'decreaseDecimal'
          );
        });
      case 'borderPreset':
        return style({
          border: {
            preset: bound.value,
            style: env.borderStyle ?? 'solid',
            color: env.borderColor ?? '#000000',
          },
        });
      case 'borderStyle':
        current.setBorderStyle(bound.value);
        return style({ border: { style: bound.value } });
      case 'borderColor':
        current.setBorderColor(bound.color);
        return style({ border: { color: bound.color } });
      case 'horizontalAlignment':
        return style({ horizontalAlignment: bound.value });
      case 'verticalAlignment':
        return style({ verticalAlignment: bound.value });
      case 'textWrapping':
        return style({ textWrapping: bound.value });
      case 'merge': {
        if (!selection) return fail('cell-selection-required');
        const ops = mergeOps(selection.sheet, selection, bound.value);
        if (ops.length === 0) return done(false);
        return edit(() => edits.applyOps(ops));
      }
      case 'undo':
        return edit(() => edits.undo());
      case 'redo':
        return edit(() => edits.redo());
      case 'save': {
        const generation = current.generation();
        markApplied?.();
        return edits.save().then((bytes) => {
          if (bridge().generation() !== generation) return fail('document-replaced');
          current.deliver(new Uint8Array(bytes));
          return done();
        }).catch((error) => {
          if (bridge().generation() !== generation) return fail('document-replaced');
          current.fail(error);
          return fail('command-failed');
        });
      }
      case 'exportPng':
        return fail('png-unavailable');
      case 'print': {
        const generation = current.generation();
        return current.afterPaint().then((painted) => {
          if (bridge().generation() !== generation) return fail('document-replaced');
          if (!painted) return fail('render-failed');
          window.print();
          return done();
        });
      }
      case 'proposalsPanel': {
        const open = args === null ? !env.proposalsPanelOpen : (bound.open as boolean);
        if (open === env.proposalsPanelOpen) return done(false);
        current.setProposalsPanelOpen(open);
        return done();
      }
      case 'proposalAccept':
        try {
          const result = edits.acceptProposal(bound.proposalId, { force: bound.force });
          markApplied?.();
          current.apply(result);
          current.markStale(bound.proposalId, null);
          current.refreshProposals();
          return done();
        } catch (error) {
          if (error instanceof StaleProposalError) {
            markApplied?.();
            current.markStale(bound.proposalId, error.cells);
            current.refreshProposals();
            return fail('proposal-stale');
          }
          current.fail(error);
          return fail('command-failed');
        }
      case 'proposalReject':
        return attempt(() => {
          const removed = edits.rejectProposal(bound.proposalId);
          markApplied?.();
          current.markStale(bound.proposalId, null);
          current.refreshProposals();
          return removed ? done() : fail('proposal-not-found');
        });
      default:
        return fail('unsupported-command');
    }
  };

  const targetOf = (id: XlsxCommandId): string => {
    if (!isCellCommand(id)) return 'document';
    const view = bridge().view();
    if (view.chartSelected) return 'chart';
    return view.selection ? `${view.sheet}:${JSON.stringify(normalizeRange(view.selection))}` : 'none';
  };

  const binding: XlsxCommandBinding = {
    requestSave: (save) => bridge().requestSave?.(save) ?? save(),
    environment,
    ordered: (id) => {
      const current = bridge();
      if (id === 'save') void current.coordinator.requestHydration('save').catch(() => {});
      return (
        !IMMEDIATE.has(id) ||
        current.coordinator.rejected.some((entry) => entry.generation === current.generation())
      );
    },
    refuse: (reason) => bridge().refuse?.(reason),
    admit: (operation, id) => {
      const readOnly = bridge().readOnly();
      return bridge().coordinator.runAfterPendingInput(async (_, applied) => {
        const current = bridge();
        const generation = current.generation();
        await current.preview();
        applied.check();
        if (bridge().generation() !== generation) throw new XlsxCommandAdmissionError('document-replaced');
        if (current.coordinator.error) throw current.coordinator.error;
        markApplied = applied;
        try {
          const result = await operation(readOnly);
          if (result !== null && typeof result === 'object' && 'ok' in result && result.ok === false) {
            const reason = 'failure' in result ? result.failure as { message: string } : { message: 'Command was refused' };
            applied.refuse(new WorkerInputRefusal(reason.message), id !== undefined && XLSX_COMMAND_DESCRIPTORS[id].mutatesDocument);
          }
          return result;
        } finally { markApplied = null; }
      }, { recover: id === undefined || XLSX_COMMAND_DESCRIPTORS[id].mutatesDocument, barrier: id === 'save' });
    },
    perform,
    capture(id) {
      const current = bridge();
      if (current.status() === 'empty') return null;
      return { generation: current.generation(), target: targetOf(id) };
    },
    resume(origin, id) {
      const current = bridge();
      if (current.generation() !== origin.generation) {
        return 'document-replaced';
      }
      return targetOf(id) === origin.target ? null : 'target-changed';
    },
    chrome: () => ({ i18n: bridge().i18n() }),
    focusEditor: () => bridge().focusGrid(),
  };
  return binding;
}

export function useWorkerXlsxCommands(
  controller: XlsxCommandController, bridge: WorkerXlsxEditorBridge
): void {
  const latest = useRef(bridge);
  latest.current = bridge;
  useLayoutEffect(() => {
    const binding = createWorkerXlsxCommandBinding(() => latest.current);
    controller.attach(binding);
    return () => controller.detach(binding);
  }, [controller]);
  useLayoutEffect(() => controller.refresh());
}

export interface PeerlessWorkerXlsxEditorBridge extends Omit<WorkerXlsxEditorBridge,
  'peer' | 'editPeer' | 'coordinator' | 'preview' | 'apply' | 'setCapturedFormat' | 'refuse'> {
  session(): (PeerlessWorkerCommandSession & Pick<
    WorkbookSession, 'editState' | 'selection'
  > & Partial<Pick<WorkbookSession, 'call' | 'pngExport'>>) | null;
  deliverPng?(bytes: Uint8Array): void;
  coordinator: PeerlessWorkerInputCoordinator;
  frame(): WorkbookCommandFrameRequest | undefined;
  selectionRevision(): number;
  formatToken(): string | null;
  formatRevision(): number;
  setFormatToken(token: string | null): void;
  selection?(): WorkbookCommandSelection | null;
  refuse?: XlsxCommandBinding['refuse'];
}

export interface PeerlessWorkerXlsxCommandBinding extends XlsxCommandBinding {
  subscribe(listener: () => void): () => void;
  refreshSubscription(): void;
  dispose(): void;
}

function peerlessRange(selection: WorkbookCommandSelection) {
  if (selection.range.kind === 'rowCol') {
    return { top: selection.range.start.row, left: selection.range.start.col,
      bottom: selection.range.end.row, right: selection.range.end.col };
  }
  const match = /^\$?([A-Za-z]+)\$?([1-9]\d*)(?::\$?([A-Za-z]+)\$?([1-9]\d*))?$/.exec(selection.range.a1);
  if (!match) return null;
  const column = (letters: string) => [...letters.toUpperCase()]
    .reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
  return { top: Number(match[2]) - 1, left: column(match[1]),
    bottom: Number(match[4] ?? match[2]) - 1, right: column(match[3] ?? match[1]) };
}

function peerlessA1(selection: XlsxSelectionEnvironment): string {
  const address = (row: number, col: number) => {
    let letters = '';
    for (let value = col + 1; value > 0; value = Math.floor((value - 1) / 26)) {
      letters = String.fromCharCode(65 + (value - 1) % 26) + letters;
    }
    return `${letters}${row + 1}`;
  };
  return `${address(selection.top, selection.left)}:${address(selection.bottom, selection.right)}`;
}

export function createPeerlessWorkerXlsxCommandBinding(
  bridge: () => PeerlessWorkerXlsxEditorBridge
): PeerlessWorkerXlsxCommandBinding {
  const listeners = new Set<() => void>();
  let watched: ReturnType<PeerlessWorkerXlsxEditorBridge['session']> = null;
  let watchedInput: PeerlessWorkerInputCoordinator | null = null;
  let identity = '';
  let pushed = '';
  let formatThroughSeq = 0;
  let offAck = () => {};
  let offInput = () => {};
  let admission: { env: XlsxCommandEnvironment; selection?: WorkbookCommandSelection;
    frame?: WorkbookCommandFrameRequest; selectionRevision?: number } | null = null;

  const selectionOf = (): WorkbookCommandSelection | undefined => {
    const current = bridge();
    const session = current.session();
    if (current.selection) return structuredClone(current.selection() ?? undefined);
    if (session?.selection) return structuredClone(session.selection);
    const view = current.view();
    const sheetId = session?.editState.sheets?.find((sheet) => sheet.index === view.sheet)?.id;
    if (!sheetId || !view.selection) return undefined;
    const range = normalizeRange(view.selection);
    return { sheetId, activeCell: { ...view.selection.focus }, range: { kind: 'rowCol',
      start: { row: range.top, col: range.left }, end: { row: range.bottom, col: range.right } } };
  };

  const sessionKey = () => {
    const session = bridge().session();
    return `${bridge().generation()}:${session?.commandIdentity.session}:${session?.commandIdentity.incarnation}:${session?.confirmedBasis.epoch}`;
  };

  const projection = () => {
    const session = bridge().session();
    const state = session?.editState;
    return JSON.stringify({
      formatting: state?.formatting, formattingAvailable: state?.formattingAvailable,
      selectionEditable: state?.selectionEditable, mergedRanges: state?.mergedRanges,
      canUndo: state?.canUndo, canRedo: state?.canRedo, readOnly: state?.readOnly,
      activeSheet: state?.activeSheet, sheets: state?.sheets, selection: session?.selection,
    });
  };
  const refreshSubscription = () => {
    const session = bridge().session();
    const coordinator = bridge().coordinator;
    const key = sessionKey();
    if (session === watched && coordinator === watchedInput && identity === key) return;
    offAck(); offInput();
    watched = session;
    watchedInput = coordinator;
    identity = key;
    formatThroughSeq = 0;
    pushed = projection();
    let pending = coordinator.pending;
    offInput = coordinator.subscribe(() => {
      if (bridge().coordinator !== coordinator || bridge().session() !== session || identity !== key || sessionKey() !== key) return;
      const next = coordinator.pending;
      if (next === pending) return;
      pending = next;
      for (const listener of listeners) listener();
    });
    offAck = session?.onAck((ack: WorkbookAck) => {
      if (bridge().session() !== session || identity !== key || sessionKey() !== key) return;
      if (!ack.selection && (!ack.editState || Object.keys(ack.editState).every((key) => key === 'calculation'))) return;
      const next = projection();
      if (next === pushed) return;
      pushed = next;
      for (const listener of listeners) listener();
    }) ?? (() => {});
  };
  const environment = (executing: boolean): XlsxCommandEnvironment => {
    refreshSubscription();
    if (executing && admission) return admission.env;
    const current = bridge();
    const view = current.view();
    const session = current.session();
    const state: Readonly<Partial<WorkbookEditState>> = session?.editState ?? {};
    const restored = current.selection ? current.selection() : session?.selection;
    const range = restored ? peerlessRange(restored) :
      current.selection ? null : view.selection ? normalizeRange(view.selection) : null;
    const sheet = restored ? state.sheets?.find((entry) => entry.id === restored.sheetId)?.index ?? view.sheet : view.sheet;
    return {
      status: current.status(), readOnly: current.readOnly() || current.coordinator.readOnly || state.readOnly === true,
      collaborative: current.collaborative(),
      selection: view.chartSelected ? 'chart' : range ? {
        sheet, ...range, merged: state.mergedRanges ?? [],
        formatting: state.formattingAvailable === false ? null : state.formatting ?? null,
      } : null,
      canUndo: current.coordinator.pending || (state.canUndo ?? false),
      canRedo: current.coordinator.pending || (state.canRedo ?? false),
      pendingInput: current.coordinator.draft !== null || current.coordinator.committed.length > 0,
      zoom: view.zoom, paintFormat: current.formatToken() !== null,
      borderStyle: view.borderStyle ?? state.formatting?.borderStyle ?? null,
      borderColor: view.borderColor ?? state.formatting?.borderColor ?? null,
      pngExport: session?.pngExport === true,
      proposals: view.proposalsAvailable ? view.proposals.map((proposal) => ({ id: proposal.id, label: proposal.agentId })) : null,
      proposalsPanelOpen: view.proposalsPanelOpen, fontFamilies: DEFAULT_FONT_FAMILIES,
      fontSizes: DEFAULT_FONT_SIZES, translate: current.translate,
      hostDisabled: state.selectionEditable === false ? new Set(
        (Object.keys(XLSX_COMMAND_DESCRIPTORS) as XlsxCommandId[]).filter(isCellCommand)
      ) : undefined,
    };
  };
  const targetOf = (id: XlsxCommandId, env = environment(false)) => {
    if (!isCellCommand(id)) return 'document';
    const selection = env.selection;
    if (selection === 'chart') return 'chart';
    return selection ? `${selection.sheet}:${selection.top}:${selection.left}:${selection.bottom}:${selection.right}` : 'none';
  };
  const perform = async <K extends XlsxCommandId>(
    id: K, args: XlsxCommandArgs[K], env: XlsxCommandEnvironment
  ): Promise<XlsxCommandResult> => {
    const current = bridge();
    const generation = current.generation();
    const session = current.session();
    const admitted = admission;
    const selection = admitted ? admitted.selection : selectionOf();
    const done = (changed = true): XlsxCommandResult => ({ ok: true, status: changed ? 'executed' : 'noop' });
    const fail = (code: XlsxCommandFailureCode): XlsxCommandResult => ({ ok: false, failure: commandReason(code, env) });
    if (id === 'zoom') {
      const scale = (args as XlsxCommandArgs['zoom']).scale;
      if (scale === env.zoom) return done(false);
      current.setZoom(scale);
      return done();
    }
    if (id === 'proposalsPanel') {
      const open = args === null ? !env.proposalsPanelOpen : (args as XlsxCommandArgs['proposalsPanel'])!.open;
      if (open === env.proposalsPanelOpen) return done(false);
      current.setProposalsPanelOpen(open);
      return done();
    }
    if (!session) return fail('editor-unavailable');
    if (id === 'paintFormat') {
      const armed = current.formatToken() !== null;
      current.setFormatToken(null);
      if (armed) return done();
    }
    const formatRevision = current.formatRevision();
    const cells = typeof env.selection === 'object' ? env.selection : null;
    const sheetId = cells ? session.editState.sheets?.find((sheet) => sheet.index === cells.sheet)?.id : undefined;
    const target: XlsxRangeTarget | undefined = cells && sheetId ? {
      sheetId, range: { kind: 'rowCol', start: { row: cells.top, col: cells.left },
        end: { row: cells.bottom, col: cells.right } },
    } : undefined;
    const argsRecord = args as unknown as Record<string, never>;
    const style = (patch: Parameters<WorkbookHandle['patchRangeStyle']>[2]): PeerlessWorkerCommandInput => {
      if (!cells || !target) throw new XlsxCommandAdmissionError('editor-unavailable');
      return { kind: 'patchRangeStyle', payload: {
        args: [cells.sheet, peerlessA1(cells), patch], target, selection,
      } };
    };
    const worker = async <T>(operation: () => Promise<T>): Promise<T> => {
      try { return await operation(); } catch (error) {
        if (bridge().generation() === generation && bridge().session() === session &&
          !(error instanceof PeerlessWorkerCommandRefusal) && !(error instanceof XlsxCommandAdmissionError) &&
          !current.coordinator.error) current.coordinator.fail(error);
        throw error;
      }
    };
    try {
      if (id === 'save') {
        const bytes = await worker(() => current.coordinator.save());
        if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
        current.deliver(bytes);
        return done();
      }
      if (id === 'print') {
        await worker(() => current.coordinator.whenPresented());
        if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
        window.print();
        return done();
      }
      if (id === 'exportPng') {
        const frame = admitted ? admitted.frame : current.frame();
        if (!session.pngExport || !session.call) return fail('png-unavailable');
        if (!frame || !current.deliverPng) return fail('editor-unavailable');
        const png = await session.call.renderPng(frame.viewport, { sheet: frame.sheet });
        if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
        current.deliverPng(png);
        return done();
      }
      let input: PeerlessWorkerCommandInput;
      switch (id) {
        case 'bold': case 'italic': case 'strikethrough': case 'fontSizeStep':
        case 'merge': case 'paintFormat':
          if (!target) return fail('cell-selection-required');
          input = { kind: 'applyOps', payload: { toolbar: { id, args, target }, selection } };
          break;
        case 'fontFamily': input = style({ fontFamily: argsRecord.family }); break;
        case 'fontSize': input = style({ fontSize: argsRecord.points }); break;
        case 'textColor': input = style({ textColor: argsRecord.color }); break;
        case 'fillColor': input = style({ fillColor: argsRecord.color }); break;
        case 'horizontalAlignment': input = style({ horizontalAlignment: argsRecord.value }); break;
        case 'verticalAlignment': input = style({ verticalAlignment: argsRecord.value }); break;
        case 'textWrapping': input = style({ textWrapping: argsRecord.value }); break;
        case 'borderPreset': input = style({ border: {
          preset: argsRecord.value, style: env.borderStyle ?? 'solid', color: env.borderColor ?? '#000000',
        } }); break;
        case 'borderStyle':
          current.setBorderStyle(argsRecord.value);
          input = style({ border: { style: argsRecord.value } }); break;
        case 'borderColor':
          current.setBorderColor(argsRecord.color);
          input = style({ border: { color: argsRecord.color } }); break;
        case 'numberFormat': case 'decimalPlaces':
          if (!cells || !target) return fail('cell-selection-required');
          if (id === 'numberFormat' && argsRecord.value === 'custom') {
            input = { kind: 'applyOps', payload: { toolbar: { id, args, target }, selection } };
          } else {
            input = { kind: 'setNumberFormat', payload: { args: [cells.sheet, peerlessA1(cells),
              id === 'decimalPlaces' ? argsRecord.direction === 'increase' ? 'increaseDecimal' : 'decreaseDecimal' : argsRecord.value,
            ], target, selection } };
          }
          break;
        case 'undo': case 'redo': input = { kind: id, payload: { args: [] } }; break;
        case 'proposalAccept': input = { kind: 'acceptProposal', payload: {
          args: [argsRecord.proposalId, { force: argsRecord.force }],
        } }; break;
        case 'proposalReject': input = { kind: 'rejectProposal', payload: { args: [argsRecord.proposalId] } }; break;
        default: return fail('unsupported-command');
      }
      const ticket = current.coordinator.command(input, { frame: admitted?.frame,
        selectionRevision: admitted?.selectionRevision });
      const status = await worker(() => ticket.result);
      if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
      if (id === 'paintFormat' && ticket.seq !== null && ticket.seq >= formatThroughSeq &&
        bridge().formatRevision() === formatRevision) {
        const result = status.result as WorkbookFormatTokenResult | undefined;
        if (typeof result?.formatToken !== 'string') {
          const error = new Error('Worker format capture omitted its token');
          current.coordinator.fail(error);
          throw error;
        }
        current.setFormatToken(result.formatToken);
        formatThroughSeq = ticket.seq;
      }
      if (id === 'proposalAccept' || id === 'proposalReject') {
        current.markStale(argsRecord.proposalId, null);
        current.refreshProposals();
      }
      if (id === 'proposalReject' && status.result === false) return fail('proposal-not-found');
      return done(status.status !== 'noop');
    } catch (error) {
      if (bridge().generation() !== generation || bridge().session() !== session) return fail('document-replaced');
      if (error instanceof PeerlessWorkerCommandRefusal) {
        if (id === 'proposalAccept' && error.status.reason === 'stale-proposal') {
          const result = error.status.result as { cells?: string[] } | undefined;
          current.markStale(argsRecord.proposalId, result?.cells ?? []);
          current.refreshProposals();
          return fail('proposal-stale');
        }
        if (error.status.reason === 'read-only') return fail('read-only');
        if (id === 'proposalReject' && error.status.reason === 'missing-target') return fail('proposal-not-found');
        return fail('input-failed');
      }
      if (error instanceof XlsxCommandAdmissionError) return fail(error.code);
      if (error instanceof Error && error.name === 'WorkbookPngUnavailableError') return fail('png-unavailable');
      return fail('command-failed');
    }
  };
  return {
    environment, ordered: (id) => !IMMEDIATE.has(id), perform,
    admit(operation, id) {
      const env = environment(false);
      const captured = { env, selection: selectionOf(), frame: structuredClone(bridge().frame()),
        selectionRevision: bridge().selectionRevision() };
      return bridge().coordinator.runAfterPendingInput(() => {
        const previous = admission;
        admission = captured;
        try { return operation(env.readOnly); }
        finally { admission = previous; }
      }, id === 'exportPng');
    },
    capture(id) {
      if (bridge().status() === 'empty') return null;
      return { generation: bridge().generation(), target: targetOf(id) };
    },
    resume(origin, id) {
      if (bridge().generation() !== origin.generation) return 'document-replaced';
      return targetOf(id, admission?.env) === origin.target ? null : 'target-changed';
    },
    refuse: (reason) => bridge().refuse?.(reason),
    chrome: () => ({ i18n: bridge().i18n() }), focusEditor: () => bridge().focusGrid(),
    subscribe(listener) {
      listeners.add(listener);
      refreshSubscription();
      return () => { listeners.delete(listener); };
    },
    refreshSubscription,
    dispose() { offAck(); offInput(); listeners.clear(); watched = null; watchedInput = null; identity = ''; },
  };
}

export function usePeerlessWorkerXlsxCommands(
  controller: XlsxCommandController, bridge: PeerlessWorkerXlsxEditorBridge
): void {
  const latest = useRef(bridge);
  latest.current = bridge;
  const binding = useRef<PeerlessWorkerXlsxCommandBinding | null>(null);
  useLayoutEffect(() => {
    const next = createPeerlessWorkerXlsxCommandBinding(() => latest.current);
    binding.current = next;
    controller.attach(next);
    const unsubscribe = next.subscribe(() => controller.refresh());
    return () => {
      unsubscribe(); next.dispose(); controller.detach(next);
      binding.current = null;
    };
  }, [controller]);
  useLayoutEffect(() => { binding.current?.refreshSubscription(); controller.refresh(); });
}
