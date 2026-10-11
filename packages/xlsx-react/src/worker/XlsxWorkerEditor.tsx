import {
  buildA11yGrid, cellAtPoint, cellRect, chartRegionAtPoint, createWorkbookRecoveryMutators,
  extendTo, fromTsv, hyperlinkAtCell, isProposalsAvailable, localNowSerial, moveFocus, normalizeRange,
  paintDisplayList, parseHyperlinkLocation, rangeRect, safeExternalHyperlink, selectionAt,
  selectionKeyReducer, toTsv, WorkbookEditPeerFailedError, WorkbookRecoveryRefusal,
  workbookEditPeerOperations, workbookSessionInternals,
} from '@betteroffice/xlsx';
import type {
  CapturedFormat, CellAddr, CellEdit, CellInputEdit, Direction, EditResult, Proposal,
  Selection, WorkbookFrame, WorkbookHandle, WorkbookReplayOp, WorkbookSheetView, XlsxEditResult,
  WorkbookCommandSelection, WorkbookSheetSummary,
} from '@betteroffice/xlsx';
import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, MouseEvent, SetStateAction, SyntheticEvent } from 'react';
import { flushSync } from 'react-dom';
import { createXlsxCommandController, XlsxCommandAdmissionError } from '../commands/createXlsxCommandStore';
import { commandForEvent } from '../commands/descriptors';
import { commandReason } from '../commands/evaluate';
import type { XlsxCommandFailureCode } from '../commands/types';
import type { InputDraft, InputSeal } from '../commands/inputCoordinator';
import { inPluginChrome, pluginEventStore } from '../commands/pluginEvents';
import { useCommandShortcuts } from '../commands/useCommandShortcuts';
import { useSaveRequest } from '../commands/useSaveRequest';
import { useWorkerXlsxCommands, usePeerlessWorkerXlsxCommands } from '../commands/useWorkerXlsxCommands';
import {
  createWorkerInputCoordinator, inputRefusal, WorkerInputRefusal, type WorkerInputCoordinatorHooks, type WorkerInputLease,
  createPeerlessWorkerInputCoordinator, PeerlessWorkerCommandRefusal, type PeerlessWorkerCommandInput,
  type PeerlessWorkerCommandTicket, type PeerlessWorkerInputEvent,
} from '../commands/workerInputCoordinator';
import { XlsxCommandContext } from '../commands/XlsxCommandProvider';
import { EditorToolbar } from '../components/EditorToolbar';
import { EditorChromeContext, PeerlessChromeContext } from '../components/EditorToolbarContext';
import type { BorderStyle } from '../components/Toolbar';
import { FormulaBarContext, type FormulaBarBinding } from '../components/toolbar/FormulaBar';
import { ToolbarCommandButton } from '../components/toolbar/ToolbarCommand';
import { LocaleProvider, useTranslation } from '../i18n';
import type { XlsxPluginEditorAccess, XlsxRevealAlignment } from '../plugins/createPluginClients';
import { createPeerlessPluginDocument } from '../plugins/createPeerlessPluginDocument';
import type { XlsxPluginDocumentAccess } from '../plugins/createPluginDocumentAccess';
import type { XlsxPluginNavigationFailureCode } from '../plugins/types';
import { PluginOverlays } from '../plugins/PluginOverlays';
import { PluginDock, useDockArea, type DockPlacement } from '../plugins/PluginPanels';
import { useXlsxPluginHost, type XlsxPluginHostBinding } from '../plugins/useXlsxPluginHost';
import { ProposalsPanel } from '../proposals/ProposalsPanel';
import { deriveLimits, expandRangeToMergedCells, positionAtPoint, scaledRect } from '../viewer/sessionGeometry';
import type { WorkerEditorApiBridge } from './createWorkerEditorApi';
import type { XlsxAsyncExportBridge } from './createAsyncEditorApi';
import {
  useEditableSessionWorkbook, EditableWorkbookSession, PeerlessWorkbookSession,
  type EditableSessionWorkbookProps, type PeerlessSessionWorkbookProps,
} from './useEditableSessionWorkbook';
import { PeerlessWorkerPaintSource, type WorkerPaintRequest, type WorkerPaintResult } from './PeerlessWorkerPaintSource';
import { WorkerPaintSource } from './WorkerPaintSource';
import { WorkerCellEcho } from './WorkerCellEcho';
import { WorkerDraftController, type WorkerNativeDraft } from './WorkerDraftController';
import { PeerlessWorkerInputs } from './PeerlessWorkerInputs';
import { WorkerPresentationStore } from './WorkerPresentationStore';
import { PeerlessSelectionEffect, PeerlessWorkerAccessibility, PeerlessWorkerExtent, PeerlessWorkerOverlay, PeerlessWorkerPresence } from './PeerlessWorkerSurface';

const BRAND = '#217346';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const visuallyHidden: CSSProperties = {
  position: 'absolute', width: 1, height: 1, margin: -1, padding: 0, border: 0,
  overflow: 'hidden', clip: 'rect(0 0 0 0)', whiteSpace: 'nowrap',
};
const chartKeys: Record<string, [number, number] | undefined> = {
  ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1],
};

type WorkerCellDraft = InputDraft & { prefill?: boolean; modified?: boolean; draftId?: number; revision?: number; unchanged?: boolean };
type PreviewPredecessor = { operation?: WorkbookReplayOp; input?: string; index: number };

function address(cell: CellAddr): string {
  let column = '';
  for (let col = cell.col + 1; col > 0; col = Math.floor((col - 1) / 26)) {
    column = String.fromCharCode(65 + (col - 1) % 26) + column;
  }
  return `${column}${cell.row + 1}`;
}

function commandSelection(target: WorkbookCommandSelection): Selection | null {
  let start: CellAddr;
  let end: CellAddr;
  if (target.range.kind === 'rowCol') ({ start, end } = target.range);
  else {
    const match = /^\$?([A-Z]+)\$?([1-9]\d*)(?::\$?([A-Z]+)\$?([1-9]\d*))?$/i.exec(target.range.a1);
    if (!match) return null;
    const cell = (letters: string, row: string): CellAddr => ({ row: Number(row) - 1,
      col: [...letters.toUpperCase()].reduce((col, letter) => col * 26 + letter.charCodeAt(0) - 64, 0) - 1 });
    start = cell(match[1], match[2]);
    end = cell(match[3] ?? match[1], match[4] ?? match[2]);
  }
  const focus = target.activeCell ?? end;
  return { focus: { ...focus }, anchor: {
    row: focus.row === start.row ? end.row : start.row,
    col: focus.col === start.col ? end.col : start.col,
  } };
}

function selectedRange(selection: Selection): string {
  const range = normalizeRange(selection);
  return `${address({ row: range.top, col: range.left })}:${address({ row: range.bottom, col: range.right })}`;
}

function fromPlugin(event: SyntheticEvent): boolean {
  return pluginEventStore(event.nativeEvent) !== null || inPluginChrome(event.target);
}

function useSyncedState<T>(initial: T) {
  const [value, setValue] = useState(initial);
  const ref = useRef(value);
  const set = useCallback((next: SetStateAction<T>) => {
    ref.current = typeof next === 'function' ? (next as (previous: T) => T)(ref.current) : next;
    setValue(ref.current);
  }, []);
  return [value, set, ref] as const;
}

function samePaint(painted: WorkerPaintResult | null, request: WorkerPaintRequest | null, sequence: number): boolean {
  if (!painted || !request || painted.sequence < sequence) return false;
  const previous = painted.request;
  return previous.generation === request.generation && previous.sheet === request.sheet &&
    previous.navigation === request.navigation && previous.zoom === request.zoom &&
    previous.dpr === request.dpr && previous.width === request.width && previous.height === request.height &&
    JSON.stringify(previous.viewport) === JSON.stringify(request.viewport);
}

function samePaintRequest(painted: WorkerPaintResult, request: WorkerPaintRequest): boolean {
  const previous = painted.request;
  return previous.generation === request.generation && previous.sheet === request.sheet &&
    previous.navigation === request.navigation && previous.zoom === request.zoom &&
    previous.dpr === request.dpr && previous.width === request.width && previous.height === request.height &&
    JSON.stringify(previous.viewport) === JSON.stringify(request.viewport);
}

function revealAxis(position: number, size: number, frozen: boolean, scroll: number, body: number,
  align: XlsxRevealAlignment): number | null {
  if (frozen) return null;
  if (align === 'start') return position;
  if (align === 'center') return Math.max(0, position - Math.max(0, body - size) / 2);
  if (position >= scroll && position + size <= scroll + body) return null;
  return position < scroll || size >= body ? position : Math.max(0, position + size - body);
}

type XlsxWorkerEditorProps = (EditableSessionWorkbookProps & { backend?: 'legacy' }) |
  (PeerlessSessionWorkbookProps & { backend: 'peerless' });

export function XlsxWorkerEditor(props: XlsxWorkerEditorProps) {
  const backend = useRef(props.backend ?? 'legacy').current;
  return backend === 'peerless' ? <LocaleProvider i18n={props.i18n}>
    <PeerlessXlsxWorkerEditor {...(props as PeerlessSessionWorkbookProps)} />
  </LocaleProvider> :
    <LegacyXlsxWorkerEditor {...(props as EditableSessionWorkbookProps)} />;
}

function LegacyXlsxWorkerEditor(props: EditableSessionWorkbookProps) {
  const { t } = useTranslation();
  const rootRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const editorInputRef = useRef<HTMLInputElement | null>(null);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);
  const editorRectRef = useRef<ReturnType<typeof cellRect>>(null);
  const runRef = useRef<EditableWorkbookSession | null>(null);
  const propsRef = useRef(props);
  propsRef.current = props;
  const [commands] = useState(createXlsxCommandController);
  const [active, setActive, activeRef] = useSyncedState(0);
  const [view, setView, viewRef] = useSyncedState<WorkbookSheetView | null>(null);
  const [selection, setSelection, selectionRef] = useSyncedState<Selection | null>(null);
  const [zoom, setZoom, zoomRef] = useSyncedState(1);
  const [editing, setEditing, editingRef] = useSyncedState<InputDraft | null>(null);
  const [formulaDraft, setFormulaDraft] = useState<string | null>(null);
  const [focusedCell, setFocusedCell] = useState<CellEdit | null>(null);
  const previewCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const [refusal, setRefusal] = useState<string | null>(null);
  const [previewDraft, setPreviewDraft] = useState<InputDraft | null>(null);
  const [painted, setPainted] = useState<WorkerPaintResult | null>(null);
  const [selectedChart, setSelectedChart, chartRef] = useSyncedState<{ id: string; movable: boolean } | null>(null);
  const [chartOffset, setChartOffset] = useState<{ x: number; y: number } | null>(null);
  const [capturedFormat, setCapturedFormat, capturedFormatRef] = useSyncedState<CapturedFormat | null>(null);
  const paintFormatSource = useRef<string | null>(null);
  const borderStyleRef = useRef<BorderStyle | undefined>(undefined);
  const borderColorRef = useRef<string | undefined>(undefined);
  const [proposals, setProposals, proposalsRef] = useSyncedState<Proposal[]>([]);
  const [proposalsPanelOpen, setProposalsPanelOpen, panelRef] = useSyncedState(false);
  const [staleFor, setStaleFor] = useState<Record<string, string[]>>({});
  const mutationRef = useRef(0);
  const navigationRef = useRef(0);
  const sourceRef = useRef<WorkerPaintSource | null>(null);
  const dragging = useRef(false);
  const clickStart = useRef<CellAddr | null>(null);
  const chartDrag = useRef<{ id: string; sheet: number; zoom: number; clientX: number; clientY: number } | null>(null);
  const nudge = useRef<{ id: string; sheet: number; dx: number; dy: number } | null>(null);
  const nudgeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const flushNudgeRef = useRef<(retained?: boolean) => void>(() => {});
  const suppressBlur = useRef(false);
  const suppressFormulaBlur = useRef(false);
  const composition = useRef<{ source: InputDraft['source']; done: Promise<boolean>; settle(ended: boolean): void } | null>(null);
  const paintWaiters = useRef<{ request: WorkerPaintRequest; sequence: number; resolve(value: boolean): void }[]>([]);
  const coordinatorRef = useRef<ReturnType<typeof createWorkerInputCoordinator> | null>(null);
  const draftRevision = useRef(0);
  const settledDraftRevision = useRef(0);
  const focus = useCallback(() => scrollRef.current?.focus({ preventScroll: true }), []);
  const reportInputError = useCallback((error: unknown) => {
    const local = inputRefusal(error);
    if (local || error instanceof XlsxCommandAdmissionError) {
      setRefusal((local ?? error as Error).message);
      return;
    }
    setRefusal(error instanceof Error ? error.message : String(error));
  }, []);
  const acceptInput = () => Boolean(runRef.current?.current && !runRef.current.retiring && !runRef.current.failure && !propsRef.current.readOnly);

  const apiBridgeRef = useRef<WorkerEditorApiBridge | null>(null);
  const [apiBridge] = useState<WorkerEditorApiBridge>(() => ({
    getPositionAtPoint: (x, y) => apiBridgeRef.current?.getPositionAtPoint?.(x, y) ?? null,
    coordinator: () => coordinatorRef.current,
    readOnly: () => propsRef.current.readOnly ?? false,
    clearSelection: () => apiBridgeRef.current?.clearSelection(),
    focus,
    refreshProposals: () => apiBridgeRef.current?.refreshProposals(),
    recoverInput: () => apiBridgeRef.current!.recoverInput(),
    selectCells: (...args) => apiBridgeRef.current?.selectCells(...args) ?? false,
    selectCellsAsync: (...args) => apiBridgeRef.current?.selectCellsAsync(...args) ?? Promise.resolve(false),
    apply: (result, op) => apiBridgeRef.current?.apply(result, op),
    previewEdits: (sheet, edits, op) => apiBridgeRef.current?.previewEdits?.(sheet, edits, op) ?? Promise.resolve(),
    canNavigateSync: () => apiBridgeRef.current?.canNavigateSync?.() ?? false,
  }));
  const { run, error, loading } = useEditableSessionWorkbook(props, commands.store, apiBridge);
  runRef.current = run;
  const inputHooks = useMemo(() => ({ current: null as WorkerInputCoordinatorHooks | null }), [run]);
  const acceptedCells = useMemo(() => new Map<string, string>(), [run]);
  const draftOperations = useMemo(() => new WeakMap<InputDraft, WorkbookReplayOp>(), [run]);
  const previewPredecessors = useMemo(() => new Map<WorkbookReplayOp, Map<string, PreviewPredecessor>>(), [run]);
  const draftDiscards = useMemo(() => new WeakMap<InputDraft, () => void>(), [run]);
  const operationPeer = useMemo(() => ({ current: run?.editPeer ?? null }), [run]);
  if (run?.editPeer && !run.recovering) operationPeer.current = run.editPeer;
  const previewOps = useMemo(() => new Map<string, WorkbookReplayOp>(), [run]);
  const unchangedPreviews = useMemo(() => new WeakSet<WorkbookReplayOp>(), [run]);
  const previewJournal = useMemo(() => [] as WorkbookReplayOp[], [run]);
  const completedPreviews = useMemo(() => new WeakMap<WorkbookReplayOp, number>(), [run]);
  const unpreviewedOps = useMemo(() => new Set<WorkbookReplayOp>(), [run]);
  const previewState = useMemo(() => ({ revision: 0, sequence: Infinity, request: null as WorkerPaintRequest | null, committed: null as InputDraft | null }), [run]);
  const retirement = useMemo(() => {
    const callbacks = new Set<() => void>();
    return {
      resolve() { for (const callback of callbacks) callback(); callbacks.clear(); },
      wait(promise: Promise<void>) {
        return new Promise<void>((resolve, reject) => {
          const complete = () => { callbacks.delete(complete); resolve(); };
          callbacks.add(complete);
          void promise.then(complete, (error) => { callbacks.delete(complete); reject(error); });
        });
      },
    };
  }, [run]);
  const coordinator = useMemo(() => {
    const owner = run;
    const input = createWorkerInputCoordinator({
      generation: () => owner?.generation ?? 0,
      capture: () => ({ sheet: activeRef.current, target: JSON.stringify(selectionRef.current) }),
      isReady: () => owner?.ready ?? false,
      isAcknowledged: () => !owner?.editPeer || owner.editPeer.acknowledgedSequence >= owner.editPeer.sentSequence,
      whenReady: () => owner?.whenHydrated() ?? Promise.reject(new XlsxCommandAdmissionError('editor-unavailable')),
      requestHydration: (reason) => owner?.requestHydration(reason),
      flushEdits: () => owner?.editPeer?.flush() ?? Promise.reject(new XlsxCommandAdmissionError('editor-unavailable')),
      async acknowledgeEdits() {
        if (!owner?.editPeer || owner.recovering || owner.retiring) return;
        await retirement.wait(workbookEditPeerOperations(owner.editPeer).whenAcknowledged());
        if (!owner.retiring) await inputHooks.current?.acknowledgeEdits?.();
      },
      preview: (draft) => inputHooks.current!.preview(draft),
      resolveDraft: (draft) => inputHooks.current!.resolveDraft!(draft),
      sameDraft: (a, b) => (a as WorkerCellDraft).draftId === (b as WorkerCellDraft).draftId,
      seal: () => inputHooks.current?.seal() ?? {},
      sync: () => inputHooks.current?.sync(),
      write: (draft, applied) => inputHooks.current?.write(draft, applied) ?? false,
      onError: (error) => owner?.fail(error),
      onRefusal(error, draft) {
        if (owner?.retiring) { owner.reportRefusal(error); return; }
        const message = error instanceof Error ? error.message : String(error);
        if (owner?.recovering) flushSync(() => setRefusal(message));
        else setRefusal(message);
        inputHooks.current?.sync?.();
        const revision = (draft as WorkerCellDraft | undefined)?.revision ?? 0;
        if (draft && !propsRef.current.readOnly && revision > settledDraftRevision.current && revision === draftRevision.current) {
          input.setDraft(draft);
          if (draft.source === 'cell') setEditing(draft);
          else setFormulaDraft(draft.value);
        }
      },
    });
    return input;
  }, [run, inputHooks, reportInputError, retirement]);
  coordinatorRef.current = coordinator;

  const capture = useCallback((): WorkerPaintRequest | null => {
    const owner = runRef.current;
    const scroll = scrollRef.current;
    const zoom = zoomRef.current;
    if (!owner?.current || !viewRef.current || !scroll || !canvasRef.current ||
      !scroll.clientWidth || !scroll.clientHeight) return null;
    return {
      generation: owner.generation, sheet: activeRef.current, navigation: navigationRef.current,
      zoom, dpr: window.devicePixelRatio || 1, width: scroll.clientWidth, height: scroll.clientHeight,
      viewport: { x: scroll.scrollLeft / zoom, y: scroll.scrollTop / zoom,
        width: scroll.clientWidth / zoom, height: scroll.clientHeight / zoom },
    };
  }, []);

  const boundary = useCallback((): Promise<void> => {
    if (!run?.current) return Promise.reject(new XlsxCommandAdmissionError('document-replaced'));
    if (run.failure && !run.recovering) return Promise.reject(run.failure);
    return Promise.resolve();
  }, [run]);

  const relinkPredecessors = useCallback((key: string, op: WorkbookReplayOp, previous?: PreviewPredecessor) => {
    for (const predecessors of previewPredecessors.values()) {
      const predecessor = predecessors.get(key);
      if (predecessor?.operation === op) {
        predecessor.operation = previous?.operation;
        predecessor.input = previous?.input;
        predecessor.index = previous?.index ?? -1;
      }
    }
    const predecessors = previewPredecessors.get(op);
    predecessors?.delete(key);
    if (!predecessors?.size) previewPredecessors.delete(op);
  }, [previewPredecessors]);

  const preview = useCallback(async (draft?: InputDraft, discardOnFailure = true): Promise<void> => {
    if (!run || run.failure || run.recovering || run.retiring) return;
    if (draft) {
      const key = `${draft.sheet}:${draft.row}:${draft.col}`;
      const preceding = previewOps.get(key);
      const committed = previewState.committed;
      const precedingInput = acceptedCells.get(key) ?? (committed &&
        `${committed.sheet}:${committed.row}:${committed.col}` === key ? committed.value : undefined);
      acceptedCells.set(key, draft.value);
      const op: WorkbookReplayOp = draftOperations.get(draft) ?? {
        method: 'editCell', args: [draft.sheet, draft.row, draft.col, draft.value], calculation: {
          nowSerial: localNowSerial(Date.now()),
          randSeed: globalThis.crypto.getRandomValues(new Uint32Array(1))[0] >>> 0,
        },
      };
      draftOperations.set(draft, op);
      if (preceding !== op) {
        let predecessors = previewPredecessors.get(op);
        if (!predecessors) previewPredecessors.set(op, predecessors = new Map());
        predecessors.set(key, {
          operation: preceding, input: precedingInput, index: preceding ? previewJournal.indexOf(preceding) : -1,
        });
      }
      if ((draft as WorkerCellDraft).unchanged) unchangedPreviews.add(op);
      if (preceding && preceding !== op && unchangedPreviews.has(preceding)) {
        const index = previewJournal.indexOf(preceding);
        if (index >= 0) previewJournal.splice(index, 1);
      }
      previewOps.set(key, op);
      if (!previewJournal.includes(op)) previewJournal.push(op);
      previewState.sequence = Infinity;
      const revision = ++previewState.revision;
      const discard = () => {
        const previous = previewPredecessors.get(op)?.get(key);
        const index = previewJournal.indexOf(op);
        if (index >= 0) previewJournal.splice(index, 1);
        relinkPredecessors(key, op, previous);
        const current = previewOps.get(key) === op;
        if (current) {
          if (previous?.operation) {
            previewOps.set(key, previous.operation);
            if (!previewJournal.includes(previous.operation)) {
              previewJournal.splice(Math.max(0, previous.index), 0, previous.operation);
            }
          } else previewOps.delete(key);
          if (previous?.operation && previous.input !== undefined) acceptedCells.set(key, previous.input);
          else acceptedCells.delete(key);
        }
        if (!previewOps.size && ((current && !unpreviewedOps.size) || revision === previewState.revision) && !run.retiring) {
          setPreviewDraft(previewState.committed);
          previewState.request = null;
        } else if (current && !run.retiring) {
          const pending = [...previewOps.keys()].reverse()[0];
          if (pending) {
            const [sheet, row, col] = pending.split(':').map(Number);
            setPreviewDraft({ ...draft, sheet, row, col, value: acceptedCells.get(pending) ?? '' });
          }
        }
        const complete = !unpreviewedOps.size && previewJournal.every((pending) => completedPreviews.has(pending));
        previewState.sequence = complete ? Math.max(0, ...previewJournal.map((pending) => completedPreviews.get(pending)!)) : Infinity;
        if (complete) sourceRef.current?.schedule();
      };
      if (discardOnFailure) draftDiscards.set(draft, discard);
      if (run.ready) {
        if (!unpreviewedOps.size) {
          flushSync(() => setPreviewDraft(draft));
        }
        return;
      }
      const request = capture();
      if (unpreviewedOps.size) { await boundary(); return; }
      if (request) {
        const internal = workbookSessionInternals.get(run.session);
        if (!internal?.preview) throw new Error('Worker cell preview is unavailable');
        let frame: WorkbookFrame;
        try {
          frame = await internal.preview(request.viewport, request.sheet,
            previewJournal.filter((pending) => !completedPreviews.has(pending)));
        }
        catch (error) {
          if (discardOnFailure) discard();
          throw error;
        }
        if (run.retiring || run.recovering || run.failure || !run.current || revision !== previewState.revision) return;
        previewState.request = request;
        previewState.sequence = previewJournal.every((pending) => completedPreviews.has(pending)) ?
          run.editPeer?.sentSequence ?? Infinity : Infinity;
        if (unpreviewedOps.size) return;
        flushSync(() => setPreviewDraft(draft));
        if (run.ready) return;
        const canvas = previewCanvasRef.current;
        const context = canvas?.getContext('2d');
        if (!canvas || !context) throw new Error('Preview canvas context is unavailable');
        canvas.width = Math.round(request.width * request.dpr);
        canvas.height = Math.round(request.height * request.dpr);
        canvas.style.width = `${request.width}px`;
        canvas.style.height = `${request.height}px`;
        paintDisplayList(context, frame.displayList, request.dpr * request.zoom);
        canvas.dataset.previewReady = 'true';
      }
    }
    await boundary();
  }, [run, capture, boundary, acceptedCells, previewOps, previewState, draftOperations, draftDiscards, previewPredecessors, relinkPredecessors, completedPreviews, unpreviewedOps, previewJournal, unchangedPreviews]);

  const previewEdits = useCallback(async (sheet: number, edits: readonly CellInputEdit[], operation?: WorkbookReplayOp) => {
    if (operation?.method === 'applyEdits' && operation.args[0].steps.some((step) => step.op !== 'setCellInputs')) {
      unpreviewedOps.add(operation);
      previewState.revision += 1;
      previewState.request = null;
      setPreviewDraft(null);
      return async () => { unpreviewedOps.delete(operation); };
    }
    if (operation) operation.calculation ??= {
      nowSerial: localNowSerial(Date.now()),
      randSeed: globalThis.crypto.getRandomValues(new Uint32Array(1))[0] >>> 0,
    };
    const preceding = new Map<string, PreviewPredecessor>();
    const recorded = new Map<string, WorkbookReplayOp>();
    for (const edit of edits) {
      if (!Number.isInteger(sheet) || sheet < 0 || sheet >= (run?.session.state.sheets.length ?? 0) ||
        !Number.isInteger(edit.row) || edit.row < 0 || edit.row > 1048575 ||
        !Number.isInteger(edit.col) || edit.col < 0 || edit.col > 16383 ||
        !edit.input.startsWith('=') && [...edit.input.replace(/^'/, '')].length > 32767) {
        throw new WorkerInputRefusal('Invalid cell input or target');
      }
    }
    for (const edit of edits) {
      const key = `${sheet}:${edit.row}:${edit.col}`;
      const op = operation ?? { method: 'editCell' as const, args: [sheet, edit.row, edit.col, edit.input] as [number, number, number, string] };
      const previous = previewOps.get(key);
      let predecessors = previewPredecessors.get(op);
      if (!predecessors) previewPredecessors.set(op, predecessors = new Map());
      const predecessor = preceding.get(key) ?? (previous === op ? predecessors.get(key) : undefined) ?? {
        operation: previous, input: acceptedCells.get(key) ?? (previewState.committed &&
          `${previewState.committed.sheet}:${previewState.committed.row}:${previewState.committed.col}` === key ?
          previewState.committed.value : undefined), index: previous ? previewJournal.indexOf(previous) : -1,
      };
      preceding.set(key, predecessor);
      predecessors.set(key, predecessor);
      acceptedCells.set(key, edit.input);
      if (previous && previous !== op && unchangedPreviews.has(previous)) {
        const index = previewJournal.indexOf(previous);
        if (index >= 0) previewJournal.splice(index, 1);
      }
      previewOps.set(key, op);
      recorded.set(key, op);
      if (!previewJournal.includes(op)) previewJournal.push(op);
    }
    const restore = async () => {
      let changed = false;
      for (const edit of edits) {
        const key = `${sheet}:${edit.row}:${edit.col}`;
        const op = recorded.get(key);
        const predecessor = preceding.get(key);
        const previous = predecessor?.operation;
        const value = predecessor?.input;
        if (op) relinkPredecessors(key, op, predecessor);
        if (previewOps.get(key) !== op) continue;
        if (previous) {
          previewOps.set(key, previous);
          if (!previewJournal.includes(previous)) previewJournal.splice(Math.max(0, predecessor!.index), 0, previous);
        }
        else previewOps.delete(key);
        if (previous && value !== undefined) acceptedCells.set(key, value);
        else acceptedCells.delete(key);
        changed = true;
      }
      for (const op of new Set(recorded.values())) {
        const index = previewJournal.indexOf(op);
        if (index >= 0) previewJournal.splice(index, 1);
      }
      if (!changed || !run || run.retiring || run.failure) return;
      const pending = [...previewOps.entries()].reverse()[0];
      if (pending) {
        const [sheet, row, col] = pending[0].split(':').map(Number);
        const draft: InputDraft = { generation: run.generation, sheet, row, col,
          value: acceptedCells.get(pending[0]) ?? '', source: 'cell' };
        draftOperations.set(draft, pending[1]);
        await preview(draft);
        if (previewJournal.every((op) => completedPreviews.has(op))) {
          previewState.sequence = run.editPeer?.sentSequence ?? Infinity;
          sourceRef.current?.schedule();
        }
      } else {
        previewState.revision += 1;
        previewState.request = null;
        previewState.sequence = run.editPeer?.sentSequence ?? Infinity;
        setPreviewDraft(previewState.committed);
      }
    };
    try {
      const draft: InputDraft | undefined = edits[0] ? { generation: run!.generation, sheet, row: edits[0].row,
        col: edits[0].col, value: edits[0].input, source: 'cell' } : undefined;
      if (draft) draftOperations.set(draft, recorded.get(`${sheet}:${draft.row}:${draft.col}`)!);
      const prepared = draft ? preview(draft, false) : Promise.resolve();
      await prepared;
    } catch (error) {
      await restore();
      throw error;
    }
    return restore;
  }, [run, preview, relinkPredecessors, acceptedCells, previewOps, previewPredecessors, previewState, draftOperations, completedPreviews, unpreviewedOps, previewJournal, unchangedPreviews]);

  const refreshProposals = useCallback(() => {
    const owner = runRef.current;
    if (!owner?.peer) { setProposals([]); return; }
    setProposals(owner.peer.listProposals());
    if (!owner.recovering) sourceRef.current?.schedule();
    commands.refresh();
  }, [commands]);

  const peerPaint = useCallback(() => {
    if (run?.retiring || run?.recovering) return;
    sourceRef.current?.schedule();
  }, [run]);

  const trackCommit = useCallback((op?: WorkbookReplayOp) => {
    if (op) {
      completedPreviews.set(op, run?.editPeer?.sentSequence ?? 0);
      unpreviewedOps.delete(op);
    }
    const complete = !unpreviewedOps.size && previewJournal.every((pending) => completedPreviews.has(pending));
    previewState.sequence = complete ? Math.max(0, ...previewJournal.map((pending) => completedPreviews.get(pending)!)) : Infinity;
    if (!run?.recovering && !run?.retiring) sourceRef.current?.schedule();
  }, [run, previewState, completedPreviews, unpreviewedOps, previewJournal]);

  const closeWrittenDraft = (draft: InputDraft) => {
    const live = coordinator.draft;
    if (live && (live.source !== draft.source || live.sheet !== draft.sheet || live.row !== draft.row ||
      live.col !== draft.col || (live as WorkerCellDraft).revision !== (draft as WorkerCellDraft).revision ||
      live.value !== draft.value && (live as WorkerCellDraft).modified !== false)) return;
    if (live) coordinator.setDraft(null);
    if (draft.source === 'cell') {
      const current = editingRef.current;
      if (current && current.sheet === draft.sheet && current.row === draft.row && current.col === draft.col &&
        (current as WorkerCellDraft).revision === (draft as WorkerCellDraft).revision &&
        (current.value === draft.value || (current as WorkerCellDraft).modified === false)) {
        suppressBlur.current = true;
        setEditing(null);
      }
    } else setFormulaDraft(null);
  };

  const apply = useCallback((result: EditResult | XlsxEditResult, op?: WorkbookReplayOp) => {
    const owner = run;
    if (!owner?.peer || owner.retiring) return;
    try {
      if (!op && previewState.request && (!('ok' in result) || result.ok) && result.applied) {
        previewState.revision += 1;
        previewState.request = null;
        setPreviewDraft(null);
      }
      trackCommit(op);
      mutationRef.current += 1;
      const info = owner.peer.sheetInfo();
      if (viewRef.current && info.activeSheet === activeRef.current) {
        setView({ ...viewRef.current, ...info, sheet: info.activeSheet, version: owner.peer.version() });
      }
      peerPaint();
      refreshProposals();
      if ((!('ok' in result) || result.ok) && result.applied) propsRef.current.onChange?.();
      commands.refresh();
    } catch (error) {
      owner.fail(error);
    }
  }, [run, commands, peerPaint, refreshProposals, trackCommit, previewState]);

  const afterPaint = useCallback((): Promise<boolean> => {
    const owner = runRef.current;
    const request = capture();
    const source = sourceRef.current;
    if (!owner?.current || owner.failure || !request || !source) return Promise.resolve(false);
    const sequence = owner.editPeer?.sentSequence ?? 0;
    if (samePaint(source.painted, request, sequence)) return Promise.resolve(true);
    source.schedule();
    return new Promise((resolve) => {
      const waiter = { request, sequence, resolve(value: boolean) { clearTimeout(timer); resolve(value); } };
      const timer = setTimeout(() => {
        paintWaiters.current = paintWaiters.current.filter((entry) => entry !== waiter);
        resolve(false);
      }, 1000);
      paintWaiters.current.push(waiter);
    });
  }, [capture]);

  const place = useCallback((sheet: number, next: Selection | null): boolean => {
    const owner = run;
    if (!owner?.peer || !owner.current || owner.retiring) return false;
    const info = owner.peer.sheetInfo();
    if (info.activeSheet !== sheet) return false;
    navigationRef.current += 1;
    setActive(sheet);
    setView({ ...viewRef.current!, ...info, sheet, version: owner.peer.version(), frozenWidth: 0, frozenHeight: 0 });
    setSelection(next ? structuredClone(next) : null);
    setSelectedChart(null);
    setCapturedFormat(null);
    paintFormatSource.current = null;
    if (!coordinator.draft) {
      setEditing(null);
      setFormulaDraft(null);
    }
    const scroll = scrollRef.current;
    if (scroll && next) {
      const at = owner.peer.cellPosition(sheet, next.focus.row, next.focus.col);
      scroll.scrollLeft = at.x * zoomRef.current;
      scroll.scrollTop = at.y * zoomRef.current;
    }
    if (!owner.recovering) { peerPaint(); sourceRef.current?.schedule(); }
    commands.refresh();
    return true;
  }, [commands, coordinator, peerPaint]);

  const recoverInput = useCallback(async () => {
    const owner = runRef.current;
    if (!owner?.current || !owner.peer || !owner.editPeer) throw new XlsxCommandAdmissionError('editor-unavailable');
    const edits = owner.editPeer;
    operationPeer.current = edits;
    let recoveryMutators: ReturnType<typeof createWorkbookRecoveryMutators>;
    try { recoveryMutators = createWorkbookRecoveryMutators(edits); }
    catch (error) {
      if (error instanceof WorkbookEditPeerFailedError) throw error;
      throw new WorkbookEditPeerFailedError(error instanceof Error ? error : new Error(String(error)));
    }
    owner.recovering = true;
    owner.editPeer = new Proxy(edits, {
      get(target, key) {
        if (key in recoveryMutators) return Reflect.get(recoveryMutators, key);
        if (key === 'flush' || key === 'save') return () => Promise.reject(owner.failure ?? new Error('Worker is unavailable'));
        return Reflect.get(target, key);
      },
    });
    try {
      flushNudgeRef.current(true);
      if (coordinator.draft) coordinator.submit(coordinator.draft);
      await coordinator.recover();
    }
    catch (error) {
      if (error instanceof XlsxCommandAdmissionError || error instanceof WorkbookEditPeerFailedError) throw error;
      throw new WorkbookEditPeerFailedError(error instanceof Error ? error : new Error(String(error)));
    }
    finally {
      owner.editPeer = edits;
      owner.recovering = false;
      coordinator.fail(owner.failure ?? edits.error ?? new Error('Worker is unavailable'));
    }
  }, [coordinator]);

  apiBridgeRef.current = {
    getPositionAtPoint: (x, y) => {
      const painted = sourceRef.current?.painted;
      if (!painted || painted.request.generation !== runRef.current?.generation) return null;
      return positionAtPoint(painted.displayList, painted.request.sheet, canvasRef.current, x, y);
    },
    coordinator: () => coordinator, readOnly: () => propsRef.current.readOnly ?? false, focus,
    clearSelection: () => {
      coordinator.settle(); setEditing(null); setFormulaDraft(null); setSelection(null); setSelectedChart(null);
      setCapturedFormat(null); commands.refresh();
    },
    refreshProposals, recoverInput, selectCells: place, previewEdits,
    canNavigateSync: () => previewState.request === null,
    selectCellsAsync: async (sheet, next) => place(sheet, next), apply,
  };

  const pluginAccessRef = useRef<XlsxPluginEditorAccess | null>(null);
  const [pluginAccess] = useState<XlsxPluginEditorAccess>(() => ({
    handle: () => runRef.current?.ready && !runRef.current.failure ? runRef.current.peer : null,
    admit: (handle, operation) => pluginAccessRef.current!.admit(handle, operation),
    readOnlyRefusal: (handle) => pluginAccessRef.current!.readOnlyRefusal(handle),
    applyEdits: (...args) => pluginAccessRef.current!.applyEdits(...args),
    admitEdits: (...args) => pluginAccessRef.current!.admitEdits!(...args),
    commands: () => commands,
    navigator: {
      selectCells: (...args) => pluginAccessRef.current!.navigator.selectCells(...args),
      scrollToCell: (...args) => pluginAccessRef.current!.navigator.scrollToCell(...args),
    },
  }));
  const pluginSelection = useMemo(() => {
    const sheetId = run?.session.state.sheets[active]?.id;
    return sheetId === undefined ? null : { sheetId, sheetIndex: active,
      cells: selection, chartId: selectedChart?.id ?? null };
  }, [run, active, selection, selectedChart]);
  const pluginHost = useXlsxPluginHost({ ...props, access: pluginAccess, commands,
    readOnly: props.readOnly ?? false, handle: run?.ready && !run.failure ? run.peer : null,
    selection: pluginSelection, canvasRef, t });
  const presentGridRef = useRef(pluginHost.presentGrid);
  presentGridRef.current = pluginHost.presentGrid;
  const [dockAreaRef, dockArea] = useDockArea(pluginHost.managed);

  const cancelPendingFrames = useCallback(() => {
    for (const waiter of paintWaiters.current.splice(0)) waiter.resolve(false);
  }, []);

  useLayoutEffect(() => {
    setRefusal(null);
    setPainted(null); setPreviewDraft(null); setSelection(null); setView(null); setEditing(null);
    setFormulaDraft(null); setSelectedChart(null); setChartOffset(null); setFocusedCell(null);
    editorRectRef.current = null;
    setCapturedFormat(null); setProposals([]); setStaleFor({}); setProposalsPanelOpen(false);
    mutationRef.current = 0;
    navigationRef.current += 1;
    dragging.current = false;
    chartDrag.current = null;
    nudge.current = null;
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    nudgeTimer.current = null;
    composition.current?.settle(false);
    composition.current = null;
    pluginHost.beginLoad();
    if (!run) return;
    const source = new WorkerPaintSource({
      generation: run.generation, capture, isCurrent: () => run.current && !run.failure,
      requestFrame: (request) => run.session.call.frame(request.viewport, { sheet: request.sheet }),
      sentSequence: () => run.editPeer?.sentSequence ?? 0,
      publish(painted) {
        const canvas = canvasRef.current;
        const context = canvas?.getContext('2d');
        if (!canvas || !context) throw new Error('Canvas context is unavailable');
        const request = painted.request;
        canvas.width = Math.round(request.width * request.dpr);
        canvas.height = Math.round(request.height * request.dpr);
        canvas.style.width = `${request.width}px`;
        canvas.style.height = `${request.height}px`;
        paintDisplayList(context, painted.displayList, request.dpr * request.zoom);
        canvas.dataset.workerSequence = String(painted.sequence);
        canvas.dataset.paintSource = painted.source;
        canvas.dataset.workerZoom = String(request.zoom);
        setPainted(painted);
        if (painted.sequence >= previewState.sequence) {
          previewOps.clear();
          previewJournal.length = 0;
          previewPredecessors.clear();
          acceptedCells.clear();
          previewState.request = null;
          previewState.committed = null;
          flushSync(() => setPreviewDraft(null));
        }
        const sheetId = run.session.state.sheets[request.sheet]?.id;
        presentGridRef.current(sheetId === undefined ? null : {
          frame: painted.displayList, zoom: request.zoom, viewport: request.viewport, version: painted.version, sheetId,
        });
        for (const waiter of paintWaiters.current.slice()) {
          const covered = samePaint(painted, waiter.request, waiter.sequence);
          if (covered || !samePaint(painted, waiter.request, 0)) {
            paintWaiters.current = paintWaiters.current.filter((entry) => entry !== waiter);
            waiter.resolve(covered);
          }
        }
        run.firstPaint();
      },
      onError: (error) => run.fail(error),
    });
    sourceRef.current = source;
    const sheet = run.session.state.activeSheet;
    setActive(sheet);
    void run.session.call.sheetView(sheet).then((next) => {
      if (!run.current || run.failure || sourceRef.current !== source) return;
      setView(next);
      setSelection(selectionAt({ row: 0, col: 0 }));
      if (scrollRef.current) {
        scrollRef.current.scrollLeft = next.initialScrollX * zoomRef.current;
        scrollRef.current.scrollTop = next.initialScrollY * zoomRef.current;
      }
      source.schedule();
    }).catch((error) => run.fail(error));
    return () => {
      source.dispose();
      if (sourceRef.current === source) sourceRef.current = null;
      cancelPendingFrames();
      presentGridRef.current(null);
    };
  }, [run, capture, cancelPendingFrames]);

  useEffect(() => {
    const scroll = scrollRef.current;
    if (!scroll || !run) return;
    const schedule = () => sourceRef.current?.schedule();
    scroll.addEventListener('scroll', schedule, { passive: true });
    const observer = new ResizeObserver(schedule);
    observer.observe(scroll);
    schedule();
    return () => { scroll.removeEventListener('scroll', schedule); observer.disconnect(); };
  }, [run, zoom]);

  useEffect(() => {
    if (typeof window.matchMedia !== 'function') return;
    let media: MediaQueryList | undefined;
    const update = () => {
      sourceRef.current?.schedule();
      media?.removeEventListener('change', update);
      media = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`);
      media.addEventListener('change', update);
    };
    update();
    return () => media?.removeEventListener('change', update);
  }, []);

  useEffect(() => {
    if (!run?.failure) return;
    sourceRef.current?.dispose();
    cancelPendingFrames();
    chartDrag.current = null;
    flushNudgeRef.current(true);
    syncDraft();
    composition.current?.settle(true);
    composition.current = null;
  }, [run, run?.failure, cancelPendingFrames]);

  useEffect(() => () => {
    cancelPendingFrames();
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    composition.current?.settle(false);
    composition.current = null;
  }, [cancelPendingFrames]);

  useEffect(() => {
    setFocusedCell(null);
    if (!run?.current || run.failure || !selection || selectedChart) return;
    if (run.ready && run.peer) { setFocusedCell(run.peer.cell(active, selection.focus.row, selection.focus.col)); return; }
    let cancelled = false;
    const navigation = navigationRef.current;
    void run.session.call.cellInputs(active, address(selection.focus)).then((result) => {
      if (!cancelled && run.current && navigationRef.current === navigation) setFocusedCell(result.cells[0]?.[0] ?? null);
    }).catch((error) => { if (!cancelled && run.current) run.fail(error); });
    return () => { cancelled = true; };
  }, [run, run?.peer, selection, active, selectedChart, painted]);

  const limits = () => deriveLimits(sourceRef.current?.painted?.displayList ?? null, viewRef.current!,
    (scrollRef.current?.clientHeight ?? 0) / zoomRef.current);
  const draftFor = (source: InputDraft['source'], cell: CellAddr, value: string): WorkerCellDraft => ({
    generation: runRef.current!.generation, sheet: activeRef.current, ...cell, source, value,
    revision: ++draftRevision.current, draftId: draftRevision.current,
  });
  const syncDraft = () => {
    const draft = coordinator.draft;
    const input = draft?.source === 'cell' ? editorInputRef.current : formulaInputRef.current;
    if (draft && input && draft.value !== input.value) {
      const next: WorkerCellDraft = { ...draft, value: input.value, prefill: false, modified: true,
        revision: ++draftRevision.current };
      coordinator.setDraft(next);
      if (next.source === 'cell') setEditing(next);
      else setFormulaDraft(next.value);
    }
  };
  const seal = (): InputSeal => {
    if (chartDrag.current) return { refused: 'gesture-active' };
    flushNudgeRef.current();
    const pending = composition.current;
    if (!pending) return {};
    if (pending.source === 'cell') {
      suppressBlur.current = true; editorInputRef.current?.blur(); suppressBlur.current = false;
    } else {
      suppressFormulaBlur.current = true; formulaInputRef.current?.blur(); suppressFormulaBlur.current = false;
    }
    return { composition: pending.done };
  };
  inputHooks.current = {
    generation: () => run?.generation ?? 0,
    capture: () => ({ sheet: active, target: JSON.stringify(selection) }),
    isReady: () => run?.ready ?? false, whenReady: () => run!.whenHydrated(),
    requestHydration: (reason) => run!.requestHydration(reason), flushEdits: () => run!.editPeer!.flush(),
    preview, seal, sync: syncDraft,
    async acknowledgeEdits() {
      if (!run?.peer || !run.editPeer) return;
      for (const op of previewJournal.slice()) {
        const sequence = completedPreviews.get(op);
        if (sequence === undefined || sequence > run.editPeer.acknowledgedSequence) continue;
        const keys = [...(previewPredecessors.get(op)?.keys() ?? [])];
        if (keys[0]) {
          const [sheet, row, col] = keys[0].split(':').map(Number);
          previewState.committed = { generation: run.generation, sheet, row, col,
            value: run.peer.cell(sheet, row, col).input, source: 'cell' };
        }
        for (const key of keys) {
          const [sheet, row, col] = key.split(':').map(Number);
          relinkPredecessors(key, op, { input: run.peer.cell(sheet, row, col).input, index: -1 });
          if (previewOps.get(key) === op) {
            previewOps.delete(key);
            acceptedCells.delete(key);
          }
        }
        previewJournal.splice(previewJournal.indexOf(op), 1);
        completedPreviews.delete(op);
        if (previewDraft && draftOperations.get(previewDraft) === op) {
          draftOperations.delete(previewDraft);
          draftDiscards.delete(previewDraft);
        }
      }
      if (!unpreviewedOps.size && previewJournal.every((op) => completedPreviews.has(op))) {
        previewState.sequence = run.editPeer.sentSequence;
      }
    },
    async resolveDraft(draft) {
      if (!draft.value.startsWith('=') && [...draft.value.replace(/^'/, '')].length > 32767) {
        throw new WorkerInputRefusal('Cell text exceeds 32,767 characters');
      }
      if ((draft as WorkerCellDraft).modified !== false) return draft;
      const key = `${draft.sheet}:${draft.row}:${draft.col}`;
      const accepted = acceptedCells.get(key);
      if (accepted !== undefined) return { ...draft, value: accepted, modified: true } as WorkerCellDraft;
      const internal = run && workbookSessionInternals.get(run.session);
      const value = run?.peer ? run.peer.cell(draft.sheet, draft.row, draft.col).input :
        internal?.cellInput ? await internal.cellInput(draft.sheet, draft.row, draft.col) :
        (await run!.session.call.cellInputs(draft.sheet, address(draft))).cells[0]?.[0]?.input ?? '';
      return { ...draft, value, modified: true, unchanged: true } as WorkerCellDraft;
    },
    write(draft, markApplied) {
      const owner = run;
      if (!owner?.ready || !owner.peer || !owner.editPeer) return false;
      if (owner.peer.cell(draft.sheet, draft.row, draft.col).input !== draft.value) {
        markApplied.check();
        const op = draftOperations.get(draft);
        const operations = workbookEditPeerOperations(owner.recovering ? operationPeer.current! : owner.editPeer);
        let result: EditResult;
        try {
          result = op ? (owner.recovering ? operations.applyRecoveryOp(op) : operations.applyQueuedOp(op)) as EditResult :
            owner.editPeer.editCell(draft.sheet, draft.row, draft.col, draft.value);
        } catch (error) {
          if (inputRefusal(error)) draftDiscards.get(draft)?.();
          throw error;
        }
        markApplied();
        apply(result, op);
      } else trackCommit(draftOperations.get(draft));
      if (!owner.retiring) closeWrittenDraft(draft);
      return true;
    },
  };

  const openEditor = (seed?: string) => {
    const owner = runRef.current;
    const selected = selectionRef.current;
    if (!acceptInput() || !owner || !selected) return;
    if (coordinator.draft) {
      syncDraft(); coordinator.settle(); setFormulaDraft(null);
    }
    const cell = selected.focus;
    const grid = sourceRef.current?.painted?.displayList.grid;
    editorRectRef.current = grid ? cellRect(grid, cell.row, cell.col) : null;
    const draft: WorkerCellDraft = {
      ...draftFor('cell', cell, seed ?? owner.peer?.cell(activeRef.current, cell.row, cell.col).input ?? ''),
      prefill: seed === undefined && !owner.peer,
      modified: seed !== undefined,
    };
    coordinator.setDraft(draft);
    setEditing(draft);
    if (seed !== undefined || owner.peer) return;
    void owner.session.call.cellInputs(draft.sheet, address(cell)).then((result) => {
      if (!owner.current || owner.failure || coordinator.draft !== draft || editingRef.current !== draft) return;
      const next = { ...draft, value: result.cells[0]?.[0]?.input ?? '', prefill: false };
      const input = editorInputRef.current;
      const start = input?.selectionStart;
      const end = input?.selectionEnd;
      flushSync(() => { setEditing(next); coordinator.setDraft(next); });
      if (input && start != null && end != null) input.setSelectionRange(start, end);
    }).catch((error) => owner.fail(error));
  };

  const commitDraft = (source: InputDraft['source'], move?: Direction) => {
    if (!acceptInput() || composition.current) return;
    syncDraft();
    const draft = coordinator.draft;
    if (draft?.source !== source) return;
    if (!coordinator.submit(draft)) return;
    if ((draft as WorkerCellDraft).modified !== false && !unpreviewedOps.size) {
      previewState.sequence = Infinity;
      if (previewCanvasRef.current) previewCanvasRef.current.dataset.previewReady = 'false';
      setPreviewDraft(draft);
    }
    if (source === 'cell') { suppressBlur.current = true; setEditing(null); }
    else setFormulaDraft(null);
    const base = selectionAt({ row: draft.row, col: draft.col });
    setSelection(move ? moveFocus(base, move, { limits: limits() }) : base);
    focus();
  };
  const cancelDraft = () => {
    coordinator.setDraft(null); suppressBlur.current = true; setEditing(null); setFormulaDraft(null); focus();
  };

  const clipboard = (kind: 'copy' | 'cut' | 'paste') => {
    const owner = runRef.current;
    const selected = selectionRef.current;
    if (!owner?.current || owner.failure || !selected || kind !== 'copy' && !acceptInput()) return;
    const sheet = activeRef.current;
    const range = normalizeRange(selected);
    let clipboardWrite: { resolve(payload: Blob): void; reject(error: unknown): void; done: Promise<boolean> } | undefined;
    if (kind !== 'paste') {
      if (typeof ClipboardItem !== 'function' || typeof navigator.clipboard.write !== 'function') {
        if (kind === 'cut') { reportInputError(new WorkerInputRefusal('Cut requires clipboard write access')); return; }
      } else {
        let resolve!: (payload: Blob) => void;
        let reject!: (error: unknown) => void;
        const payload = new Promise<Blob>((done, failed) => { resolve = done; reject = failed; });
        void payload.catch(() => {});
        try {
          const done = navigator.clipboard.write([new ClipboardItem({ 'text/plain': payload })]).then(
            () => true, (error) => { reject(error); throw new WorkerInputRefusal(error instanceof Error ? error.message : 'Clipboard write failed', kind === 'cut'); }
          );
          void done.catch(() => {});
          clipboardWrite = { resolve, reject, done };
        } catch (error) { reject(error); reportInputError(new WorkerInputRefusal(error instanceof Error ? error.message : 'Clipboard write failed')); return; }
      }
    }
    const captureClipboard = () => kind === 'paste' ? navigator.clipboard.readText() : '';
    let mutation: Extract<WorkbookReplayOp, { method: 'editCells' }> | undefined;
    let discardPreview: (() => Promise<void>) | undefined;
    let clipboardRefusal: unknown;
    void clipboardWrite?.done.catch((error) => {
      clipboardRefusal = error;
      if (kind === 'cut' && discardPreview) void discardPreview().catch(reportInputError);
      reportInputError(error);
    });
    const prepareMutation = async (text: string) => {
      if (kind !== 'copy' && !mutation) {
        const edits: CellInputEdit[] = [];
        if (kind === 'paste') {
          fromTsv(text).forEach((row, dr) => row.forEach((input, dc) =>
            edits.push({ row: range.top + dr, col: range.left + dc, input })));
        } else {
          for (let row = range.top; row <= range.bottom; row++) {
            for (let col = range.left; col <= range.right; col++) edits.push({ row, col, input: '' });
          }
        }
        mutation = { method: 'editCells', args: [sheet, edits] };
      }
      if (mutation) discardPreview = await previewEdits(sheet, mutation.args[1], mutation);
      if (clipboardRefusal) { await discardPreview?.(); throw clipboardRefusal; }
    };
    void coordinator.clipboard(captureClipboard, async (text, _, markApplied) => {
      if (kind !== 'copy' && !mutation) await prepareMutation(text);
      if (kind !== 'paste') {
        const copied = toTsv(owner.peer!.rangeCells(sheet, selectedRange(selected)));
        let written: boolean;
        if (clipboardWrite) {
          clipboardWrite.resolve(new Blob([copied], { type: 'text/plain' }));
          written = await clipboardWrite.done;
        } else {
          try { await navigator.clipboard.writeText(copied); written = true; }
          catch (error) { throw new WorkerInputRefusal(error instanceof Error ? error.message : 'Clipboard write failed'); }
        }
        if (!owner.current) throw new XlsxCommandAdmissionError('document-replaced');
        markApplied.check();
        if (!written) throw new WorkerInputRefusal('Clipboard write was refused', kind === 'cut');
        if (kind === 'copy') return;
      }
      const op = mutation!;
      const edits = op.args[1];
      if (edits.length === 0) return;
      markApplied.check();
      const operations = workbookEditPeerOperations(owner.recovering ? operationPeer.current! : owner.editPeer!);
      const result = (owner.recovering ? operations.applyRecoveryOp(op) : operations.applyQueuedOp(op)) as EditResult;
      markApplied();
      apply(result, op);
      if (kind === 'paste' && selectionRef.current === selected && activeRef.current === sheet) {
        setSelection({ anchor: { row: range.top, col: range.left }, focus: {
          row: Math.max(...edits.map((edit) => edit.row)), col: Math.max(...edits.map((edit) => edit.col)),
        } });
      }
    }, { sheet, target: JSON.stringify(range) }, kind !== 'copy' ? prepareMutation : undefined, kind !== 'copy').catch(async (error) => {
      clipboardWrite?.reject(error);
      if (inputRefusal(error)) await discardPreview?.();
      reportInputError(error);
    });
  };

  const clearCells = () => {
    if (!acceptInput() || !selectionRef.current) return;
    const sheet = activeRef.current;
    const range = normalizeRange(selectionRef.current);
    const edits: CellInputEdit[] = [];
    for (let row = range.top; row <= range.bottom; row++) {
      for (let col = range.left; col <= range.right; col++) edits.push({ row, col, input: '' });
    }
    const op: WorkbookReplayOp = { method: 'editCells', args: [sheet, edits] };
    let discardPreview: (() => Promise<void>) | undefined;
    void coordinator.runAfterPendingInput(async (_, markApplied) => {
      markApplied.check();
      const operations = workbookEditPeerOperations(run!.recovering ? operationPeer.current! : run!.editPeer!);
      const result = (run!.recovering ? operations.applyRecoveryOp(op) : operations.applyQueuedOp(op)) as EditResult;
      markApplied(); apply(result, op);
    }, { kind: 'input', target: { sheet, target: JSON.stringify(range) },
      prepare: async () => { discardPreview = await previewEdits(sheet, edits, op); } }).catch(async (error) => {
      if (inputRefusal(error)) await discardPreview?.();
      reportInputError(error);
    });
  };

  const navigate = (sheet: number, next: Selection) => {
    const owner = run;
    if (!owner?.current || owner.failure) return;
    const target = structuredClone(next);
    void coordinator.runAfterPendingInput(async (_, markApplied) => {
      markApplied.check();
      owner.editPeer!.setActiveSheet(sheet);
      markApplied(); place(sheet, target);
    }, { kind: 'input', target: { sheet, target: JSON.stringify(target) } }).catch(reportInputError);
  };

  const moveChart = (sheet: number, id: string, dx: number, dy: number, retained = false) => {
    const owner = run;
    if (!owner?.current || !retained && !acceptInput() || dx === 0 && dy === 0) return;
    const operation = async (_: unknown, markApplied: WorkerInputLease) => {
      await boundary();
      markApplied.check();
      const result = owner.editPeer!.moveChart(sheet, id, dx, dy);
      markApplied(); apply(result); setChartOffset(null);
    };
    const options = { kind: 'chart' as const, target: { sheet, target: id }, input: { dx, dy } };
    void (retained ? coordinator.input(operation, options) :
      coordinator.runAfterPendingInput(operation, options)).catch(reportInputError);
  };
  const flushNudge = (retained = false) => {
    const pending = nudge.current;
    nudge.current = null;
    if (pending) moveChart(pending.sheet, pending.id, pending.dx, pending.dy, retained);
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    nudgeTimer.current = null;
  };
  flushNudgeRef.current = flushNudge;
  if (run) run.beforeRetire = () => {
    coordinator.retire();
    retirement.resolve();
    sourceRef.current?.dispose();
    syncDraft();
    composition.current?.settle(true);
    composition.current = null;
    flushNudge(true);
  };
  const moveChartRef = useRef(moveChart);
  moveChartRef.current = moveChart;
  useEffect(() => {
    const stop = (event: globalThis.MouseEvent) => {
      dragging.current = false;
      const drag = chartDrag.current;
      chartDrag.current = null;
      if (drag && event.button === 0) moveChartRef.current(drag.sheet, drag.id,
        (event.clientX - drag.clientX) / drag.zoom, (event.clientY - drag.clientY) / drag.zoom);
      else setChartOffset(null);
    };
    const cancel = () => { dragging.current = false; chartDrag.current = null; setChartOffset(null); flushNudgeRef.current(); };
    window.addEventListener('mouseup', stop);
    window.addEventListener('blur', cancel);
    return () => { window.removeEventListener('mouseup', stop); window.removeEventListener('blur', cancel); };
  }, []);

  const point = (event: MouseEvent) => {
    const painted = sourceRef.current?.painted;
    const canvas = canvasRef.current;
    if (!painted || painted.request.sheet !== activeRef.current || !canvas) return { cell: null, chart: null };
    const rect = canvas.getBoundingClientRect();
    const x = (event.clientX - rect.left) / painted.request.zoom;
    const y = (event.clientY - rect.top) / painted.request.zoom;
    return { cell: painted.geometry ? cellAtPoint(painted.geometry, x, y) : null,
      chart: chartRegionAtPoint(painted.displayList.charts, x, y) };
  };
  const onMouseDown = (event: MouseEvent) => {
    if (fromPlugin(event) || runRef.current?.failure || editorInputRef.current?.contains(event.target as Node)) return;
    const { cell, chart } = point(event);
    const dismissing = editingRef.current !== null;
    if (dismissing) commitDraft('cell');
    else if (coordinator.draft?.source === 'formula') commitDraft('formula');
    flushNudgeRef.current();
    chartDrag.current = null;
    clickStart.current = null;
    if (chart) {
      setSelectedChart({ id: chart.id, movable: chart.movable });
      if (acceptInput() && chart.movable && event.button === 0) {
        chartDrag.current = { id: chart.id, sheet: activeRef.current,
          zoom: sourceRef.current?.painted?.request.zoom ?? zoomRef.current,
          clientX: event.clientX, clientY: event.clientY };
      }
      focus(); event.preventDefault(); return;
    }
    setSelectedChart(null);
    if (!cell || !selectionRef.current) return;
    setSelection(event.shiftKey ? extendTo(selectionRef.current, cell, limits()) : selectionAt(cell));
    clickStart.current = dismissing ? null : cell;
    dragging.current = true;
    focus();
  };
  const onMouseMove = (event: MouseEvent<HTMLDivElement>) => {
    if (fromPlugin(event) || runRef.current?.failure) return;
    const drag = chartDrag.current;
    if (drag) {
      if (!event.buttons) { chartDrag.current = null; setChartOffset(null); return; }
      setChartOffset({ x: event.clientX - drag.clientX, y: event.clientY - drag.clientY });
      event.currentTarget.style.cursor = 'move'; event.currentTarget.title = ''; return;
    }
    const { cell, chart } = point(event);
    if (dragging.current && cell && selectionRef.current) {
      setSelection(extendTo(selectionRef.current, cell, limits())); return;
    }
    const hyperlink = cell && sourceRef.current?.painted ?
      hyperlinkAtCell(sourceRef.current.painted.displayList, cell.row, cell.col) : null;
    event.currentTarget.style.cursor = chart?.movable && acceptInput() ? 'move' : chart || hyperlink ? 'pointer' : 'default';
    event.currentTarget.title = chart ? '' : hyperlink?.tooltip ?? '';
  };
  const onClick = (event: MouseEvent) => {
    if (fromPlugin(event) || runRef.current?.failure) return;
    const { cell } = point(event);
    const start = clickStart.current;
    clickStart.current = null;
    const frame = sourceRef.current?.painted?.displayList;
    if (!cell || !start || !frame || cell.row !== start.row || cell.col !== start.col) return;
    const hyperlink = hyperlinkAtCell(frame, cell.row, cell.col);
    if (!hyperlink) return;
    const href = safeExternalHyperlink(hyperlink);
    if (href) { window.open(href, '_blank', 'noopener,noreferrer'); return; }
    if (!hyperlink.location) return;
    const sheets = runRef.current!.session.state.sheets;
    const destination = parseHyperlinkLocation(hyperlink.location, sheets[activeRef.current]?.name ?? '');
    if (!destination) return;
    const sheet = sheets.findIndex((entry) => entry.name.toLowerCase() === destination.sheetName.toLowerCase());
    if (sheet >= 0) navigate(sheet, selectionAt(destination));
  };
  const onKeyDown = (event: KeyboardEvent) => {
    if (fromPlugin(event) || runRef.current?.failure || !viewRef.current || !selectionRef.current ||
      editingRef.current || commandForEvent(event)) return;
    const mod = event.metaKey || event.ctrlKey;
    const chart = chartRef.current;
    if (chart) {
      if (event.key === 'Escape') {
        chartDrag.current = null; nudge.current = null; setChartOffset(null); setSelectedChart(null);
        if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
        nudgeTimer.current = null; event.preventDefault(); return;
      }
      const delta = !mod && chartKeys[event.key];
      if (delta) {
        if (acceptInput() && chart.movable) {
          const step = event.shiftKey ? 10 : 1;
          const paintedZoom = sourceRef.current?.painted?.request.zoom ?? zoomRef.current;
          const base = nudge.current ?? { id: chart.id, sheet: activeRef.current, dx: 0, dy: 0 };
          nudge.current = { ...base, dx: base.dx + delta[0] * step, dy: base.dy + delta[1] * step };
          setChartOffset({ x: nudge.current.dx * paintedZoom, y: nudge.current.dy * paintedZoom });
          if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
          nudgeTimer.current = setTimeout(() => flushNudgeRef.current(), 250);
        }
        event.preventDefault();
      }
      return;
    }
    if (mod) {
      const key = event.key.toLowerCase();
      if (key === 'c' || key === 'x' || key === 'v') {
        clipboard(key === 'c' ? 'copy' : key === 'x' ? 'cut' : 'paste'); event.preventDefault(); return;
      }
    }
    const action = selectionKeyReducer(selectionRef.current, event, limits());
    if (action.type === 'move') setSelection(action.selection);
    else if (action.type === 'startEdit') openEditor(action.initialInput);
    else if (action.type === 'clear') clearCells();
    if (action.type !== 'none') event.preventDefault();
  };

  const reveal = (peer: WorkbookHandle, sheet: number, row: number, col: number, align: XlsxRevealAlignment,
    live: () => boolean) => {
    const owner = runRef.current;
    const scroll = scrollRef.current;
    if (!owner?.current || !live() || !scroll || activeRef.current !== sheet) return;
    const info = peer.sheetInfo();
    const at = peer.cellPosition(sheet, row, col);
    const next = peer.cellPosition(sheet, Math.min(row + 1, 1_048_575), Math.min(col + 1, 16_383));
    const zoom = zoomRef.current;
    const frozen = sourceRef.current?.painted?.displayList.grid;
    const x = revealAxis(at.x, next.x - at.x, col < info.frozenCols, scroll.scrollLeft / zoom,
      scroll.clientWidth / zoom - (frozen?.colOffsets[info.frozenCols] ?? 0), align);
    const y = revealAxis(at.y, next.y - at.y, row < info.frozenRows, scroll.scrollTop / zoom,
      scroll.clientHeight / zoom - (frozen?.rowOffsets[info.frozenRows] ?? 0), align);
    if (x !== null) scroll.scrollLeft = x * zoom;
    if (y !== null) scroll.scrollTop = y * zoom;
    if (!owner.recovering) sourceRef.current?.schedule();
  };
  const validCell = (peer: WorkbookHandle, sheet: number, cell: CellAddr | undefined) => {
    if (!cell || !Number.isInteger(cell.row) || !Number.isInteger(cell.col) || cell.row < 0 || cell.col < 0 ||
      cell.row > 1_048_575 || cell.col > 16_383) return false;
    try { peer.cellPosition(sheet, cell.row, cell.col); return true; } catch { return false; }
  };
  pluginAccessRef.current = {
    handle: pluginAccess.handle, commands: () => commands,
    async admit(handle, operation) {
      try {
        const value = await coordinator.runAfterPendingInput(async (_, markApplied) => {
          if (runRef.current?.peer !== handle) throw new XlsxCommandAdmissionError('document-replaced');
          await boundary();
          markApplied.check();
          const value = operation();
          if (runRef.current?.recovering && value !== null && typeof value === 'object' &&
            ('refused' in value || 'ok' in value && value.ok === false)) {
            throw new XlsxCommandAdmissionError('input-failed');
          }
          if (value !== null && typeof value === 'object' && ('refused' in value || 'ok' in value && value.ok === false)) {
            markApplied.refuse(new WorkerInputRefusal('Plugin request was refused'));
          } else markApplied();
          return value;
        }, { kind: 'plugin', recover: false });
        return { ok: true, value };
      } catch (error) {
        return { ok: false, code: error instanceof XlsxCommandAdmissionError && error.code === 'document-replaced'
          ? 'document-replaced' : 'input-failed', error };
      }
    },
    async admitEdits(handle, request, authorize, commit) {
      const owner = run;
      const facade = owner?.editPeer;
      if (!owner || !facade || owner.peer !== handle) return {
        ok: false, code: 'document-replaced', error: new XlsxCommandAdmissionError('document-replaced'),
      };
      const op: WorkbookReplayOp = { method: 'applyEdits', args: [structuredClone(request)] };
      const operations = workbookEditPeerOperations(facade);
      const readOnly = pluginAccessRef.current!.readOnlyRefusal(handle);
      try {
        const value = await coordinator.runAfterPendingInput(async (_, applied) => {
          await boundary();
          applied.check();
          const denied = authorize(owner.recovering);
          if (denied) { applied.refuse(new WorkerInputRefusal('Plugin write permission was refused')); return denied; }
          if (readOnly) { applied.refuse(new WorkerInputRefusal(readOnly.failure.message)); return readOnly; }
          const write = () => owner.recovering ? operations.applyRecoveryOp(op) : operations.applyQueuedOp(op);
          let result: XlsxEditResult;
          try { result = (owner.recovering ? write() : commit(write)) as XlsxEditResult; }
          catch (error) {
            if (!(error instanceof WorkbookRecoveryRefusal)) throw error;
            result = error.result as XlsxEditResult;
          }
          if (!result.ok) {
            applied.refuse(new WorkerInputRefusal(result.failure.message), true);
            return result;
          }
          applied();
          apply(result);
          return result;
        }, { kind: 'plugin', input: op });
        return { ok: true, value };
      } catch (error) {
        return { ok: false, code: error instanceof XlsxCommandAdmissionError && error.code === 'document-replaced'
          ? 'document-replaced' : 'input-failed', error };
      }
    },
    readOnlyRefusal: (handle) => propsRef.current.readOnly ? {
      ok: false, version: handle.version(), failure: { code: 'read-only', message: 'The editor is read-only' },
    } : null,
    applyEdits(handle, request, authorize, commit) {
      const denied = authorize();
      if (denied) return denied;
      const readOnly = pluginAccessRef.current!.readOnlyRefusal(handle);
      if (readOnly) return readOnly;
      const result = commit(() => runRef.current!.editPeer!.applyEdits(request));
      apply(result); return result;
    },
    navigator: {
      selectCells(peer, target, wantsFocus, live) {
        const sheet = peer.sheetInfo().sheetIds.indexOf(target?.sheetId);
        if (sheet < 0 || !validCell(peer, sheet, target?.selection?.anchor) ||
          !validCell(peer, sheet, target?.selection?.focus)) return 'missing-target';
        runRef.current!.editPeer!.setActiveSheet(sheet);
        place(sheet, target.selection);
        reveal(peer, sheet, target.selection.focus.row, target.selection.focus.col, 'nearest', live);
        if (wantsFocus) focus();
        return null;
      },
      scrollToCell(peer, target, align, live) {
        const sheet = peer.sheetInfo().sheetIds.indexOf(target?.sheetId);
        if (sheet < 0 || !validCell(peer, sheet, target)) return 'missing-target';
        if (sheet !== activeRef.current) {
          runRef.current!.editPeer!.setActiveSheet(sheet); place(sheet, null);
        }
        reveal(peer, sheet, target.row, target.col, align, live); return null;
      },
    },
  };

  const requestSave = useSaveRequest({
    document: () => runRef.current?.current && !runRef.current.retiring && !runRef.current.failure ? runRef.current : null,
    onSaveRequest: props.onSaveRequest,
    fail: reportInputError,
  });
  useWorkerXlsxCommands(commands, {
    requestSave,
    peer: () => run?.peer ?? null, editPeer: () => run?.editPeer ?? null,
    status: () => run?.ready ? 'ready' : run && !run.failure || loading ? 'loading' : 'empty',
    readOnly: () => propsRef.current.readOnly ?? false, collaborative: () => false,
    mutation: () => mutationRef.current, generation: () => run?.generation ?? 0,
    view: () => ({ sheet: activeRef.current, selection: selectionRef.current, chartSelected: chartRef.current !== null,
      zoom: zoomRef.current, capturedFormat: capturedFormatRef.current, borderStyle: borderStyleRef.current,
      borderColor: borderColorRef.current, proposals: proposalsRef.current, proposalsAvailable: isProposalsAvailable(),
      proposalsPanelOpen: panelRef.current, pngExport: false }),
    coordinator, preview: () => preview(), recovering: () => runRef.current?.recovering ?? false,
    i18n: () => propsRef.current.i18n, translate: t, apply, fail: (error) => run?.fail(error),
    setZoom: (scale) => { setZoom(scale); sourceRef.current?.schedule(); },
    setProposalsPanelOpen, setCapturedFormat: (format, source) => {
      setCapturedFormat(format); paintFormatSource.current = source;
    },
    setBorderStyle: (style) => { borderStyleRef.current = style; },
    setBorderColor: (color) => { borderColorRef.current = color; },
    markStale: (id, cells) => setStaleFor(({ [id]: _removed, ...rest }) => cells ? { ...rest, [id]: cells } : rest),
    refreshProposals,
    deliver(bytes) {
      if (propsRef.current.onSave) propsRef.current.onSave(bytes);
      else {
        const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: XLSX_MIME }));
        const anchor = document.createElement('a');
        anchor.href = url; anchor.download = propsRef.current.fileName ?? 'workbook.xlsx';
        document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
      }
    },
    afterPaint, focusGrid: focus, refuse: (reason) => setRefusal(reason.message),
  });
  useCommandShortcuts(commands, rootRef);

  useEffect(() => {
    if (!capturedFormat || !selection || !acceptInput() || dragging.current) return;
    const range = normalizeRange(selection);
    const key = `${active}:${range.top}:${range.left}:${range.bottom}:${range.right}`;
    if (key === paintFormatSource.current) return;
    const format = capturedFormat;
    const a1 = selectedRange(selection);
    setCapturedFormat(null); paintFormatSource.current = null;
    void coordinator.runAfterPendingInput(async (_, markApplied) => {
      await boundary();
      markApplied.check();
      const result = run!.editPeer!.applyFormat(active, a1, format);
      markApplied(); apply(result); focus();
    }, { kind: 'input', target: { sheet: active, target: a1 } }).catch(reportInputError);
  }, [capturedFormat, selection, active, coordinator, boundary, apply, reportInputError, focus]);

  useLayoutEffect(() => {
    if (!props.readOnly) return;
    syncDraft();
    settledDraftRevision.current = draftRevision.current;
    const source = coordinator.draft?.source;
    suppressBlur.current = true;
    suppressFormulaBlur.current = true;
    composition.current?.settle(true);
    composition.current = null;
    setEditing(null);
    setFormulaDraft(null);
    setCapturedFormat(null);
    paintFormatSource.current = null;
    dragging.current = false;
    clickStart.current = null;
    chartDrag.current = null;
    nudge.current = null;
    setChartOffset(null);
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    nudgeTimer.current = null;
    if (coordinator.draft || coordinator.pending) void coordinator.drain().catch(reportInputError);
    if (source === 'cell') focus();
    suppressFormulaBlur.current = false;
  }, [props.readOnly, coordinator]);

  const startComposition = (source: InputDraft['source']) => {
    if (!acceptInput()) return;
    if (source === 'formula' && coordinator.draft?.source !== 'formula' && selectionRef.current) {
      syncDraft(); coordinator.settle(); suppressBlur.current = true; setEditing(null);
      coordinator.setDraft(draftFor(source, selectionRef.current.focus, formulaInputRef.current?.value ?? ''));
    }
    let settle!: (ended: boolean) => void;
    const done = new Promise<boolean>((resolve) => { settle = resolve; });
    composition.current = { source, done, settle };
  };
  const endComposition = () => {
    if (!composition.current) return;
    syncDraft(); composition.current?.settle(true); composition.current = null;
  };
  const attachCellInput = useCallback((input: HTMLInputElement | null) => {
    if (!input) inputHooks.current?.sync?.();
    editorInputRef.current = input;
    if (input) suppressBlur.current = false;
    if (!input && composition.current?.source === 'cell') {
      composition.current.settle(true); composition.current = null;
    }
  }, [inputHooks]);
  const attachFormulaInput = useCallback((input: HTMLInputElement | null) => {
    if (!input) inputHooks.current?.sync?.();
    formulaInputRef.current = input;
    if (!input && composition.current?.source === 'formula') {
      composition.current.settle(true); composition.current = null;
    }
  }, [inputHooks]);
  useLayoutEffect(() => {
    if (!editing || run?.failure) return;
    const input = editorInputRef.current;
    if (input && document.activeElement !== input) {
      input.focus({ preventScroll: true }); input.setSelectionRange(input.value.length, input.value.length);
    }
  }, [editing !== null]);

  const formulaBar: FormulaBarBinding = {
    a1: selection ? address(selection.focus) : '',
    value: (composition.current?.source === 'formula' ? formulaInputRef.current?.value : undefined) ?? formulaDraft ?? focusedCell?.input ?? '',
    disabled: !view || !selection || Boolean(error), readOnly: !acceptInput() || selectedChart !== null,
    inputRef: attachFormulaInput,
    onChange(value) {
      const selected = selectionRef.current;
      if (!acceptInput() || !selected || chartRef.current) return;
      let draft = coordinator.draft;
      if (draft && draft.source !== 'formula') {
        syncDraft(); coordinator.settle(); suppressBlur.current = true; setEditing(null);
        draft = null;
      }
      setFormulaDraft(value);
      coordinator.setDraft({ ...(draft ?? draftFor('formula', selected.focus, value)), value, modified: true,
        revision: ++draftRevision.current } as WorkerCellDraft);
    },
    onCommit: (move) => commitDraft('formula', move), onCancel: cancelDraft,
    onBlur: () => { if (!suppressFormulaBlur.current) commitDraft('formula'); },
    onCompositionStart: () => startComposition('formula'), onCompositionEnd: endComposition,
  };
  const frame = painted?.displayList;
  useEffect(() => {
    if (!selectedChart || !frame || frame.charts?.some((chart) => chart.id === selectedChart.id)) return;
    flushNudgeRef.current(); chartDrag.current = null; setChartOffset(null); setSelectedChart(null);
  }, [selectedChart, frame]);
  const paintedZoom = painted?.request.zoom ?? zoom;
  const visibleSelection = painted?.request.sheet === active ? selection : null;
  const merged = painted?.mergedRanges ?? [];
  const grid = frame?.grid;
  const selected = visibleSelection ? expandRangeToMergedCells(normalizeRange(visibleSelection), merged) : null;
  const focused = visibleSelection ? expandRangeToMergedCells({ top: visibleSelection.focus.row,
    left: visibleSelection.focus.col, bottom: visibleSelection.focus.row, right: visibleSelection.focus.col }, merged) : null;
  const selectionRect = grid && selected ? rangeRect(grid, selected) : null;
  const focusRect = grid && focused ? rangeRect(grid, focused) : null;
  const editRect = editing ? (grid && editing.sheet === painted?.request.sheet ?
    cellRect(grid, editing.row, editing.col) : null) ?? editorRectRef.current : null;
  const previewRect = previewDraft && grid && previewDraft.sheet === painted?.request.sheet ?
    rangeRect(grid, expandRangeToMergedCells({ top: previewDraft.row, left: previewDraft.col,
      bottom: previewDraft.row, right: previewDraft.col }, merged)) : null;
  const chartRect = selectedChart ? frame?.charts?.find((chart) => chart.id === selectedChart.id)?.rect : null;
  const sheets = run?.session.state.sheets ?? [];
  const a11yGrid = frame && painted ? buildA11yGrid(frame, visibleSelection, sheets[painted.request.sheet]?.name ?? '', {
    gridLabel: t('a11y.gridLabel'), rowHeaderLabel: t('a11y.rowHeaderLabel'), columnHeaderLabel: t('a11y.columnHeaderLabel'),
    cellLabel: t('a11y.cellLabel'), cellLabelSelected: t('a11y.cellLabelSelected'), emptyCellLabel: t('a11y.emptyCellLabel'),
    emptyCellLabelSelected: t('a11y.emptyCellLabelSelected'),
  }) : null;
  const outline = (rect: NonNullable<typeof selectionRect>, border: string, background?: string): CSSProperties => {
    const scaled = scaledRect(rect, paintedZoom);
    return { position: 'absolute', left: scaled.x, top: scaled.y, width: scaled.w, height: scaled.h,
      boxSizing: 'border-box', border, background };
  };
  const defaultToolbar = <EditorToolbar mode="commands">
    <EditorToolbar.Toolbar />
    <div style={{ display: 'flex', padding: '4px 8px' }}>
      <ToolbarCommandButton id="save" />
      <ToolbarCommandButton id="exportPng" />
      <EditorToolbar.FormulaBar />
      {isProposalsAvailable() && <ToolbarCommandButton id="proposalsPanel" />}
    </div>
  </EditorToolbar>;
  const toolbar = props.showToolbar === false ? null : props.toolbar === undefined ? props.readOnly ? null : defaultToolbar : props.toolbar;
  const renderDock = (placement: DockPlacement) => pluginHost.managed ? <PluginDock host={pluginHost.host}
    placement={placement} activations={pluginHost.activations.filter((activation) => activation.plugin.panel?.placement === placement)}
    available={dockArea} /> : null;

  return <XlsxCommandContext.Provider value={commands.store}>
    <EditorChromeContext.Provider value={true}>
      <FormulaBarContext.Provider value={formulaBar}>
        <div ref={rootRef} className={props.className} role="application" aria-label={t('editor.appLabel')}
          style={{ position: 'relative', display: 'flex', flexDirection: 'column', width: '100%', height: '100%',
            minWidth: 0, color: '#202124', background: '#ffffff', fontFamily: 'ui-sans-serif, system-ui, sans-serif' }}>
          {toolbar != null && <div data-testid="xlsx-toolbar" style={{ flex: '0 0 auto' }}>{toolbar}</div>}
          <div ref={dockAreaRef} data-testid="xlsx-workspace" style={{ position: 'relative', display: 'flex', flex: 1, minWidth: 0, minHeight: 0 }}>
            {renderDock('left')}
            <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, minHeight: 0 }}>
              {isProposalsAvailable() && proposalsPanelOpen && <ProposalsPanel proposals={proposals} staleFor={staleFor}
                style={{ top: 4, right: 8, width: 'min(320px, calc(100% - 16px))', maxHeight: 'min(420px, calc(100% - 12px))' }} />}
              <div ref={scrollRef} data-testid="xlsx-scroll" tabIndex={0} onKeyDown={onKeyDown}
                onMouseDown={onMouseDown} onMouseMove={onMouseMove} onClick={onClick}
                onDoubleClick={(event) => {
                  if (fromPlugin(event) || editingRef.current || point(event).chart) return;
                  const cell = point(event).cell;
                  if (cell && frame && hyperlinkAtCell(frame, cell.row, cell.col)) return;
                  openEditor();
                }}
                onMouseLeave={(event) => { event.currentTarget.style.cursor = 'default'; event.currentTarget.title = ''; }}
                style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none' }}>
                <div style={{ position: 'absolute', top: 0, left: 0, width: view ? view.contentWidth * zoom : '100%',
                  height: view ? view.contentHeight * zoom : '100%' }} />
                <div style={{ position: 'sticky', top: 0, left: 0, width: 0, height: 0 }}>
                  <canvas ref={canvasRef} style={{ display: 'block', position: 'absolute', top: 0, left: 0 }} />
                  {previewDraft && (run?.ready ? <div data-testid="xlsx-commit-preview" aria-hidden
                    style={{ ...outline(previewRect ?? { x: 0, y: 0, w: 96, h: 24 }, 'none', '#fff'),
                      color: '#000', padding: `0 ${8 * paintedZoom}px`, overflow: 'hidden', whiteSpace: 'pre',
                      display: previewDraft.sheet === active && (!grid || previewRect) ? 'block' : 'none',
                      font: `${13 * paintedZoom}px system-ui, sans-serif`,
                      lineHeight: `${(previewRect?.h ?? 24) * paintedZoom}px`, pointerEvents: 'none' }}>{previewDraft.value}</div> :
                    <canvas ref={previewCanvasRef} data-testid="xlsx-commit-preview" aria-hidden
                      style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none' }}>{previewDraft.value}</canvas>)}
                  {pluginHost.managed && <PluginOverlays host={pluginHost.host} activations={pluginHost.activations}
                    layerRef={pluginHost.overlayLayerRef} width={(frame?.width ?? 0) * paintedZoom} height={(frame?.height ?? 0) * paintedZoom} />}
                  <div data-testid="xlsx-overlay-host" style={{ position: 'absolute', top: 0, left: 0, width: 0, height: 0, pointerEvents: 'none' }}>
                    {selectionRect && !selectedChart && <div data-testid="xlsx-selection"
                      style={outline(selectionRect, `1px solid ${BRAND}`, 'rgba(33, 115, 70, 0.12)')} />}
                    {focusRect && !selectedChart && !editing && <div style={outline(focusRect, `2px solid ${BRAND}`)} />}
                    {chartRect && <div data-testid="xlsx-chart-selection" data-chart-id={selectedChart?.id} aria-hidden
                      style={{ ...outline(chartRect, `2px solid ${BRAND}`),
                        transform: `translate(${chartOffset?.x ?? 0}px, ${chartOffset?.y ?? 0}px)`, boxShadow: '0 1px 6px rgba(0, 0, 0, 0.25)' }} />}
                    {editing && editRect && <input ref={attachCellInput} data-testid="xlsx-cell-editor"
                      value={(composition.current?.source === 'cell' ? editorInputRef.current?.value : undefined) ?? editing.value}
                      disabled={Boolean(error) || props.readOnly}
                      onChange={(event) => {
                        if (!acceptInput()) return;
                        const draft: WorkerCellDraft = { ...editing, value: event.target.value, prefill: false, modified: true,
                          revision: ++draftRevision.current };
                        setEditing(draft); coordinator.setDraft(draft);
                      }}
                      onCompositionStart={() => startComposition('cell')} onCompositionEnd={endComposition}
                      onKeyDown={(event) => {
                        if (!commandForEvent(event)) event.stopPropagation();
                        if (event.nativeEvent.isComposing || event.keyCode === 229) return;
                        if (event.key === 'Enter') { commitDraft('cell', event.shiftKey ? 'up' : 'down'); event.preventDefault(); }
                        else if (event.key === 'Tab') { commitDraft('cell', event.shiftKey ? 'left' : 'right'); event.preventDefault(); }
                        else if (event.key === 'Escape') { cancelDraft(); event.preventDefault(); }
                      }}
                      onBlur={() => {
                        if (suppressBlur.current) { suppressBlur.current = false; return; }
                        commitDraft('cell');
                      }}
                      style={{ ...outline(editRect, `2px solid ${BRAND}`, '#ffffff'), padding: '0 3px',
                        font: `${13 * paintedZoom}px system-ui, sans-serif`, pointerEvents: 'auto', outline: 'none' }} />}
                  </div>
                </div>
              </div>
              {renderDock('bottom')}
            </div>
            {renderDock('right')}
          </div>
          {a11yGrid && <>
            <div style={visuallyHidden} role="grid" aria-label={a11yGrid.label}>
              <div role="row"><span role="columnheader" />{a11yGrid.columnHeaders.map((header) =>
                <span key={header.col} role="columnheader">{header.label}</span>)}</div>
              {a11yGrid.rows.map((row) => <div key={row.row} role="row"><span role="rowheader">{row.header}</span>
                {row.cells.map((cell) => <span key={cell.col} role="gridcell" aria-selected={cell.selected}>{cell.label}</span>)}
              </div>)}
            </div>
            {a11yGrid.charts.map((chart, index) => <div key={`${index}:${chart.label}`} style={visuallyHidden} role="img" aria-label={chart.label} />)}
          </>}
          {refusal && <div data-testid="xlsx-input-refusal" role="alert" style={{ position: 'relative', zIndex: 1 }}>{refusal}</div>}
          {error && <div data-testid="xlsx-error" role="alert" style={{ position: 'absolute', inset: 0, display: 'grid',
            placeItems: 'center', padding: 16, textAlign: 'center', color: '#b00020' }}>{t('editor.openError')}: {error.message}</div>}
          {sheets.length > 0 && <div data-testid="xlsx-sheet-tabs" role="tablist" aria-label={t('editor.sheetTabsLabel')}
            style={{ display: 'flex', gap: 2, padding: '4px 6px', borderTop: '1px solid #e0e0e0', background: '#fafafa', overflowX: 'auto' }}>
            {sheets.map((sheet, index) => <button key={sheet.id} role="tab" aria-selected={index === active} disabled={Boolean(error)}
              onClick={() => navigate(index, selectionAt({ row: 0, col: 0 }))}
              style={{ border: 'none', padding: '4px 12px', cursor: 'pointer', background: index === active ? '#ffffff' : 'transparent',
                borderBottom: index === active ? `2px solid ${BRAND}` : '2px solid transparent', fontWeight: index === active ? 600 : 400 }}>
              {sheet.name}
            </button>)}
          </div>}
        </div>
      </FormulaBarContext.Provider>
    </EditorChromeContext.Provider>
  </XlsxCommandContext.Provider>;
}


const PeerlessDefaultToolbar = memo(function PeerlessDefaultToolbar() {
  return <EditorToolbar mode="commands"><EditorToolbar.Toolbar />
    <div style={{ display: 'flex', padding: '4px 8px' }}>
      <ToolbarCommandButton id="save" /><ToolbarCommandButton id="exportPng" /><EditorToolbar.FormulaBar />
      {isProposalsAvailable() && <ToolbarCommandButton id="proposalsPanel" />}
    </div>
  </EditorToolbar>;
});

function PeerlessXlsxWorkerEditor(props: PeerlessSessionWorkbookProps) {
  const { t } = useTranslation();
  const rootRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const pluginBindingRef = useRef<XlsxPluginHostBinding | null>(null);
  useLayoutEffect(() => () => pluginBindingRef.current?.host.close('unmounted'), []);
  const propsRef = useRef(props);
  propsRef.current = props;
  const [commands] = useState(createXlsxCommandController);
  const [active, setActive, activeRef] = useSyncedState(0);
  const [presentation] = useState(() => new WorkerPresentationStore());
  const { setSelection, selectionRef, setView, viewRef, setPainted } = presentation;
  const [catalog, setCatalog, catalogRef] = useSyncedState<WorkbookSheetSummary[]>([]);
  const [zoom, setZoom, zoomRef] = useSyncedState(1);
  const [refusal, setRefusal] = useState<{ code: XlsxCommandFailureCode; detail?: string } | null>(null);
  const [selectedChart, setSelectedChart, chartRef] = useSyncedState<{ id: string; movable: boolean } | null>(null);
  const [chartOffset, setChartOffset] = useState<{ x: number; y: number } | null>(null);
  const [formatToken, setFormatTokenState, formatTokenRef] = useSyncedState<string | null>(null);
  const formatSource = useRef<string | null>(null);
  const formatRevision = useRef(0);
  const applyingFormat = useRef<number | null>(null);
  const setFormatToken = (token: string | null) => {
    formatRevision.current++;
    applyingFormat.current = null;
    setFormatTokenState(token);
    if (!token) formatSource.current = null;
  };
  const borderStyle = useRef<BorderStyle | undefined>(undefined);
  const borderColor = useRef<string | undefined>(undefined);
  const [proposals, setProposals, proposalsRef] = useSyncedState<Proposal[]>([]);
  const [proposalsPanelOpen, setProposalsPanelOpen, panelRef] = useSyncedState(false);
  const [staleFor, setStaleFor] = useState<Record<string, string[]>>({});
  const navigation = useRef(0);
  const causalSelection = useRef<number | null>(null);
  const mutation = useRef(0);
  const sourceRef = useRef<PeerlessWorkerPaintSource | null>(null);
  const echoRef = useRef<WorkerCellEcho | null>(null);
  const dragging = useRef(false);
  const [selectionFinished, setSelectionFinished] = useState(0);
  const clickStart = useRef<CellAddr | null>(null);
  const chartDrag = useRef<{ generation: number; sheet: number; id: string; zoom: number; x: number; y: number } | null>(null);
  const nudge = useRef<{ generation: number; sheet: number; id: string; dx: number; dy: number } | null>(null);
  const nudgeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const revealing = useRef<{ generation: number; navigation: number } | null>(null);
  const flushNudgeRef = useRef<(retained?: boolean) => void>(() => {});
  const focus = useCallback(() => scrollRef.current?.focus({ preventScroll: true }), []);
  const exportInput = useRef<ReturnType<typeof createPeerlessWorkerInputCoordinator> | null>(null);
  const [bridge] = useState<WorkerEditorApiBridge & XlsxAsyncExportBridge>(() => ({
    coordinator: () => null, readOnly: () => propsRef.current.readOnly ?? false, focus,
    exportSheet: () => activeRef.current, exportCoordinator: () => exportInput.current,
    clearSelection: () => {}, refreshProposals: () => {}, recoverInput: async () => {},
    selectCells: () => false, selectCellsAsync: async () => false, apply: () => {},
  }));
  const { run, error, loading } = useEditableSessionWorkbook(props, commands.store, bridge, {
    backend: 'peerless', initialViewport: () => {
      const scroll = scrollRef.current;
      if (!scroll?.clientWidth || !scroll.clientHeight) return undefined;
      return { width: scroll.clientWidth, height: scroll.clientHeight, zoom: zoomRef.current,
        dpr: window.devicePixelRatio || 1, requestGeneration: 0 };
    },
  });
  const ownerRef = useRef(run);
  ownerRef.current = run;
  const editable = () => Boolean(ownerRef.current?.ready && !ownerRef.current.retiring &&
    !ownerRef.current.failure && !ownerRef.current.readOnly && !propsRef.current.readOnly && !coordinatorRef.current?.readOnly);
  const report = useCallback((error: unknown) => {
    if (error instanceof PeerlessWorkerCommandRefusal && error.status.reason === 'read-only') {
      setRefusal({ code: 'read-only' });
    } else if (error instanceof XlsxCommandAdmissionError) setRefusal({ code: error.code });
    else setRefusal({ code: 'input-failed', detail: error instanceof Error ? error.message : String(error) });
  }, []);
  const rangeTarget = (sheet: number, selected: Selection): WorkbookCommandSelection => {
    const sheetId = ownerRef.current?.session.editState.sheets?.find((entry) => entry.index === sheet)?.id ??
      ownerRef.current?.session.state.sheets.find((entry) => entry.index === sheet)?.id;
    if (!sheetId) throw new XlsxCommandAdmissionError('editor-unavailable');
    const range = normalizeRange(selected);
    return { sheetId, range: { kind: 'rowCol', start: { row: range.top, col: range.left },
      end: { row: range.bottom, col: range.right } }, activeCell: { ...selected.focus } };
  };
  const selectedTarget = () => {
    const owner = ownerRef.current;
    const sheets = owner?.session.editState.sheets ?? owner?.session.state.sheets;
    return selectionRef.current && sheets?.some((sheet) => sheet.index === activeRef.current)
      ? rangeTarget(activeRef.current, selectionRef.current) : undefined;
  };
  const limits = () => deriveLimits(sourceRef.current?.painted?.displayList ?? null, viewRef.current!,
    (scrollRef.current?.clientHeight ?? 0) / zoomRef.current);
  const sameSelection = (first: Selection | null, second: Selection) => first !== null &&
    first.anchor.row === second.anchor.row && first.anchor.col === second.anchor.col &&
    first.focus.row === second.focus.row && first.focus.col === second.focus.col;
  const predict = (sheet: number, next: Selection) => {
    if (sheet === activeRef.current && sameSelection(selectionRef.current, next)) return;
    navigation.current += 1;
    setActive(sheet);
    setSelection(structuredClone(next));
    setSelectedChart(null);
    draftController.selectionChanged();
    commands.refresh();
  };
  const reveal = (sheet: number, cell: CellAddr) => {
    const owner = ownerRef.current;
    const frame = sourceRef.current?.painted;
    if (!owner?.current || owner.failure || frame?.geometry && frame.request.sheet === sheet &&
      cellRect(frame.geometry, cell.row, cell.col)) return;
    const at = navigation.current;
    if (revealing.current?.generation === owner.generation && revealing.current.navigation === at) return;
    const request = { generation: owner.generation, navigation: at };
    revealing.current = request;
    void owner.session.call.cellGeometry(sheet, cell.row, cell.col).then((geometry) => {
      const scroll = scrollRef.current;
      if (!owner.current || ownerRef.current !== owner || navigation.current !== at || !scroll) return;
      scroll.scrollLeft = geometry.scrollPosition.x * zoomRef.current;
      scroll.scrollTop = geometry.scrollPosition.y * zoomRef.current;
      sourceRef.current?.schedule();
    }).catch((error) => {
      if (owner.current && ownerRef.current === owner && navigation.current === at) owner.fail(error);
    }).finally(() => { if (revealing.current === request) revealing.current = null; });
  };
  const requests = useMemo(() => new Map<number, WorkerPaintRequest>(), [run]);
  const selectionCommands = useMemo(() => new Set<number>(), [run]);
  const selectionRevisions = useMemo(() => new Map<number, number>(), [run]);
  const inputEvents = useRef<(event: PeerlessWorkerInputEvent, publish?: () => void) => void>(() => {});
  const coordinatorRef = useRef<ReturnType<typeof createPeerlessWorkerInputCoordinator> | null>(null);
  const draftController = useMemo(() => new WorkerDraftController({
    generation: () => run?.generation ?? 0,
    identity: () => `${run?.generation}:${run?.session.commandIdentity.session}:${run?.session.commandIdentity.incarnation}:${run?.session.confirmedBasis.epoch}`,
    navigation: () => navigation.current, selection: () => selectionRef.current, sheet: () => activeRef.current,
    target: (sheet, row, col) => rangeTarget(sheet, selectionAt({ row, col })),
    resolveSheet: (sheetId) => (run?.session.editState.sheets ?? run?.session.state.sheets)?.find((sheet) => sheet.id === sheetId)?.index,
    editable,
    prefill: async (sheet, cell) => (await run!.session.call.cellInputs(sheet, address(cell))).cells[0]?.[0]?.input ?? '',
    select: (sheet, cell, move) => {
      const base = selectionAt(cell);
      predict(sheet, move ? moveFocus(base, move, { limits: limits() }) : base);
      reveal(sheet, selectionRef.current!.focus);
    },
    selectionTarget: selectedTarget,
    dependency: () => causalSelection.current,
    navigate: () => {
      const selected = selectedTarget();
      if (selected) void coordinatorRef.current?.command({ kind: 'select', payload: { selection: selected } }).result.catch(report);
    },
    committed: (ticket) => { void ticket.result.catch(report); commands.refresh(); },
    fail: (error) => run?.fail(error), focus,
  }), [run]);
  const capture = useCallback((): WorkerPaintRequest | null => {
    const owner = ownerRef.current;
    const scroll = scrollRef.current;
    if (!owner?.current || !scroll?.clientWidth || !scroll.clientHeight || !canvasRef.current || !viewRef.current) return null;
    const scale = zoomRef.current;
    return { generation: owner.generation, sheet: activeRef.current, navigation: navigation.current,
      coversThroughSeq: Math.max(owner.session.throughSeq, coordinatorRef.current?.submittedThroughSeq ?? 0),
      width: scroll.clientWidth, height: scroll.clientHeight, zoom: scale, dpr: window.devicePixelRatio || 1,
      viewport: { x: scroll.scrollLeft / scale, y: scroll.scrollTop / scale,
        width: scroll.clientWidth / scale, height: scroll.clientHeight / scale } };
  }, []);
  const coordinator = useMemo(() => createPeerlessWorkerInputCoordinator({
    generation: () => run?.generation ?? 0, session: () => run?.session ?? null,
    isReady: () => run?.ready ?? false, readOnly: () => propsRef.current.readOnly ?? false,
    seal: (): InputSeal => {
      if (chartDrag.current) return { refused: 'gesture-active' };
      flushNudgeRef.current();
      return draftController.seal();
    },
    sync: draftController.sync,
    compositionId: () => draftController.compositionId,
    dependency: (draft) => (draft as WorkerNativeDraft).after,
    target: (sheet, row, col) => rangeTarget(sheet, selectionAt({ row, col })),
    draftTarget: (draft) => (draft as WorkerNativeDraft).target,
    frame: () => sourceRef.current?.captureCommandFrame() ?? undefined,
    selectionRevision: () => navigation.current,
    onEvent: (event, publish) => inputEvents.current(event, publish), onError: (error) => run?.fail(error),
  }), [run, draftController]);
  coordinatorRef.current = coordinator;
  exportInput.current = coordinator;
  draftController.attach(coordinator);
  const pluginNavigator = useRef<XlsxPluginDocumentAccess['navigator'] | null>(null);
  const pluginNavigation = useRef<{ owner: PeerlessWorkbookSession; sheetId: string;
    clear: boolean; ticket: PeerlessWorkerCommandTicket | null } | null>(null);
  const acknowledgedSelection = useRef<{ owner: PeerlessWorkbookSession; sheetId: string;
    target: string | undefined; cells: Selection | null } | null>(null);
  const [pluginAccess] = useState(() => createPeerlessPluginDocument({
    owner: () => ownerRef.current, coordinator: () => coordinatorRef.current,
    readOnly: () => propsRef.current.readOnly ?? false, commands: () => commands,
    selection(owner) {
      const target = owner.session.selection;
      const sheets = owner.session.editState.sheets ?? owner.session.state.sheets;
      const activeSheet = owner.session.editState.activeSheet ?? owner.session.state.activeSheet;
      const sheet = sheets.find((entry) => entry.index === activeSheet);
      const selected = target && sheets.find((entry) => entry.id === target.sheetId);
      const current = sheet ?? selected;
      const projection = acknowledgedSelection.current;
      const cells = projection?.owner === owner && projection.sheetId === current?.id &&
        projection.target === JSON.stringify(target) ? projection.cells
        : target && selected?.id === current?.id ? commandSelection(target) : null;
      return current ? { sheetId: current.id, sheetIndex: current.index,
        cells: chartRef.current ? null : cells,
        chartId: chartRef.current?.id ?? null } : null;
    },
    navigator: {
      selectCells: (...args) => pluginNavigator.current!.selectCells(...args),
      scrollToCell: (...args) => pluginNavigator.current!.scrollToCell(...args),
    },
  }));
  const pluginIdentity = pluginAccess.identity();
  const pluginSelection = pluginIdentity ? pluginAccess.selection!(pluginIdentity) : null;
  const pluginHost = useXlsxPluginHost({ ...props, access: pluginAccess, commands,
    readOnly: props.readOnly ?? false, handle: pluginIdentity, selection: pluginSelection, canvasRef, t });
  pluginBindingRef.current = pluginHost;
  const [dockAreaRef, dockArea] = useDockArea(pluginHost.managed);
  useLayoutEffect(() => run?.session.bindCollaboration?.(async (update) => {
    const status = await coordinator.command({ kind: 'applyUpdate', payload: { update } }, { after: null }).result;
    if (status.status !== 'applied' && status.status !== 'noop') throw new Error(status.reason ?? 'Remote update refused');
    return status;
  }), [run, coordinator]);
  useLayoutEffect(() => { run?.connectRecoveryInput(coordinator, draftController); }, [run, coordinator, draftController]);
  const proposalsRevision = useRef(0);
  const refreshProposals = useCallback(() => {
    const owner = ownerRef.current;
    if (!owner?.ready || !isProposalsAvailable()) return;
    const revision = ++proposalsRevision.current;
    const ticket = coordinator.command({ kind: 'applyOps', payload: { toolbar: { id: 'refreshProposals', args: null } } });
    void ticket.result.then((status) => {
      if (ownerRef.current !== owner || !owner.current || revision !== proposalsRevision.current) return;
      const result = status.result as { proposals?: Proposal[] } | undefined;
      if (result?.proposals && JSON.stringify(result.proposals) !== JSON.stringify(proposalsRef.current)) {
        setProposals(result.proposals); commands.refresh();
      }
    }).catch(report);
  }, [coordinator, commands, report]);
  const geometryRevision = useRef(0);
  const refreshSheetGeometry = useCallback((paint = false) => {
    const owner = ownerRef.current;
    if (!owner?.ready || owner.failure) return;
    const sheet = activeRef.current;
    const revision = ++geometryRevision.current;
    void owner.session.call.sheetView(sheet).then((next) => {
      if (ownerRef.current !== owner || !owner.current || owner.failure ||
        revision !== geometryRevision.current || activeRef.current !== next.sheet) return;
      const changedSheet = viewRef.current?.sheet !== next.sheet;
      setView(next);
      if (paint || changedSheet) sourceRef.current?.schedule();
    }).catch((error) => {
      if (ownerRef.current === owner && owner.current && !owner.failure) owner.fail(error);
    });
  }, [setView, viewRef]);
  const selectionKey = () => {
    const selected = selectionRef.current;
    if (!selected) return '';
    const range = normalizeRange(selected);
    return `${activeRef.current}:${range.top}:${range.left}:${range.bottom}:${range.right}`;
  };
  inputEvents.current = (event, publish) => {
    if (!run?.current || run.failure) return;
    if (event.type === 'dispatched' && event.seq !== null && event.command) {
      const command = event.command;
      const revision = command.kind === 'editCells' &&
        JSON.stringify(command.payload.selection) === JSON.stringify(selectedTarget())
        ? navigation.current : event.selectionRevision ?? navigation.current;
      selectionRevisions.set(event.seq, revision);
      while (selectionRevisions.size > 256) selectionRevisions.delete(selectionRevisions.keys().next().value!);
      if (command.frame) requests.set(event.seq, { ...command.frame as WorkerPaintRequest,
        navigation: revision, coversThroughSeq: event.seq });
      while (requests.size > 256) requests.delete(requests.keys().next().value!);
      if (command.kind === 'select' || command.kind === 'undo' || command.kind === 'redo' || 'selection' in command.payload) {
        selectionCommands.add(event.seq); causalSelection.current = event.seq;
      }
      if (command.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'paintFormat') {
        const target = command.payload.toolbar.target;
        const sheet = run.session.editState.sheets?.find((entry) => entry.id === target?.sheetId)?.index;
        if (sheet !== undefined && target?.range.kind === 'rowCol') formatSource.current =
          `${sheet}:${target.range.start.row}:${target.range.start.col}:${target.range.end.row}:${target.range.end.col}`;
      }
      if (event.draft) {
        echoRef.current?.patch(event.seq, [{ ...event.draft, text: event.draft.value }]);
        draftController.submitted(event.draft, event.seq);
      }
      else if (command.kind === 'editCells') {
        const [sheet, edits] = command.payload.args;
        echoRef.current?.paste(event.seq, sheet, edits);
      }
      else if (command.kind === 'applyOps' && 'toolbar' in command.payload && command.payload.toolbar.id === 'clear') {
        const target = command.payload.toolbar.target;
        const sheet = run.session.editState.sheets?.find((entry) => entry.id === target?.sheetId)?.index;
        if (sheet !== undefined && target?.range.kind === 'rowCol') echoRef.current?.clear(event.seq, {
          sheet, top: target.range.start.row, left: target.range.start.col,
          bottom: target.range.end.row, right: target.range.end.col,
        });
      }
      return;
    }
    if (event.type === 'refused' || event.type === 'cancelled') {
      if (event.seq !== null) echoRef.current?.refuse(event.seq);
      if (event.type === 'refused' && event.draft) {
        draftController.refused(event.seq, event.draft);
        report(new PeerlessWorkerCommandRefusal(event.status ?? { seq: event.seq ?? 0, status: 'refused' }, event.target, event.draft));
      }
      return;
    }
    if (event.type !== 'coverage') return;
    echoRef.current?.acknowledge(event.ack);
    const ack = event.ack;
    const pendingNavigation = pluginNavigation.current;
    const navigationStatus = pendingNavigation?.owner === run
      ? ack.statuses.find((status) => status.seq === pendingNavigation.ticket?.seq) : undefined;
    if (pendingNavigation?.clear && navigationStatus &&
      (navigationStatus.status === 'applied' || navigationStatus.status === 'noop')) {
      acknowledgedSelection.current = { owner: run, sheetId: pendingNavigation.sheetId,
        target: JSON.stringify(run.session.selection), cells: null };
    } else if (ack.statuses.some((status) => selectionCommands.has(status.seq) &&
      (status.status === 'applied' || status.status === 'noop'))) {
      acknowledgedSelection.current = null;
    }
    for (const status of ack.statuses) {
      const result = status.result as { remoteUpdate?: boolean; proposals?: Proposal[] } | undefined;
      if (!result?.remoteUpdate) continue;
      if (result.proposals) setProposals(result.proposals);
      sourceRef.current?.schedule();
      commands.refresh();
    }
    if (ack.editState?.sheets && JSON.stringify(ack.editState.sheets) !== JSON.stringify(catalogRef.current)) {
      setCatalog(ack.editState.sheets);
    }
    const latestSelection = Math.max(0, ...selectionCommands);
    const reconcile = ack.throughSeq >= latestSelection && selectionRevisions.get(ack.throughSeq) === navigation.current &&
      !draftController.ownsFocus;
    if (ack.selection && reconcile) {
      const target = ack.selection;
      const sheet = run.session.editState.sheets?.find((entry) => entry.id === target.sheetId)?.index;
      const restored = commandSelection(target);
      if (sheet !== undefined && restored) predict(sheet, restored);
    } else if (!ack.selection && ack.editState?.activeSheet !== undefined && reconcile && ack.editState.activeSheet !== activeRef.current) {
      predict(ack.editState.activeSheet, selectionRef.current ?? selectionAt({ row: 0, col: 0 }));
    }
    publish?.();
    if (ack.frame) {
      const frameRequest = requests.get(ack.frame.coversThroughSeq);
      if (!frameRequest || !sourceRef.current?.install(ack.frame,
        reconcile ? { ...frameRequest, navigation: navigation.current } : frameRequest)) sourceRef.current?.schedule();
    }
    for (const seq of requests.keys()) if (seq <= (sourceRef.current?.coversThroughSeq ?? 0)) requests.delete(seq);
    for (const seq of selectionCommands) if (seq <= ack.throughSeq) selectionCommands.delete(seq);
    for (const seq of selectionRevisions.keys()) if (seq <= ack.throughSeq) selectionRevisions.delete(seq);
    if (causalSelection.current !== null && causalSelection.current <= ack.throughSeq) causalSelection.current = null;
    if (ack.statuses.some((status) => status.committed && status.status === 'applied')) {
      mutation.current += 1;
      refreshSheetGeometry();
      propsRef.current.onChange?.();
    }
  };

  const pluginReveal = async (owner: PeerlessWorkbookSession, sheet: number, cell: CellAddr,
    align: XlsxRevealAlignment, live: () => boolean, version: string): Promise<XlsxPluginNavigationFailureCode | null> => {
    const revision = navigation.current;
    const painted = sourceRef.current?.painted;
    const cachedView = viewRef.current?.sheet === sheet && viewRef.current.version === version ? viewRef.current : null;
    const rect = painted?.version === version && painted.request.sheet === sheet && painted.geometry
      ? cellRect(painted.geometry, cell.row, cell.col) : null;
    const [geometry, view] = await Promise.all([
      rect && painted ? { version, rect }
        : owner.session.call.cellGeometry(sheet, cell.row, cell.col),
      cachedView ?? owner.session.call.sheetView(sheet),
    ]);
    if (ownerRef.current !== owner || !owner.current || owner.failure || !live() ||
      navigation.current !== revision || activeRef.current !== sheet) return 'unsupported';
    if (geometry.version !== version || view.version !== version || owner.session.confirmedBasis.version !== version) {
      return 'stale-version';
    }
    const scroll = scrollRef.current;
    if (!scroll) return 'layout-unavailable';
    if (viewRef.current !== view) { geometryRevision.current += 1; setView(view); }
    const scale = zoomRef.current;
    const left = scroll.scrollLeft / scale;
    const top = scroll.scrollTop / scale;
    const x = revealAxis(geometry.rect.x + (rect && painted && cell.col >= view.frozenCols ? painted.request.viewport.x : 0),
      geometry.rect.w, cell.col < view.frozenCols, left, scroll.clientWidth / scale - view.frozenWidth, align);
    const y = revealAxis(geometry.rect.y + (rect && painted && cell.row >= view.frozenRows ? painted.request.viewport.y : 0),
      geometry.rect.h, cell.row < view.frozenRows, top, scroll.clientHeight / scale - view.frozenHeight, align);
    const movedX = x !== null && x !== left;
    const movedY = y !== null && y !== top;
    if (movedX) scroll.scrollLeft = x * scale;
    if (movedY) scroll.scrollTop = y * scale;
    if (movedX || movedY || painted?.request.sheet !== sheet || painted.version !== version) sourceRef.current?.schedule();
    return null;
  };
  const pluginNavigate = async (identity: object, target: { sheetId: string; selection?: Selection; row?: number; col?: number },
    align: XlsxRevealAlignment, wantsFocus: boolean, live: () => boolean,
    expectVersion: string): Promise<XlsxPluginNavigationFailureCode | null> => {
    const owner = ownerRef.current;
    if (!owner || owner !== identity || !owner.current || owner.failure || !live()) return 'unsupported';
    const sheet = (owner.session.editState.sheets ?? owner.session.state.sheets).find((entry) => entry.id === target?.sheetId);
    const cell = target?.selection?.focus ?? { row: target?.row!, col: target?.col! };
    const valid = (value: CellAddr | undefined) => value && Number.isInteger(value.row) && Number.isInteger(value.col) &&
      value.row >= 0 && value.row < 1_048_576 && value.col >= 0 && value.col < 16_384;
    if (!sheet || !valid(cell) || target.selection && !valid(target.selection.anchor)) return 'missing-target';
    let pendingNavigation: typeof pluginNavigation.current = null;
    try {
      if (target.selection || sheet.index !== activeRef.current) {
        const selected = target.selection ? rangeTarget(sheet.index, target.selection) : undefined;
        const revision = navigation.current;
        let ticket: PeerlessWorkerCommandTicket | null = null;
        let acknowledgedVersion: string | null = null;
        pendingNavigation = { owner, sheetId: sheet.id, clear: !target.selection,
          ticket: null as PeerlessWorkerCommandTicket | null };
        pluginNavigation.current = pendingNavigation;
        const offAck = owner.session.onAck((ack) => {
          const status = ack.statuses.find((entry) => entry.seq === ticket?.seq);
          if (!status) return;
          acknowledgedVersion = ack.statuses.some((entry) => entry.seq !== ticket?.seq && entry.committed) ||
            !status.committed && ack.version !== expectVersion ? null : ack.version;
        });
        try {
          ticket = coordinator.command({ kind: 'setActiveSheet', payload: { args: [sheet.index],
            target: { sheetId: sheet.id, range: { kind: 'rowCol', start: cell, end: cell } },
            ...(selected ? { selection: selected } : {}), expectVersion } }, {
            frame: undefined,
            dispatch(write) {
              if (ownerRef.current !== owner || !owner.current || owner.failure || !live()) {
                throw new XlsxCommandAdmissionError('document-replaced');
              }
              return write();
            },
          });
          pendingNavigation.ticket = ticket;
          await ticket.result;
        } finally { offAck(); }
        if (!acknowledgedVersion || owner.session.confirmedBasis.version !== acknowledgedVersion) return 'stale-version';
        if (ownerRef.current !== owner || !owner.current || owner.failure || !live() ||
          navigation.current > revision + 1 || activeRef.current !== sheet.index ||
          target.selection && !sameSelection(selectionRef.current, target.selection)) return 'unsupported';
        if (!target.selection) setSelection(null);
        expectVersion = acknowledgedVersion;
      }
      const revision = navigation.current;
      const failure = await pluginReveal(owner, sheet.index, cell, align, live, expectVersion);
      if (failure) return failure;
      if (ownerRef.current !== owner || !owner.current || owner.retiring || owner.failure || !live() ||
        navigation.current !== revision || activeRef.current !== sheet.index) return 'unsupported';
      if (owner.session.confirmedBasis.version !== expectVersion) return 'stale-version';
      if (target.selection) {
        flushSync(() => setSelectedChart(null));
        if (ownerRef.current !== owner || !owner.current || owner.retiring || owner.failure || !live() ||
          navigation.current !== revision || activeRef.current !== sheet.index) return 'unsupported';
        if (owner.session.confirmedBasis.version !== expectVersion) return 'stale-version';
        const source = sourceRef.current;
        const painted = source?.painted;
        const request = source?.captureCommandFrame();
        const seq = pendingNavigation?.ticket?.seq;
        if (source && seq && painted?.version === expectVersion && request &&
          samePaintRequest(painted, { ...request, navigation: painted.request.navigation })) {
          source.install({ ...owner.session.commandIdentity, sheet: request.sheet, viewport: request.viewport,
            requestGeneration: request.requestGeneration!, zoom: request.zoom, dpr: request.dpr,
            version: painted.version, epoch: owner.session.confirmedBasis.epoch, sequence: painted.sequence,
            coversThroughSeq: seq, displayList: painted.displayList, mergedRanges: [...painted.mergedRanges] },
          { ...request, coversThroughSeq: seq });
        }
      }
      if (wantsFocus) focus();
      return null;
    } catch (error) {
      if (error instanceof PeerlessWorkerCommandRefusal) {
        if (error.status.reason === 'stale-version') return 'stale-version';
        if (error.status.reason === 'missing-target' || error.status.reason === 'stale-target') return 'missing-target';
      }
      if (error instanceof RangeError) return 'missing-target';
      throw error;
    } finally {
      if (pluginNavigation.current === pendingNavigation) {
        pluginNavigation.current = null;
        if (ownerRef.current === owner && owner.current && !owner.failure &&
          activeRef.current !== viewRef.current?.sheet) refreshSheetGeometry(true);
      }
    }
  };
  pluginNavigator.current = {
    selectCells: (identity, target, wantsFocus, live, version) => pluginNavigate(identity,
      structuredClone(target), 'nearest', wantsFocus, live, version),
    scrollToCell: (identity, target, align, live, version) => pluginNavigate(identity,
      structuredClone(target), align, false, live, version),
  };

  useLayoutEffect(() => {
    const cancelGestures = () => {
      if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
      nudgeTimer.current = null;
      nudge.current = null; chartDrag.current = null; dragging.current = false; clickStart.current = null;
      revealing.current = null;
    };
    pluginHost.beginLoad();
    cancelGestures();
    setRefusal(null); presentation.reset();
    setSelectedChart(null); setChartOffset(null); setFormatToken(null); setProposals([]); setStaleFor({});
    setProposalsPanelOpen(false); formatSource.current = null;
    setCatalog(run?.session.editState.sheets ?? run?.session.state.sheets ?? []);
    navigation.current += 1;
    causalSelection.current = null;
    acknowledgedSelection.current = null;
    pluginNavigation.current = null;
    if (!run) return;
    const source = new PeerlessWorkerPaintSource({
      generation: run.generation, commandIdentity: run.session.commandIdentity,
      capture, isCurrent: () => run.current && !run.failure,
      sentSequence: () => coordinator.submittedThroughSeq,
      requestFrame: async (request) => {
        if (run.ready) await coordinator.whenApplied(Math.min(request.coversThroughSeq ?? 0, coordinator.submittedThroughSeq));
        const coverage = Math.max(request.coversThroughSeq ?? 0, run.session.throughSeq);
        const frame = await run.session.call.frame(request.viewport, { sheet: request.sheet });
        return { ...frame, coversThroughSeq: 'coversThroughSeq' in frame && typeof frame.coversThroughSeq === 'number'
          ? frame.coversThroughSeq : coverage };
      },
      publish(frame) {
        const canvas = canvasRef.current;
        const context = canvas?.getContext('2d');
        if (!canvas || !context) throw new Error('Canvas context is unavailable');
        const request = frame.request;
        canvas.width = Math.round(request.width * request.dpr);
        canvas.height = Math.round(request.height * request.dpr);
        canvas.style.width = `${request.width}px`; canvas.style.height = `${request.height}px`;
        paintDisplayList(context, frame.displayList, request.dpr * request.zoom);
        canvas.dataset.workerSequence = String(frame.sequence);
        canvas.dataset.commandCoverage = String(frame.coversThroughSeq ?? 0);
        canvas.dataset.paintSource = frame.source; canvas.dataset.workerZoom = String(request.zoom);
      },
      onInstalled(frame) {
        setPainted(frame);
        const sheetId = (run.session.editState.sheets ?? run.session.state.sheets).find((sheet) => sheet.index === frame.request.sheet)?.id;
        pluginHost.presentGrid(sheetId ? { frame: frame.displayList, version: frame.version,
          sheetId, zoom: frame.request.zoom, viewport: frame.request.viewport } : null);
        if (frame.coversThroughSeq !== undefined) run.session.markPresented({
          ...run.session.commandIdentity, coversThroughSeq: frame.coversThroughSeq,
        });
        run.firstPaint();
      },
      onError: (error) => run.fail(error),
    });
    sourceRef.current = source;
    const echo = new WorkerCellEcho({ source, context: () => canvasRef.current?.getContext('2d') ?? null });
    echoRef.current = echo;
    const sheet = run.session.editState.activeSheet ?? run.session.state.activeSheet;
    setActive(sheet);
    const openView = (next: WorkbookSheetView) => {
      if (!run.current || run.failure || sourceRef.current !== source) return;
      setView(next);
      setSelection(selectionAt({ row: 0, col: 0 }));
      const scroll = scrollRef.current;
      if (scroll) {
        const content = scroll.firstElementChild as HTMLElement | null;
        if (content) {
          content.style.width = `${next.contentWidth * zoomRef.current}px`;
          content.style.height = `${next.contentHeight * zoomRef.current}px`;
        }
        scroll.scrollLeft = next.initialScrollX * zoomRef.current;
        scroll.scrollTop = next.initialScrollY * zoomRef.current;
      }
    };
    const initial = run.session.initialPaint;
    if (initial) { openView(initial.sheetView); if (!source.installInitial(initial.frame)) source.schedule(); }
    else void run.session.call.sheetView(sheet).then((next) => {
      openView(next); source.schedule();
    }).catch((error) => run.fail(error));
    run.beforeRetire = () => { pluginHost.beginLoad(); cancelGestures(); presentation.cancel(); coordinator.retire(); draftController.dispose(); source.dispose(); };
    return () => {
      cancelGestures(); presentation.cancel();
      echo.dispose(); source.dispose(); coordinator.retire(); draftController.dispose();
      if (sourceRef.current === source) sourceRef.current = null;
      if (echoRef.current === echo) echoRef.current = null;
    };
  }, [run, coordinator, draftController, capture]);

  useEffect(() => {
    if (!run?.current || run.failure) return;
    const scroll = scrollRef.current;
    if (!scroll) return;
    const schedule = () => {
      const source = sourceRef.current;
      const request = capture();
      const frame = source?.painted;
      if (frame && request && (frame.coversThroughSeq ?? -1) >= coordinator.submittedThroughSeq &&
        samePaintRequest(frame, request)) return;
      source?.schedule();
    };
    const observer = new ResizeObserver(schedule);
    observer.observe(scroll); scroll.addEventListener('scroll', schedule, { passive: true });
    schedule();
    return () => { observer.disconnect(); scroll.removeEventListener('scroll', schedule); };
  }, [run, run?.failure, zoom, capture, coordinator]);
  useEffect(() => {
    if (typeof window.matchMedia !== 'function') return;
    let media: MediaQueryList | undefined;
    const changed = () => {
      sourceRef.current?.schedule(); media?.removeEventListener('change', changed);
      media = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`);
      media.addEventListener('change', changed);
    };
    changed();
    return () => media?.removeEventListener('change', changed);
  }, []);
  useEffect(() => {
    if (!run?.ready || active === viewRef.current?.sheet) return;
    if (pluginNavigation.current?.owner === run && pluginNavigation.current.sheetId ===
      run.session.editState.sheets?.find((sheet) => sheet.index === active)?.id) return;
    refreshSheetGeometry(true);
  }, [run, active, refreshSheetGeometry]);
  const mode = useRef<{ owner: PeerlessWorkbookSession | null; readOnly: boolean }>({ owner: null, readOnly: false });
  useLayoutEffect(() => {
    if (!run?.ready || mode.current.owner === run && mode.current.readOnly === Boolean(props.readOnly)) return;
    const previous = mode.current;
    mode.current = { owner: run, readOnly: Boolean(props.readOnly) };
    if (previous.owner !== run && Boolean(run.session.editState.readOnly) === Boolean(props.readOnly)) return;
    if (props.readOnly) {
      flushNudgeRef.current(true);
      draftController.finish(); setFormatToken(null); chartDrag.current = null; nudge.current = null;
      setChartOffset(null);
    }
    void coordinator.setReadOnly(Boolean(props.readOnly)).result.catch(report);
  }, [run, run?.ready, props.readOnly, coordinator, draftController]);
  useEffect(() => {
    if (!run?.failure) return;
    echoRef.current?.dispose(); sourceRef.current?.dispose(); coordinator.fail(run.failure); draftController.dispose();
    chartDrag.current = null; nudge.current = null;
  }, [run, run?.failure, coordinator, draftController]);

  const command = (input: PeerlessWorkerCommandInput) => {
    const frame = sourceRef.current?.captureCommandFrame() ?? undefined;
    const selectionRevision = navigation.current;
    void coordinator.runAfterPendingInput(() => coordinator.command(input, { frame, selectionRevision }).result).catch(report);
  };
  const clear = () => {
    const target = selectedTarget();
    if (!editable() || !target) return;
    command({ kind: 'applyOps', payload: { toolbar: { id: 'clear', args: null, target }, selection: target } });
  };
  const clipboard = (kind: 'copy' | 'cut' | 'paste') => {
    const owner = ownerRef.current;
    const selected = selectionRef.current;
    if (!owner?.ready || owner.failure || !selected || kind !== 'copy' && !editable()) return;
    const sheet = activeRef.current;
    const at = navigation.current;
    const target = rangeTarget(sheet, selected);
    const range = normalizeRange(selected);
    if (kind === 'paste') {
      const ticket = coordinator.clipboard(() => navigator.clipboard.readText(), (text) => {
        const rows = fromTsv(text);
        const edits: CellInputEdit[] = [];
        let lastCol = 0;
        rows.forEach((row, dr) => {
          lastCol = Math.max(lastCol, row.length - 1);
          row.forEach((input, dc) => edits.push({ row: range.top + dr, col: range.left + dc, input }));
        });
        const next: Selection = { anchor: { row: range.top, col: range.left }, focus: {
          row: range.top + Math.max(0, rows.length - 1),
          col: range.left + lastCol,
        } };
        if (ownerRef.current === owner && navigation.current === at && !draftController.ownsFocus) predict(sheet, next);
        return { kind: 'editCells', payload: { args: [sheet, edits], target, selection: rangeTarget(sheet, next) } };
      });
      void ticket.result.catch(report);
      return;
    }
    const checkOwner = () => {
      if (ownerRef.current !== owner || !owner.current || owner.retiring) throw new XlsxCommandAdmissionError('document-replaced');
      if (owner.failure) throw owner.failure;
    };
    let resolve!: (blob: Blob) => void;
    let reject!: (error: unknown) => void;
    const payload = new Promise<Blob>((done, fail) => { resolve = done; reject = fail; });
    void payload.catch(() => {});
    let write: Promise<void> | undefined;
    try {
      if (typeof ClipboardItem === 'function' && typeof navigator.clipboard?.write === 'function') {
        write = navigator.clipboard.write([new ClipboardItem({ 'text/plain': payload })]);
        void write.catch(() => {});
      } else if (kind === 'cut' && typeof navigator.clipboard?.writeText !== 'function') {
        throw new WorkerInputRefusal('Cut requires clipboard write access');
      }
    } catch (error) { reject(error); report(error); return; }
    const copied = coordinator.runAfterPendingInput(async () => {
      checkOwner();
      const result = await owner.session.call.cellInputs(sheet, selectedRange(selected));
      checkOwner();
      return toTsv(result.cells);
    }, kind === 'cut' ? true : 'posted');
    const written = copied.then(async (text) => {
      checkOwner();
      if (write) { resolve(new Blob([text], { type: 'text/plain' })); await write; }
      else await navigator.clipboard.writeText(text);
      checkOwner();
    });
    if (kind === 'cut') {
      const ticket = coordinator.clipboard(() => written, () => ({ kind: 'applyOps', payload: {
        toolbar: { id: 'clear', args: null, target }, selection: target,
      } }));
      void ticket.result.catch((error) => { reject(error); report(error); });
    } else void written.catch((error) => { reject(error); report(error); });
  };
  const navigate = (sheet: number, next: Selection) => {
    if (!run?.ready || run.failure || draftController.isComposing('cell') || draftController.isComposing('formula')) return;
    draftController.commit('cell'); draftController.commit('formula');
    const target = rangeTarget(sheet, next);
    predict(sheet, next);
    reveal(sheet, next.focus);
    command({ kind: 'setActiveSheet', payload: { args: [sheet], target, selection: target } });
  };
  const moveChart = (sheet: number, id: string, dx: number, dy: number, retained = false, generation = run?.generation) => {
    if (!run?.ready || !run.current || run.retiring || generation !== run.generation || run.failure ||
      !retained && !editable() || dx === 0 && dy === 0) return;
    const target = rangeTarget(sheet, selectionRef.current ?? selectionAt({ row: 0, col: 0 }));
    command({ kind: 'moveChart', payload: { args: [sheet, id, dx, dy], target } });
    setChartOffset(null);
  };
  const flushNudge = (retained = false) => {
    const pending = nudge.current;
    nudge.current = null;
    if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    nudgeTimer.current = null;
    if (pending) moveChart(pending.sheet, pending.id, pending.dx, pending.dy, retained, pending.generation);
  };
  flushNudgeRef.current = flushNudge;
  const gestures = useRef({ moveChart, flushNudge });
  gestures.current = { moveChart, flushNudge };
  useEffect(() => {
    const stop = (event: globalThis.MouseEvent) => {
      const selecting = dragging.current;
      dragging.current = false;
      if (selecting && event.button === 0) setSelectionFinished((revision) => revision + 1);
      const drag = chartDrag.current;
      chartDrag.current = null;
      if (drag && event.button === 0) gestures.current.moveChart(drag.sheet, drag.id,
        (event.clientX - drag.x) / drag.zoom, (event.clientY - drag.y) / drag.zoom, false, drag.generation);
      else setChartOffset(null);
    };
    const cancel = () => { dragging.current = false; chartDrag.current = null; setChartOffset(null); gestures.current.flushNudge(); };
    window.addEventListener('mouseup', stop); window.addEventListener('blur', cancel);
    return () => {
      window.removeEventListener('mouseup', stop); window.removeEventListener('blur', cancel);
      if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
    };
  }, []);

  usePeerlessWorkerXlsxCommands(commands, {
    session: () => run?.session ?? null, coordinator,
    status: () => run?.ready ? 'ready' : run && !run.failure || loading ? 'loading' : 'empty',
    readOnly: () => propsRef.current.readOnly ?? false, collaborative: () => propsRef.current.collaboration !== undefined,
    generation: () => run?.generation ?? 0, mutation: () => mutation.current,
    selection: () => selectedTarget() ?? null,
    frame: () => sourceRef.current?.captureCommandFrame() ?? undefined,
    selectionRevision: () => navigation.current,
    view: () => ({ sheet: activeRef.current, selection: selectionRef.current, chartSelected: chartRef.current !== null,
      zoom: zoomRef.current, capturedFormat: null, borderStyle: borderStyle.current, borderColor: borderColor.current,
      proposals: proposalsRef.current, proposalsAvailable: isProposalsAvailable(), proposalsPanelOpen: panelRef.current,
      pngExport: run?.session.pngExport === true }),
    formatToken: () => formatTokenRef.current,
    formatRevision: () => formatRevision.current,
    setFormatToken,
    setZoom: (scale) => setZoom(scale),
    setBorderStyle: (value) => { borderStyle.current = value; }, setBorderColor: (value) => { borderColor.current = value; },
    setProposalsPanelOpen: (open) => { setProposalsPanelOpen(open); if (open) refreshProposals(); },
    markStale: (id, cells) => setStaleFor(({ [id]: _removed, ...rest }) => cells ? { ...rest, [id]: cells } : rest),
    refreshProposals, i18n: () => propsRef.current.i18n, translate: t,
    fail: (error) => run?.fail(error), afterPaint: () => coordinator.whenPresented().then(() => true, () => false),
    deliver(bytes) {
      if (propsRef.current.onSave) propsRef.current.onSave(bytes);
      else {
        const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: XLSX_MIME }));
        const anchor = document.createElement('a');
        anchor.href = url; anchor.download = propsRef.current.fileName ?? 'workbook.xlsx';
        document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
      }
    },
    deliverPng(bytes) {
      const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' }));
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `${(propsRef.current.fileName ?? 'workbook.xlsx').replace(/\.xlsx$/i, '')}.png`;
      document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
    },
    focusGrid: focus, refuse: (reason) => setRefusal({ code: reason.code as XlsxCommandFailureCode }),
  });
  useCommandShortcuts(commands, rootRef);
  const applySelectionFormat = useCallback(() => {
    const selection = selectionRef.current;
    const active = activeRef.current;
    if (!formatToken || formatTokenRef.current !== formatToken || !selection || !editable() || dragging.current ||
      applyingFormat.current !== null || selectionKey() === formatSource.current) return;
    const target = rangeTarget(active, selection);
    const owner = run;
    const revision = ++formatRevision.current;
    applyingFormat.current = revision;
    const frame = sourceRef.current?.captureCommandFrame() ?? undefined;
    const selectionRevision = navigation.current;
    void coordinator.runAfterPendingInput(() => coordinator.command({ kind: 'applyFormat',
      payload: { target, formatToken, selection: target },
    }, { frame, selectionRevision }).result).then(() => {
      if (ownerRef.current !== owner || !owner?.current || formatRevision.current !== revision) return;
      setFormatToken(null); commands.refresh();
    }).catch(report).finally(() => {
      if (applyingFormat.current === revision) applyingFormat.current = null;
    });
  }, [formatToken, run, coordinator, selectionFinished]);

  const point = (event: MouseEvent) => {
    const frame = sourceRef.current?.painted;
    const canvas = canvasRef.current;
    if (!frame || frame.request.sheet !== activeRef.current || !canvas) return { cell: null, chart: null };
    const rect = canvas.getBoundingClientRect();
    const x = (event.clientX - rect.left) / frame.request.zoom;
    const y = (event.clientY - rect.top) / frame.request.zoom;
    return { cell: frame.geometry ? cellAtPoint(frame.geometry, x, y) : null,
      chart: chartRegionAtPoint(frame.displayList.charts, x, y) };
  };
  const openCell = (seed?: string) => {
    const selected = selectionRef.current;
    if (!run?.ready || !editable() || !selected) return;
    reveal(activeRef.current, selected.focus);
    draftController.open(seed);
  };
  const onMouseDown = (event: MouseEvent) => {
    if (fromPlugin(event) || !run?.ready || run.failure || draftController.cellInput?.contains(event.target as Node) ||
      draftController.isComposing('cell') || draftController.isComposing('formula')) return;
    draftController.commit('cell'); draftController.commit('formula'); flushNudge();
    const { cell, chart } = point(event);
    clickStart.current = null;
    if (chart) {
      setSelectedChart({ id: chart.id, movable: chart.movable });
      if (editable() && chart.movable && event.button === 0) chartDrag.current = {
        generation: run.generation, sheet: activeRef.current, id: chart.id,
        zoom: sourceRef.current?.painted?.request.zoom ?? zoomRef.current,
        x: event.clientX, y: event.clientY,
      };
      focus(); event.preventDefault(); return;
    }
    setSelectedChart(null);
    if (!cell || !selectionRef.current) return;
    const next = event.shiftKey ? extendTo(selectionRef.current, cell, limits()) : selectionAt(cell);
    predict(activeRef.current, next);
    command({ kind: 'select', payload: { selection: rangeTarget(activeRef.current, next) } });
    clickStart.current = cell; dragging.current = true; focus();
  };
  const onMouseMove = (event: MouseEvent<HTMLDivElement>) => {
    if (fromPlugin(event) || run?.failure) return;
    const drag = chartDrag.current;
    if (drag) {
      setChartOffset({ x: event.clientX - drag.x, y: event.clientY - drag.y });
      event.currentTarget.style.cursor = 'move'; return;
    }
    const { cell, chart } = point(event);
    if (dragging.current && cell && selectionRef.current) {
      const next = extendTo(selectionRef.current, cell, limits());
      predict(activeRef.current, next);
      command({ kind: 'select', payload: { selection: rangeTarget(activeRef.current, next) } });
      return;
    }
    const link = cell && sourceRef.current?.painted ? hyperlinkAtCell(sourceRef.current.painted.displayList, cell.row, cell.col) : null;
    event.currentTarget.style.cursor = chart?.movable && editable() ? 'move' : chart || link ? 'pointer' : 'default';
    event.currentTarget.title = chart ? '' : link?.tooltip ?? '';
  };
  const onClick = (event: MouseEvent) => {
    if (fromPlugin(event) || run?.failure) return;
    const { cell } = point(event);
    const start = clickStart.current;
    clickStart.current = null;
    const frame = sourceRef.current?.painted?.displayList;
    if (!cell || !start || !frame || cell.row !== start.row || cell.col !== start.col) return;
    const link = hyperlinkAtCell(frame, cell.row, cell.col);
    if (!link) return;
    const href = safeExternalHyperlink(link);
    if (href) { window.open(href, '_blank', 'noopener,noreferrer'); return; }
    const sheets = run?.session.editState.sheets ?? run?.session.state.sheets ?? [];
    const destination = link.location ? parseHyperlinkLocation(link.location, sheets[activeRef.current]?.name ?? '') : null;
    if (!destination) return;
    const sheet = sheets.findIndex((entry) => entry.name.toLowerCase() === destination.sheetName.toLowerCase());
    if (sheet >= 0) navigate(sheet, selectionAt(destination));
  };
  const onKeyDown = (event: KeyboardEvent) => {
    if (fromPlugin(event) || !run?.ready || run.failure || !selectionRef.current || !viewRef.current ||
      draftController.isComposing('formula') || commandForEvent(event)) return;
    const mod = event.metaKey || event.ctrlKey;
    const draft = draftController.getSnapshot().cell;
    if (draft) {
      if (draftController.cellInput || draftController.key('cell', event)) return;
      if (event.key === 'Enter') draftController.commit('cell', event.shiftKey ? 'up' : 'down');
      else if (event.key === 'Tab') draftController.commit('cell', event.shiftKey ? 'left' : 'right');
      else if (event.key === 'Escape') draftController.cancel();
      else if (!mod && event.key === 'Backspace') draftController.change('cell', draft.value.slice(0, -1), null);
      else if (!mod && event.key.length === 1) draftController.change('cell', draft.value + event.key, null);
      else return;
      event.preventDefault(); return;
    }
    const chart = chartRef.current;
    if (chart) {
      if (event.key === 'Escape') {
        chartDrag.current = null; nudge.current = null; setChartOffset(null); setSelectedChart(null);
        event.preventDefault(); return;
      }
      const delta = !mod && chartKeys[event.key];
      if (delta && editable() && chart.movable) {
        const step = event.shiftKey ? 10 : 1;
        const base = nudge.current ?? { generation: run.generation, sheet: activeRef.current, id: chart.id, dx: 0, dy: 0 };
        nudge.current = { ...base, dx: base.dx + delta[0] * step, dy: base.dy + delta[1] * step };
        setChartOffset({ x: nudge.current.dx * zoomRef.current, y: nudge.current.dy * zoomRef.current });
        if (nudgeTimer.current !== null) clearTimeout(nudgeTimer.current);
        const owner = run;
        nudgeTimer.current = setTimeout(() => {
          if (owner.current && !owner.retiring && ownerRef.current === owner) gestures.current.flushNudge();
        }, 250); event.preventDefault();
      }
      return;
    }
    if (mod && ['c', 'x', 'v'].includes(event.key.toLowerCase())) {
      clipboard(event.key.toLowerCase() === 'c' ? 'copy' : event.key.toLowerCase() === 'x' ? 'cut' : 'paste');
      event.preventDefault(); return;
    }
    const action = selectionKeyReducer(selectionRef.current, event, limits());
    if (action.type === 'move') {
      predict(activeRef.current, action.selection);
      reveal(activeRef.current, action.selection.focus);
      command({ kind: 'select', payload: { selection: rangeTarget(activeRef.current, action.selection) } });
    } else if (action.type === 'startEdit') openCell(action.initialInput);
    else if (action.type === 'clear') clear();
    if (action.type !== 'none') event.preventDefault();
  };

  const sheets = catalog;
  const defaultToolbar = useMemo(() => <PeerlessDefaultToolbar />, []);
  const toolbar = props.showToolbar === false ? null : props.toolbar === undefined ? props.readOnly ? null : defaultToolbar : props.toolbar;
  const renderDock = (placement: DockPlacement) => pluginHost.managed ? <PluginDock host={pluginHost.host}
    placement={placement} activations={pluginHost.activations.filter((activation) => activation.plugin.panel?.placement === placement)}
    available={dockArea} /> : null;
  const painted = sourceRef.current?.painted;
  return <XlsxCommandContext.Provider value={commands.store}><EditorChromeContext.Provider value={true}>
    <PeerlessChromeContext.Provider value={true}>
    <PeerlessWorkerInputs controller={draftController} presentation={presentation} run={run} active={active}
      selectedChart={selectedChart !== null} error={Boolean(error)} readOnly={Boolean(props.readOnly)} navigation={navigation}>
      <div ref={rootRef} className={props.className} role="application" aria-label={t('editor.appLabel')}
        onKeyDownCapture={(event) => {
          if (!fromPlugin(event) && event.key === 'Escape') { setFormatToken(null); commands.refresh(); }
        }}
        style={{ position: 'relative', display: 'flex', flexDirection: 'column', width: '100%', height: '100%',
          minWidth: 0, color: '#202124', background: '#ffffff', fontFamily: 'ui-sans-serif, system-ui, sans-serif' }}>
        <PeerlessSelectionEffect store={presentation} onChange={applySelectionFormat} />
        {toolbar != null && <div data-testid="xlsx-toolbar" style={{ flex: '0 0 auto' }}>{toolbar}</div>}
        <div ref={dockAreaRef} data-testid="xlsx-workspace" style={{ position: 'relative', display: 'flex', flex: 1, minHeight: 0 }}>
          {renderDock('left')}
          <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0, minHeight: 0 }}>
          {isProposalsAvailable() && proposalsPanelOpen && <ProposalsPanel proposals={proposals} staleFor={staleFor}
            style={{ top: 4, right: 8, width: 'min(320px, calc(100% - 16px))', maxHeight: 'min(420px, calc(100% - 12px))' }} />}
          <div ref={scrollRef} data-testid="xlsx-scroll" tabIndex={0} onKeyDown={onKeyDown}
            onMouseDown={onMouseDown} onMouseMove={onMouseMove} onClick={onClick}
            onDoubleClick={(event) => {
              if (fromPlugin(event) || point(event).chart) return;
              const cell = point(event).cell;
              const frame = sourceRef.current?.painted?.displayList;
              if (cell && frame && hyperlinkAtCell(frame, cell.row, cell.col)) return;
              openCell();
            }}
            onMouseLeave={(event) => { event.currentTarget.style.cursor = 'default'; event.currentTarget.title = ''; }}
            style={{ position: 'relative', flex: 1, overflow: 'auto', minHeight: 0, outline: 'none' }}>
            <PeerlessWorkerExtent store={presentation} zoom={zoom} />
            <div style={{ position: 'sticky', top: 0, left: 0, width: 0, height: 0 }}>
              <canvas ref={canvasRef} style={{ display: 'block', position: 'absolute', top: 0, left: 0 }} />
              <PeerlessWorkerOverlay store={presentation} active={active} zoom={zoom} selectedChart={selectedChart}
                chartOffset={chartOffset} controller={draftController} run={run} disabled={Boolean(error) || Boolean(props.readOnly)} />
              <PeerlessWorkerPresence store={presentation} run={run} provider={props.collaboration?.provider}
                active={active} sheets={sheets} />
              {pluginHost.managed && <PluginOverlays host={pluginHost.host} activations={pluginHost.activations}
                layerRef={pluginHost.overlayLayerRef} width={(painted?.displayList.width ?? 0) * (painted?.request.zoom ?? zoom)}
                height={(painted?.displayList.height ?? 0) * (painted?.request.zoom ?? zoom)} />}
            </div>
          </div>
          {renderDock('bottom')}
          </div>
          {renderDock('right')}
        </div>
        <PeerlessWorkerAccessibility store={presentation} active={active} sheets={sheets} />
        {refusal && <div data-testid="xlsx-input-refusal" role="alert">
          {commandReason(refusal.code, { translate: t }).message}{refusal.detail ? `: ${refusal.detail}` : ''}
        </div>}
        {error && <div data-testid="xlsx-error" role="alert">{t('editor.openError')}: {error.message}</div>}
        {sheets.length > 0 && <div data-testid="xlsx-sheet-tabs" role="tablist" aria-label={t('editor.sheetTabsLabel')}
          style={{ display: 'flex', gap: 2, padding: '4px 6px', borderTop: '1px solid #e0e0e0', background: '#fafafa', overflowX: 'auto' }}>
          {sheets.map((sheet) => <button key={sheet.id} role="tab" aria-selected={sheet.index === active} disabled={Boolean(error)}
            onClick={() => navigate(sheet.index, selectionAt({ row: 0, col: 0 }))}
            style={{ border: 'none', padding: '4px 12px', cursor: 'pointer', background: sheet.index === active ? '#ffffff' : 'transparent',
              borderBottom: sheet.index === active ? `2px solid ${BRAND}` : '2px solid transparent', fontWeight: sheet.index === active ? 600 : 400 }}>
            {sheet.name}
          </button>)}
        </div>}
      </div>
    </PeerlessWorkerInputs>
    </PeerlessChromeContext.Provider>
  </EditorChromeContext.Provider></XlsxCommandContext.Provider>;
}
