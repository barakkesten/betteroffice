import { useEffect, useRef, useSyncExternalStore } from 'react';
import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx/editor';
import type { Proposal } from '@betteroffice/pptx';
import type { PptxCommandStore } from '../commands/types';
import { useTranslation } from '../i18n';
import { PeerlessPresentationOverlay } from './PeerlessPresentationOverlay';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { peerlessFrameSource } from './peerlessFrameSource';

export interface PeerlessReviewState { open: boolean; selectedId: string | null; diff: boolean; presenting: boolean }

function Diff({ run, id, onError }: { run: PeerlessPresentation; id: string; onError(error: Error): void }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const visual = useSyncExternalStore(run.subscribeVisual, run.getVisualSnapshot, run.getVisualSnapshot);
  useEffect(() => {
    let cancelled = false;
    const index = run.client.state.slides.findIndex((slide) => slide.id === run.client.state.activeSlideId);
    void (async () => {
      const source = await peerlessFrameSource(run, index);
      try {
        const diff = await run.client.access!('layoutProposalDiffSlide', id, index);
        run.require();
        const scale = visual?.fit.scale ?? 1;
        const buffer = document.createElement('canvas');
        const context = buffer.getContext('2d');
        if (!context) throw new Error('Canvas is unavailable');
        sizeCanvasForSlide(buffer, diff.frame, 1, scale);
        await paintSlide(context, diff.frame, 1, scale, { resolveImage: source.resolveImage });
        if (!cancelled && canvas.current && run.current) {
          sizeCanvasForSlide(canvas.current, diff.frame, 1, scale);
          canvas.current.getContext('2d')?.drawImage(buffer, 0, 0);
        }
      } finally { source.release(); }
    })().catch((error: unknown) => { if (!cancelled && run.current) onError(error instanceof Error ? error : new Error(String(error))); });
    return () => { cancelled = true; };
  }, [run, id, visual, onError]);
  return <canvas ref={canvas} data-testid="pptx-proposal-diff" style={{ position: 'absolute', left: 0, top: 0 }} />;
}

export function PeerlessReview({ run, state, proposals, commands, onError, exit }: {
  run: PeerlessPresentation; state: PeerlessReviewState; proposals: Proposal[]; commands: PptxCommandStore;
  onError(error: Error): void; exit(): void;
}) {
  const { t } = useTranslation();
  return <>
    {state.open && <aside data-testid="pptx-proposals-panel" aria-label={t('proposals.title')} style={{ width: 280, overflowY: 'auto' }}>
      <strong>{t('proposals.title')}</strong>
      <button type="button" aria-label={t('proposals.close')} onClick={() => { void commands.execute('proposalsPanel', { open: false }); }}>×</button>
      {proposals.map((proposal) => <div key={proposal.id}>
        <span>{proposal.note ?? proposal.agentId}</span>
        <button type="button" onClick={() => { void commands.execute('proposalSelect', { proposalId: proposal.id }).then(() => commands.execute('proposalDiff', { enabled: true })); }}>{t('proposals.preview')}</button>
        <button type="button" disabled={proposal.staleTargets.length > 0} onClick={() => { void commands.execute('proposalAccept', { proposalId: proposal.id }); }}>{t('proposals.accept')}</button>
        <button type="button" onClick={() => { void commands.execute('proposalReject', { proposalId: proposal.id }); }}>{t('proposals.reject')}</button>
      </div>)}
    </aside>}
    {state.diff && state.selectedId && <Diff run={run} id={state.selectedId} onError={onError} />}
    {state.presenting && <PeerlessPresentationOverlay frameSource={(index) => peerlessFrameSource(run, index)}
      slideCount={run.client.state.slides.length} startIndex={run.client.state.slides.findIndex((slide) => slide.id === run.client.state.activeSlideId)}
      label={t('presentation.label')} counterLabel={(current, total) => t('presentation.slideCounter', { current, total })}
      previousLabel={t('presentation.previousSlide')} nextLabel={t('presentation.nextSlide')} exitLabel={t('presentation.exit')}
      onExit={exit} onError={(error) => onError(error instanceof Error ? error : new Error(String(error)))} />}
  </>;
}
