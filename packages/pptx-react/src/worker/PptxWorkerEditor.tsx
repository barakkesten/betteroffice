import type { PEERLESS_FLAG, PptxEditorProps } from '../PptxEditor';
import { PptxEditorContent } from '../PptxEditor';
import type { PptxWorkerEditorApi } from './createWorkerEditorApi';
import { useEditableSessionPresentation } from './useEditableSessionPresentation';

/** @experimental */
export type PptxWorkerEditorProps = Omit<PptxEditorProps, 'onReady' | 'readOnly' | typeof PEERLESS_FLAG> & {
  experimentalWorkerOpen: true;
  readOnly?: false;
  onReady?: (api: PptxWorkerEditorApi) => void;
};

export function PptxWorkerEditor(props: PptxWorkerEditorProps) {
  const backend = useEditableSessionPresentation(props);
  const { onReady, ...content } = props;
  return <PptxEditorContent {...content} backend={backend} onWorkerReady={onReady} />;
}
