import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { PptxEditorContent, type PptxAsyncEditorProps, type PptxEditorApi } from './PptxEditor';
import type { PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';
import { createInThreadAsyncEditorApi, asyncCollaboration } from './createInThreadAsyncEditorApi';

export function PptxAsyncInThreadSession(props: PptxAsyncEditorProps) {
  const callbacks = useRef(props);
  callbacks.current = props;
  const [api, setApi] = useState<PptxAsyncEditorApi | null>(null);
  const readOnly = useRef(props.readOnly);
  readOnly.current = props.readOnly;
  const ready = useCallback((value: PptxEditorApi) => {
    setApi(createInThreadAsyncEditorApi(value, () => Boolean(readOnly.current)));
  }, []);
  const collaboration = useMemo(() => asyncCollaboration(props.collaboration), [props.collaboration]);
  useEffect(() => {
    if (!api) return;
    const cleanup = callbacks.current.onReady?.(api);
    return typeof cleanup === 'function' ? cleanup : undefined;
  }, [api]);
  return <PptxEditorContent {...props} collaboration={collaboration} onReady={ready} />;
}
