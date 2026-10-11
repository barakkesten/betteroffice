import type { PresentationHandle } from '../wasm/loader';
import type { PresentationCommandMutation } from './commandProtocol';
import { findTopLevelShape } from './commandShapes';

export function prepareCommandOperation(handle: PresentationHandle, mutation: PresentationCommandMutation): (() => unknown) | null {
  if (!mutation || !Array.isArray(mutation.args) || mutation.args.length > 4) return null;
  const { method, args } = mutation;
  const snapshot = handle.snapshot();
  if (method === 'insertSlide') {
    const [index, layout] = args;
    if (!Number.isSafeInteger(index) || index < 0 || index > snapshot.slides.length ||
      layout !== undefined && (typeof layout !== 'string' || !snapshot.slides.some((slide) => slide.layoutPartPath === layout))) return null;
    return () => handle.insertSlide(index, layout);
  }
  const slide = snapshot.slides.find((slide) => slide.id === args[0]);
  if (!slide) return null;
  if (method === 'deleteSlide') return () => handle.deleteSlide(slide.id);
  if (method === 'moveSlide') {
    const index = args[1];
    return Number.isSafeInteger(index) && index >= 0 && index < snapshot.slides.length ? () => handle.moveSlide(slide.id, index) : null;
  }
  const shape = findTopLevelShape(slide, args[1]);
  if (!shape || shape.id !== args[1]) return null;
  switch (method) {
    case 'setShapeFill': {
      const color = args[2];
      return shape.kind === 'shape' && (color === null || /^#[0-9a-f]{6}$/i.test(color)) ? () => handle.setShapeFill(slide.id, shape.id, color) : null;
    }
    case 'setShapeStroke': {
      const stroke = args[2];
      return shape.kind === 'shape' && stroke && (stroke.color === undefined || typeof stroke.color === 'string' && /^#[0-9a-f]{6}$/i.test(stroke.color)) &&
        (stroke.widthPt === undefined || Number.isFinite(stroke.widthPt) && stroke.widthPt >= 0 && stroke.widthPt <= 1584)
        ? () => handle.setShapeStroke(slide.id, shape.id, stroke) : null;
    }
    case 'setShapeAdjust': {
      const adjustments = args[2];
      return shape.kind === 'shape' && adjustments && typeof adjustments === 'object' && !Array.isArray(adjustments) &&
        Object.keys(adjustments).length <= 64 && Object.entries(adjustments).every(([key, value]) =>
          /^[a-zA-Z][a-zA-Z0-9]*$/.test(key) && Number.isFinite(value) && Math.abs(value) <= 1e6)
        ? () => handle.setShapeAdjust(slide.id, shape.id, adjustments) : null;
    }
    case 'removeShape': return () => handle.removeShape(slide.id, shape.id);
    case 'bringShapeToFront': return () => handle.bringShapeToFront(slide.id, shape.id);
    case 'sendShapeToBack': return () => handle.sendShapeToBack(slide.id, shape.id);
    case 'bringShapeForward': return () => handle.bringShapeForward(slide.id, shape.id);
    case 'sendShapeBackward': return () => handle.sendShapeBackward(slide.id, shape.id);
    default: return null;
  }
}
