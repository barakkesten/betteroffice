import { useMemo, useRef } from 'react';
import {
  validReservationsPayload,
  type DocxReservationsPayload,
} from '@betteroffice/docx/yrs';
import type { DocxPluginActivation, DocxPluginHost } from './createDocxPluginHost';
import type { DocxReservationSet } from './types';

interface Evaluated {
  definition: DocxPluginActivation['plugin'];
  generate: DocxPluginActivation['plugin']['reservations'];
  state: unknown;
  version: string;
  mode: string;
  readOnly: boolean;
  set: DocxReservationSet | null;
}

/** The set a plugin returned, refused as `invalid-payload` unless well formed. */
function checkedSet(value: unknown): DocxReservationSet | null {
  if (value === null || value === undefined) return null;
  const set = value as DocxReservationSet;
  if (
    typeof set !== 'object' ||
    !validReservationsPayload({ version: set.version, sets: [{ owner: 'plugin', items: set.items }] })
  ) {
    throw new TypeError('invalid-payload');
  }
  return set;
}

/**
 * Every plugin's reservations at the document version plugins see, as one payload. Plugins are
 * asked again only when their state, the version or the write mode changed, so a new layout
 * alone never re-evaluates them. Returns the previous payload while a plugin's set is at another
 * version, and null while no plugin reserves anything.
 */
export function usePluginReservations(
  host: DocxPluginHost,
  activations: readonly DocxPluginActivation[]
): DocxReservationsPayload | null {
  const evaluated = useRef(new WeakMap<object, Evaluated>());
  const previous = useRef<{ json: string; payload: DocxReservationsPayload | null }>({
    json: 'null',
    payload: null,
  });
  return useMemo(() => {
    const version = host.version();
    const sets: { owner: string; items: DocxReservationSet['items'] }[] = [];
    let defined = false;
    for (const activation of activations) {
      const generate = activation.plugin.reservations;
      if (!generate) continue;
      defined = true;
      const { snapshot } = activation.context;
      const state = activation.context.state;
      let entry = evaluated.current.get(activation.activation);
      if (
        !entry ||
        entry.definition !== activation.plugin ||
        entry.generate !== generate ||
        entry.state !== state ||
        entry.version !== snapshot.version ||
        entry.mode !== snapshot.mode ||
        entry.readOnly !== snapshot.readOnly
      ) {
        entry = {
          definition: activation.plugin,
          generate,
          state,
          version: snapshot.version,
          mode: snapshot.mode,
          readOnly: snapshot.readOnly,
          set: host.guard(
            activation.pluginId,
            'reservations',
            (context) => checkedSet(generate.call(activation.plugin, context)),
            null
          ),
        };
        evaluated.current.set(activation.activation, entry);
      }
      if (!entry.set) continue;
      if (entry.set.version !== version) return previous.current.payload;
      sets.push({ owner: activation.pluginId, items: entry.set.items });
    }
    const payload = defined ? { version, sets } : null;
    const json = JSON.stringify(payload);
    if (json !== previous.current.json) previous.current = { json, payload };
    return previous.current.payload;
  }, [host, activations]);
}
