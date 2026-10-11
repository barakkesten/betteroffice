import { describe, expect, it } from 'bun:test';
import { SessionFailure, type SessionFailureCode } from '../types';
import { createWorkerEditSession } from './session';
import type {
  Checkpoint, CommandAck, CommandConnection, CommandEnvelope, RecoveryEvent,
  RecoveryHooks, WorkerEditSessionOptions,
} from './types';

type Ack = CommandAck<string, string, string>;
type Options = WorkerEditSessionOptions<string, string, string, string, string, string>;
type ReopenInput = Parameters<RecoveryHooks<string, string, string, string, string>['reopen']>[0];

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

function worker(incarnation = 0, checkpoint?: Checkpoint) {
  const edits: Array<{ envelope: CommandEnvelope<string>; reply: ReturnType<typeof deferred<Ack>> }> = [];
  const checkpoints: Array<ReturnType<typeof deferred<Checkpoint>>> = [];
  const listeners = new Set<(failure: SessionFailure) => void>();
  const removed: Array<(failure: SessionFailure) => void> = [];
  let failure: SessionFailure | undefined;
  let disposals = 0;
  const initial: Ack = {
    session: 'test', incarnation, epoch: checkpoint?.epoch ?? 1,
    version: checkpoint?.version ?? 0, throughSeq: checkpoint?.throughSeq ?? 0, outcomes: [],
    frame: { coversThroughSeq: checkpoint?.throughSeq ?? 0, revision: 0, payload: 'open' },
  };
  const connection: CommandConnection<string, string, string, string> = {
    initial,
    client: {
      call: {
        edit(envelope) {
          const reply = deferred<Ack>();
          edits.push({ envelope, reply });
          return reply.promise;
        },
        checkpoint() {
          const reply = deferred<Checkpoint>();
          checkpoints.push(reply);
          return reply.promise;
        },
      },
      get failure() { return failure; },
      onFailure(listener) {
        listeners.add(listener);
        return () => { listeners.delete(listener); removed.push(listener); };
      },
      dispose() { disposals += 1; return Promise.resolve(); },
    },
  };
  function ack(index: number, extra: Partial<Ack> = {}) {
    const seq = edits[index]!.envelope.seq;
    edits[index]!.reply.resolve({ ...initial, throughSeq: seq, version: initial.version + index + 1,
      outcomes: [{ seq, status: 'applied' }], frame: undefined, ...extra });
  }
  function fail(code: SessionFailureCode = 'crash') {
    failure = new SessionFailure(code, `worker ${code}`);
    for (const listener of [...listeners]) listener(failure);
    return failure;
  }
  return { connection, edits, checkpoints, ack, fail,
    staleFailure: () => { for (const listener of removed) listener(new SessionFailure('crash', 'obsolete')); },
    get disposals() { return disposals; } };
}

function harness(overrides: Partial<Options> = {}) {
  const first = worker();
  const queued: Array<() => void> = [];
  const reopens: Array<{ input: ReopenInput; reply: ReturnType<typeof deferred<CommandConnection<string, string, string, string>>> }> = [];
  const events: RecoveryEvent[] = [];
  const session = createWorkerEditSession({
    source: 'original', connection: first.connection,
    recovery: {
      reopen(input) {
        const reply = deferred<CommandConnection<string, string, string, string>>();
        reopens.push({ input, reply });
        return reply.promise;
      },
      checkpoint: (connection) => connection.client.call.checkpoint(),
      onRecovery: (event) => { events.push(event); },
    },
    enqueue: (callback) => { queued.push(callback); }, now: () => 0,
    requestFullFrame: (connection) => Promise.resolve(connection.initial), ...overrides,
  });
  function pump() { while (queued.length) queued.shift()!(); }
  function reopen(index: number) {
    const input = reopens[index]!.input;
    const next = worker(input.incarnation, input.checkpoint);
    reopens[index]!.reply.resolve(next.connection);
    return next;
  }
  return { first, session, reopens, events, pump, reopen };
}

async function turns() {
  for (let index = 0; index < 16; index += 1) await Promise.resolve();
}

describe('worker edit recovery', () => {
  it('cancels indexed replay descendants in bounded chunks and keeps independent history', async () => {
    const h = harness({ maxInFlight: 256, checkpointPolicy: () => false });
    const root = h.session.submit('insert', 'root');
    const children = Array.from({ length: 130 }, () => h.session.submit('insert', 'child', { after: root }));
    const independent = h.session.submit('insert', 'independent');
    h.pump();
    h.first.ack(0);
    await root.ack;
    h.first.fail();
    await turns();
    const next = h.reopen(0);
    await turns();
    next.ack(0, { version: 0, outcomes: [{ seq: 1, status: 'refused', reason: 'missing' }] });
    await turns();
    let firstCancelled = false;
    let lastCancelled = false;
    void children[0]!.ack.then(() => { firstCancelled = true; });
    void children[129]!.ack.then(() => { lastCancelled = true; });
    await Promise.resolve();
    expect(firstCancelled).toBe(true);
    expect(lastCancelled).toBe(false);
    expect(next.edits).toHaveLength(1);
    h.pump();
    await turns();
    expect(lastCancelled).toBe(true);
    expect(next.edits.map((edit) => edit.envelope.seq)).toEqual([1, 132]);
    expect(await children[129]!.ack).toMatchObject({ status: 'cancelled', reason: 'dependency-refused' });
    next.ack(1, { version: 1, frame: { coversThroughSeq: 132, revision: 1, payload: 'independent' } });
    await turns();
    expect(await independent.ack).toMatchObject({ seq: 132, status: 'applied' });
    expect(h.events.some((event) => event.phase === 'recovered')).toBe(true);
    await h.session.dispose();
  });

  it('replays exactly once after crash before ack, after ack before checkpoint, and after checkpoint', async () => {
    for (const boundary of ['before-ack', 'after-ack', 'after-checkpoint']) {
      const h = harness({ checkpointPolicy: () => boundary === 'after-checkpoint' });
      const ticket = h.session.submit('insert', 'a');
      h.pump();
      if (boundary !== 'before-ack') {
        h.first.ack(0);
        await ticket.ack;
        h.pump();
      }
      if (boundary === 'after-checkpoint') {
        h.first.checkpoints[0]!.resolve({ state: new Uint8Array([7]), epoch: 1, version: 1, throughSeq: 1 });
        await turns();
      }
      h.first.fail();
      await turns();
      const next = h.reopen(0);
      await turns();
      if (boundary === 'after-checkpoint') {
        expect(next.edits).toHaveLength(0);
        expect(h.reopens[0]!.input.checkpoint?.throughSeq).toBe(1);
      } else {
        expect(next.edits.map((edit) => edit.envelope.seq)).toEqual([1]);
        expect(next.edits[0]!.envelope.replay).toBe(true);
        next.ack(0, { frame: { coversThroughSeq: 1, revision: 1, payload: 'a' } });
        await turns();
      }
      expect(await ticket.ack).toMatchObject({ status: 'applied', seq: 1 });
      expect(h.events.some((event) => event.phase === 'recovered')).toBe(true);
      await h.session.dispose();
    }
  });

  it('continues seq across incarnations, retains original replay seqs, and dedupes duplicate delivery', async () => {
    const h = harness();
    const first = h.session.submit('local', 'a');
    const remote = h.session.submit('remote', 'b', { lane: 'collab', after: first });
    h.pump();
    h.first.ack(0);
    h.first.ack(1);
    await h.session.flush();
    h.first.fail();
    await turns();
    const next = h.reopen(0);
    const later = h.session.submit('local', 'c');
    await turns();
    expect(later.seq).toBeUndefined();
    next.ack(0, { outcomes: [{ seq: 1, status: 'applied' }, { seq: 1, status: 'duplicate' }] });
    await turns();
    next.ack(1, { frame: { coversThroughSeq: 2, revision: 2, payload: 'ab' } });
    await turns();
    h.pump();
    expect(next.edits.map((edit) => [edit.envelope.seq, edit.envelope.replay])).toEqual([[1, true], [2, true], [3, undefined]]);
    expect(next.edits[1]!.envelope.afterSeq).toBe(1);
    expect(await first.ack).toMatchObject({ status: 'applied' });
    expect(remote.seq).toBe(2);
    await h.session.dispose();
  });

  it('recovers unknown apply outcomes, excludes the thrower, and cancels descendants', async () => {
    for (const named of [true, false]) {
      const h = harness({ maxInFlight: 2 });
      const thrower = h.session.submit('insert', 'a');
      const child = h.session.submit('insert', 'b', { after: thrower });
      const unsent = h.session.submit('insert', 'c', { after: child });
      const independent = h.session.submit('insert', 'd');
      h.pump();
      const error = new Error('apply failed');
      if (named) Object.assign(error, { name: 'CommandOutcomeUnknown', refusal: { seq: 1 } });
      h.first.edits[0]!.reply.reject(error);
      await turns();
      expect(await thrower.ack).toMatchObject({ status: 'refused', reason: 'unknown-outcome' });
      expect(await child.ack).toMatchObject({ status: 'cancelled', reason: 'dependency-refused' });
      expect(await unsent.presented).toBe('dropped');
      expect(h.reopens[0]!.input.refusedSeqs).toEqual([1, 2]);
      const next = h.reopen(0);
      await turns();
      h.pump();
      expect(next.edits).toHaveLength(1);
      expect(next.edits[0]!.envelope).toMatchObject({ seq: 3, payload: 'd' });
      expect(independent.seq).toBe(3);
      await h.session.dispose();
    }
  });

  it('retries crash during replay, queues new submissions, and terminates after attempts exhaust', async () => {
    const h = harness();
    const pending = h.session.submit('insert', 'a');
    h.pump();
    h.first.fail();
    await turns();
    const next = h.reopen(0);
    const queued = h.session.submit('insert', 'b');
    await turns();
    next.fail();
    await turns();
    expect(h.reopens).toHaveLength(2);
    const final = h.reopen(1);
    await turns();
    const failure = final.fail('trap');
    await turns();
    expect(await pending.ack.catch((error) => error)).toBe(failure);
    expect(await queued.ack.catch((error) => error)).toBe(failure);
    expect(await pending.presented).toBe('dropped');
    expect(await queued.presented).toBe('dropped');
    expect(h.reopens).toHaveLength(2);
    expect(h.events[h.events.length - 1]?.phase).toBe('failed');
    await h.session.dispose();
  });

  it('requests full frames on a replacement after a replay connection fails with a pending request', async () => {
    const requests: Array<{
      connection: CommandConnection<string, string, string, string>;
      reply: ReturnType<typeof deferred<Ack>>;
    }> = [];
    const h = harness({ requestFullFrame(connection) {
      const reply = deferred<Ack>();
      requests.push({ connection, reply });
      return reply.promise;
    } });
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.first.fail();
    await turns();
    const next = h.reopen(0);
    await turns();
    next.ack(0, { frame: { coversThroughSeq: 1, revision: 1, baseRevision: 9, payload: 'a' } });
    await turns();
    expect(requests).toHaveLength(1);
    expect(next.edits).toHaveLength(2);
    next.fail();
    await turns();
    const final = h.reopen(1);
    await turns();
    final.ack(0, { frame: { coversThroughSeq: 1, revision: 1, baseRevision: 9, payload: 'a' } });
    await turns();
    expect(requests).toHaveLength(2);
    expect(requests[1]!.connection).toBe(final.connection);
    requests[0]!.reply.resolve({ ...next.connection.initial, throughSeq: 1, version: 1,
      frame: { coversThroughSeq: 1, revision: 99, payload: 'obsolete' } });
    final.ack(1, { frame: { coversThroughSeq: 2, revision: 2, baseRevision: 9, payload: 'b' } });
    await turns();
    expect(requests).toHaveLength(2);
    expect(h.session.state.presentedThroughSeq).toBe(0);
    requests[1]!.reply.resolve({ ...final.connection.initial, throughSeq: 2, version: 2,
      frame: { coversThroughSeq: 2, revision: 3, payload: 'full' } });
    expect(await first.presented).toBe('presented');
    expect(await second.presented).toBe('presented');
    expect(h.events.some((event) => event.phase === 'recovered')).toBe(true);
    await h.session.dispose();
  });

  it('cancels only the active replay call and ignores its late reply after replacement', async () => {
    const h = harness();
    const tickets = ['a', 'b', 'c'].map((payload) => h.session.submit('insert', payload));
    h.pump();
    h.first.fail();
    await turns();
    const next = h.reopen(0);
    await turns();
    next.ack(0);
    await turns();
    next.ack(1);
    await turns();
    expect(next.edits).toHaveLength(3);
    next.fail();
    await turns();
    expect(h.reopens).toHaveLength(2);
    const final = h.reopen(1);
    await turns();
    next.ack(2, { selection: 'obsolete', frame: { coversThroughSeq: 3, revision: 99, payload: 'obsolete' } });
    await turns();
    expect(h.session.state.throughSeq).toBe(0);
    expect(h.session.state.selection).toBeUndefined();
    expect(final.edits).toHaveLength(1);
    for (let index = 0; index < 3; index += 1) {
      final.ack(index, { frame: { coversThroughSeq: index + 1, revision: index + 1, payload: 'full' } });
      await turns();
    }
    expect(await Promise.all(tickets.map((ticket) => ticket.presented))).toEqual(['presented', 'presented', 'presented']);
    expect(h.events.filter((event) => event.phase === 'recovered')).toHaveLength(1);
    await h.session.dispose();
  });

  it('bypasses recorded-applied basis checks and reports replay refusal/throw divergence', async () => {
    for (const throws of [false, true]) {
      const h = harness();
      const first = h.session.submit('strict', 'a', { expectVersion: true });
      const child = h.session.submit('insert', 'b', { after: first });
      h.pump();
      h.first.ack(0);
      h.first.ack(1);
      await h.session.flush();
      h.first.fail();
      await turns();
      expect(h.reopens[0]!.input.replayAppliedSeqs).toEqual([1, 2]);
      const next = h.reopen(0);
      await turns();
      expect(next.edits[0]!.envelope).toMatchObject({ seq: 1, basis: { version: 0 }, expectVersion: true, replay: true });
      if (throws) next.edits[0]!.reply.reject(new Error('replay failed'));
      else next.ack(0, { version: 0, outcomes: [{ seq: 1, status: 'refused', reason: 'changed' }] });
      await turns();
      expect(h.events.some((event) => event.phase === 'diverged' && event.diverged.includes(1))).toBe(true);
      expect(await first.ack).toMatchObject({ status: 'applied' });
      expect(await first.presented).toBe('dropped');
      expect(await child.presented).toBe('dropped');
      expect(next.edits).toHaveLength(1);
      await h.session.dispose();
    }
  });

  it('validates checkpoints, truncates logs, preserves tombstones, and reports historyReset', async () => {
    const h = harness({ checkpointPolicy: () => true });
    const first = h.session.submit('insert', 'a');
    h.pump();
    h.first.ack(0);
    await first.ack;
    h.pump();
    h.first.checkpoints[0]!.resolve({ state: new Uint8Array([1]), epoch: 1, version: 1, throughSeq: 1 });
    await turns();
    const refused = h.session.submit('insert', 'b');
    h.pump();
    h.first.ack(1, { version: 1, outcomes: [{ seq: 2, status: 'refused', reason: 'locked' }] });
    await refused.ack;
    h.first.fail();
    await turns();
    expect(h.reopens[0]!.input.checkpoint?.state).toEqual(new Uint8Array([1]));
    expect(h.reopens[0]!.input.refusedSeqs).toEqual([2]);
    expect(h.reopens[0]!.input.replayAppliedSeqs).toEqual([]);
    const next = h.reopen(0);
    await turns();
    expect(next.edits).toHaveLength(0);
    expect(h.events.filter((event) => event.phase === 'recovered')[0]?.historyReset).toBe(true);
    await h.session.dispose();
  });

  it('rejects wrong-epoch, future-watermark, and mismatched-version checkpoints', async () => {
    for (const checkpoint of [
      { epoch: 2, version: 1, throughSeq: 1 },
      { epoch: 1, version: 1, throughSeq: 2 },
      { epoch: 1, version: 2, throughSeq: 1 },
    ]) {
      const h = harness({ checkpointPolicy: () => true });
      const ticket = h.session.submit('insert', 'a');
      h.pump();
      h.first.ack(0);
      await ticket.ack;
      h.pump();
      h.first.checkpoints[0]!.resolve({ state: new Uint8Array([1]), ...checkpoint });
      await turns();
      h.first.fail();
      await turns();
      expect(h.reopens[0]!.input.checkpoint).toBeUndefined();
      expect(h.reopens[0]!.input.replayAppliedSeqs).toEqual([1]);
      await h.session.dispose();
    }
  });

  it('retains checkpointed refusals in the epoch ledger', async () => {
    const h = harness({ checkpointPolicy: ({ applied }) => applied > 0 });
    const refused = h.session.submit('insert', 'a');
    const applied = h.session.submit('insert', 'b');
    h.pump();
    h.first.ack(0, { version: 0, outcomes: [{ seq: 1, status: 'refused', reason: 'locked' }] });
    h.first.ack(1, { version: 1 });
    await h.session.flush();
    h.pump();
    h.first.checkpoints[0]!.resolve({ state: new Uint8Array([1]), epoch: 1, version: 1, throughSeq: 2 });
    await turns();
    h.first.fail();
    await turns();
    expect(h.reopens[0]!.input.refusedSeqs).toEqual([1]);
    expect(h.reopens[0]!.input.checkpoint?.throughSeq).toBe(2);
    expect(await refused.presented).toBe('dropped');
    expect(await applied.ack).toMatchObject({ status: 'applied' });
    await h.session.dispose();
  });

  it('validates exact versions at initial and truncated checkpoint boundaries', async () => {
    for (const retained of [false, true]) {
      const h = harness({ checkpointPolicy: () => true });
      const first = h.session.submit('insert', 'a');
      h.pump();
      h.first.ack(0);
      await first.ack;
      h.pump();
      if (retained) {
        h.first.checkpoints[0]!.resolve({ state: new Uint8Array([1]), epoch: 1, version: 1, throughSeq: 1 });
        await turns();
        const second = h.session.submit('insert', 'b');
        h.pump();
        h.first.ack(1);
        await second.ack;
        h.pump();
      }
      h.first.checkpoints[retained ? 1 : 0]!.resolve({
        state: new Uint8Array([99]), epoch: 1,
        version: retained ? 2 : 1, throughSeq: retained ? 1 : 0,
      });
      await turns();
      h.first.fail();
      await turns();
      if (retained) {
        expect(h.reopens[0]!.input.checkpoint).toMatchObject({
          state: new Uint8Array([1]), version: 1, throughSeq: 1,
        });
        expect(h.reopens[0]!.input.replayAppliedSeqs).toEqual([2]);
      } else {
        expect(h.reopens[0]!.input.checkpoint).toBeUndefined();
        expect(h.reopens[0]!.input.replayAppliedSeqs).toEqual([1]);
      }
      await h.session.dispose();
    }
  });

  it('accepts a checkpoint at the exact initial boundary', async () => {
    const h = harness({ checkpointPolicy: () => true });
    const ticket = h.session.submit('insert', 'a');
    h.pump();
    h.first.ack(0);
    await ticket.ack;
    h.pump();
    h.first.checkpoints[0]!.resolve({ state: new Uint8Array([7]), epoch: 1, version: 0, throughSeq: 0 });
    await turns();
    h.first.fail();
    await turns();
    expect(h.reopens[0]!.input.checkpoint).toMatchObject({ version: 0, throughSeq: 0 });
    expect(h.reopens[0]!.input.replayAppliedSeqs).toEqual([1]);
    await h.session.dispose();
  });

  it('rejects checkpoints whose boundary version was never acknowledged', async () => {
    const h = harness({ checkpointPolicy: () => true });
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.first.ack(1, { version: 2, outcomes: [{ seq: 1, status: 'applied' }, { seq: 2, status: 'applied' }] });
    await first.ack;
    await second.ack;
    h.pump();
    h.first.checkpoints[0]!.resolve({ state: new Uint8Array([1]), epoch: 1, version: 1, throughSeq: 1 });
    await turns();
    h.first.fail();
    await turns();
    expect(h.reopens[0]!.input.checkpoint).toBeUndefined();
    expect(h.reopens[0]!.input.replayAppliedSeqs).toEqual([1, 2]);
    await h.session.dispose();
  });

  it('blocks descendants of already presented replay divergence across restarts', async () => {
    const h = harness({ checkpointPolicy: ({ applied }) => applied > 0 });
    const first = h.session.submit('insert', 'a');
    const child = h.session.submit('insert', 'b', { after: first });
    h.pump();
    h.first.ack(0);
    h.first.ack(1, { frame: { coversThroughSeq: 2, revision: 2, payload: 'ab' } });
    await child.presented;
    expect(await first.presented).toBe('presented');
    h.first.fail();
    await turns();
    const next = h.reopen(0);
    await turns();
    next.ack(0, { version: 0, outcomes: [{ seq: 1, status: 'refused', reason: 'changed' }] });
    await turns();
    expect(next.edits).toHaveLength(1);
    const independent = h.session.submit('insert', 'independent');
    h.pump();
    next.ack(1, { version: 1, frame: { coversThroughSeq: 3, revision: 3, payload: 'independent' } });
    await independent.ack;
    h.pump();
    next.checkpoints[0]!.resolve({ state: new Uint8Array([1]), epoch: 1, version: 1, throughSeq: 3 });
    await turns();
    const late = h.session.submit('insert', 'c', { after: child });
    let lateOutcome: unknown;
    void late.ack.then((outcome) => { lateOutcome = outcome; });
    await turns();
    expect(lateOutcome).toMatchObject({ status: 'cancelled', reason: 'dependency-refused' });
    next.fail();
    await turns();
    expect(h.reopens[1]!.input.refusedSeqs).toEqual([1, 2]);
    expect(h.reopens[1]!.input.checkpoint?.throughSeq).toBe(3);
    const again = h.reopen(1);
    await turns();
    expect(again.edits).toHaveLength(0);
    expect(await child.ack).toMatchObject({ status: 'applied' });
    await h.session.dispose();
  });

  it('validates boundary versions from reversed acknowledgements', async () => {
    const h = harness({ checkpointPolicy: () => true });
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.first.ack(1, { version: 2 });
    await second.ack;
    h.first.ack(0, { version: 1 });
    await first.ack;
    h.pump();
    h.first.checkpoints[0]!.resolve({ state: new Uint8Array([1]), epoch: 1, version: 0, throughSeq: 1 });
    await turns();
    h.first.fail();
    await turns();
    expect(h.reopens[0]!.input.checkpoint).toBeUndefined();
    await h.session.dispose();
  });

  it('ignores obsolete checkpoint replies and callbacks after restarting', async () => {
    const h = harness({ checkpointPolicy: () => true });
    const first = h.session.submit('insert', 'a');
    h.pump();
    h.first.ack(0);
    await first.ack;
    h.pump();
    h.first.fail();
    await turns();
    const next = h.reopen(0);
    h.first.checkpoints[0]!.resolve({ state: new Uint8Array([99]), epoch: 1, version: 1, throughSeq: 1 });
    h.first.staleFailure();
    await turns();
    expect(h.reopens).toHaveLength(1);
    expect(next.edits).toHaveLength(1);
    next.ack(0);
    await turns();
    h.first.ack(0, { version: 99, selection: 'obsolete' });
    await turns();
    expect(h.session.state.version).toBe(1);
    expect(h.session.state.selection).toBeUndefined();
    await h.session.dispose();
  });

  it('recovers eligible failure codes without re-settling promises; message/disposed do not reopen', async () => {
    for (const code of ['crash', 'trap', 'out-of-memory', 'silence', 'message', 'disposed'] as const) {
      const h = harness();
      const ticket = h.session.submit('insert', 'a');
      h.pump();
      const failure = h.first.fail(code);
      await turns();
      if (code === 'message' || code === 'disposed') {
        expect(h.reopens).toHaveLength(0);
        expect(await ticket.ack.catch((error) => error)).toBe(failure);
        expect(await ticket.presented).toBe('dropped');
      } else {
        const next = h.reopen(0);
        await turns();
        next.ack(0, { frame: { coversThroughSeq: 1, revision: 1, payload: 'a' } });
        expect(await ticket.ack).toMatchObject({ status: 'applied' });
      }
      await h.session.dispose();
    }
  });

  it('defaults checkpoint policy to 128 applied commands or one MiB of retained commands', async () => {
    for (const bytes of [false, true]) {
      const h = harness({ maxInFlight: 1, commandBytes: bytes ? () => 1024 * 1024 : undefined });
      const count = bytes ? 1 : 128;
      for (let index = 0; index < count; index += 1) {
        const ticket = h.session.submit('insert', 'a');
        h.pump();
        h.first.ack(index);
        await ticket.ack;
      }
      h.pump();
      expect(h.first.checkpoints).toHaveLength(1);
      await h.session.dispose();
    }
  });

  it('drains in-flight edits and waits for individual outcomes before checkpointing', async () => {
    const h = harness({ checkpointPolicy: () => true, maxInFlight: 2 });
    const first = h.session.submit('insert', 'a');
    const second = h.session.submit('insert', 'b');
    h.pump();
    h.first.ack(1, { version: 2 });
    await second.ack;
    h.pump();
    expect(h.first.checkpoints).toHaveLength(0);
    h.first.ack(0, { version: 1 });
    await first.ack;
    h.pump();
    expect(h.first.checkpoints).toHaveLength(1);
    h.first.checkpoints[0]!.resolve({ state: new Uint8Array([2]), epoch: 1, version: 2, throughSeq: 2 });
    await turns();
    h.first.fail();
    await turns();
    expect(h.reopens[0]!.input.checkpoint?.throughSeq).toBe(2);
    expect(h.reopens[0]!.input.replayAppliedSeqs).toEqual([]);
    await h.session.dispose();
  });

  it('accounts applied commands and bytes across outcomes and checkpoint truncation', async () => {
    const samples: Array<{ applied: number; bytes: number }> = [];
    const h = harness({
      maxInFlight: 1, commandBytes: () => 7,
      checkpointPolicy(stats) { samples.push(stats); return stats.bytes >= 21; },
    });
    for (const [index, status] of (['applied', 'refused', 'noop'] as const).entries()) {
      const ticket = h.session.submit('insert', 'a');
      h.pump();
      h.first.ack(index, { version: 1, outcomes: [{ seq: index + 1, status }] });
      await ticket.ack;
    }
    h.pump();
    expect(samples[samples.length - 1]).toEqual({ applied: 1, bytes: 21 });
    h.first.checkpoints[0]!.resolve({ state: new Uint8Array([1]), epoch: 1, version: 1, throughSeq: 3 });
    await turns();
    const next = h.session.submit('insert', 'b');
    h.pump();
    h.first.ack(3, { version: 1, outcomes: [{ seq: 4, status: 'noop' }] });
    await next.ack;
    h.pump();
    expect(samples[samples.length - 1]).toEqual({ applied: 0, bytes: 7 });
    await h.session.dispose();
  });

  it('resets document recovery data on epoch replacement', async () => {
    const old = harness({ checkpointPolicy: () => true });
    const previous = old.session.submit('insert', 'old');
    old.pump();
    old.first.ack(0);
    await previous.ack;
    old.pump();
    old.first.checkpoints[0]!.resolve({ state: new Uint8Array([7]), epoch: 1, version: 1, throughSeq: 1 });
    await turns();
    await old.session.dispose();
    const replacement = worker();
    replacement.connection.initial = { ...replacement.connection.initial, session: 'new', epoch: 2, throughSeq: 1, version: 0, frame: undefined };
    const inputs: ReopenInput[] = [];
    const next = worker(1);
    next.connection.initial = { ...replacement.connection.initial, incarnation: 1 };
    const queued: Array<() => void> = [];
    const session = createWorkerEditSession({
      source: 'replacement', connection: replacement.connection,
      recovery: { reopen: (input) => { inputs.push(input); return Promise.resolve(next.connection); } },
      requestFullFrame: () => Promise.resolve(next.connection.initial),
      enqueue: (callback) => { queued.push(callback); },
      now: () => 0,
    });
    const fresh = session.submit('insert', 'fresh');
    while (queued.length) queued.shift()!();
    expect(fresh.seq).toBe(2);
    replacement.fail();
    await turns();
    expect(inputs[0]).toMatchObject({ source: 'replacement', epoch: 2, session: 'new', checkpoint: undefined });
    expect(inputs[0]!.replayAppliedSeqs).toEqual([]);
    expect(inputs[0]!.refusedSeqs).toEqual([]);
    expect(next.edits[0]!.envelope).toMatchObject({ seq: 2, payload: 'fresh', basis: { epoch: 2, version: 0 }, replay: true });
    await session.dispose();
  });

  it('retains source and checkpoint bytes across reopen-hook mutation and failed attempts', async () => {
    const first = worker();
    const bytes = new Uint8Array([4]);
    const snapshot = new Uint8Array([8]);
    const inputs: Array<{ source: number[]; checkpoint: number[] }> = [];
    const queued: Array<() => void> = [];
    const session = createWorkerEditSession({
      source: bytes, connection: first.connection,
      recovery: {
        checkpoint: (connection) => connection.client.call.checkpoint(),
        reopen(input) {
          inputs.push({ source: [...input.source], checkpoint: [...input.checkpoint!.state] });
          input.source.fill(99);
          input.checkpoint!.state.fill(99);
          return Promise.reject(new SessionFailure('crash', 'reopen failed'));
        },
      },
      checkpointPolicy: () => true, enqueue: (callback) => { queued.push(callback); },
      now: () => 0,
      requestFullFrame: () => Promise.resolve(first.connection.initial),
    });
    bytes.fill(55);
    const ticket = session.submit('insert', 'a');
    while (queued.length) queued.shift()!();
    first.ack(0);
    await ticket.ack;
    while (queued.length) queued.shift()!();
    first.checkpoints[0]!.resolve({ state: snapshot, epoch: 1, version: 1, throughSeq: 1 });
    await turns();
    snapshot.fill(55);
    first.fail();
    await turns();
    expect(inputs).toEqual([{ source: [4], checkpoint: [8] }, { source: [4], checkpoint: [8] }]);
    expect(await ticket.ack).toMatchObject({ status: 'applied' });
    await session.dispose();
  });
});
