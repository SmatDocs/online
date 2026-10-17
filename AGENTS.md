# Agent instructions

This rule is mirrored from the SmartDocs workspace `AGENTS.md` (one level above this
repository). Keep the copies identical; when the rule changes, update the workspace
file and every repository copy (backend, frontend, PPB-Demo, online_update).

## Production impact first

Applies to every agent and every change that adds or reshapes a feature, a
background loop or worker, a queue, a stored document or table, a poll, a retry,
a dependency, or anything else on the request path. Build only what is needed,
never just because it is possible.

### Explain before building, and wait for the owner's go-ahead

Write a short impact note, local Markdown or in the chat, answering:

1. **Need.** The problem, who has it, why now, and what happens if we don't
   build it. Name the simpler alternative and why it is not enough.
2. **Latency on the user path.**
   - The time added per message or request, and the effect on time to first
     token.
   - The number of *sequential* network round trips, and where each one goes.
     From Europe, Firestore `nam5` (US) costs ~150–450 ms per round trip;
     Postgres, KeyDB and Milvus on the private network cost milliseconds.
   - Anything that blocks the event loop or holds Python's GIL beside chat.
3. **Complexity and growth.** The work per request and per background tick, as a
   function of users, chats, cameras, documents, events and history. Anything
   that rescans "everything ever" on a timer is not acceptable; query only the
   open or due items.
4. **Recurring cost, with a monthly $ estimate.**
   - Reads and writes per minute, and the bytes each operation moves (the
     document size, not just the count).
   - Egress: Firestore to Finland or Contabo is internet egress at
     ~$0.118/GiB, and 1 MB/min sustained ≈ $4.75/month.
   - LLM calls and tokens, and CPU, RAM and disk on the host it runs on.
5. **Where it runs.**
   - Finland serves users only.
   - Background, heavy or periodic work runs on the worker host (Contabo),
     or Acuon for experiments, never inside the API process beside chat.
   - Work that runs per replica must be leased so it runs once.
6. **Failure behaviour.**
   - Retries must back off and stop; a permanent error must not retry
     forever.
   - What happens on deploy, restart and failover.
   - Old or finished items must be retired, not re-queued.
7. **Data placement and replication.**
   - Keep hot-path data near compute.
   - No large blobs in documents that are read often: store big data once,
     reference it by id or hash, and keep the hot document small.
   - Postgres replication publishes an explicit list of tables between two
     writers, so a new table needs the replication step on both nodes and
     natural or UUID keys (no sequences).
8. **How we will measure it in production.** Name the log line, metric or
   probe, and record before/after numbers once it is deployed.

Then build, deploy through the pipeline, measure, and report the numbers
against the note.

### Lessons this rule comes from (October 2026)

- **Agent inbox on Firestore in the US:** about 12 sequential round trips per
  message, adding 3.7 s median (7.2 s p90) before every agent turn. On Postgres
  it takes about 20 ms. See the workspace report `reports/ai-systems/agent-turn-overhead-2026-10-16.md`.
- **Video recovery scan:** re-read every camera run ever, every 30 s, in every
  replica. A movement-map bug then retried failed monitors every 2 minutes,
  forever (600–2,000 times per run).
- **Run documents of 215 KB:** 95% was a movement map stored twice, re-read by
  every transaction. That was about 93 MB/min (~$440/month) for one camera.
- **Video loops inside the chat process:** held the GIL 50% of the time and
  quadrupled Luna latency. See
  the workspace report `reports/infrastructure/servers/firestore-and-latency-fixes-2026-10-16.md`.
