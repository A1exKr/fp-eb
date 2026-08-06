# FP-GEN Development Plan

Purpose: single source of truth for what we build next in `fpgen_mvp/`, in what order, and what each
step depends on. Platform-side asks are tracked separately in `EXABASE_REQUESTS.md` (EB-xxx).

Last updated: 2026-08-06

---

## 1. Where we are

FP-GEN MVP is a FastAPI application with a deterministic pipeline:

```
RFP text/file -> parser -> fee_engine -> relevant_selector -> proposal_builder -> DB
                                                                              -> JSX bundle
                                                                              -> INDD export
```

- The **only** place the app talks to an LLM is `app/services/openai_service.py`
  (`json_completion`, `text_completion`), used by exactly two call sites:
  `services/parser.py` (RFP -> structured JSON) and `services/proposal_builder.py` (cover-letter polish).
- The service already switches on `LLM_PROVIDER=openai|litellm`, so an OpenAI-compatible gateway is
  supported today with configuration only.
- Fees, reference-project selection, persistence and both exports are pure server-side code. No LLM
  performs arithmetic or writes files.

## 2. Guiding constraints

1. **No personal LLM credential in production.** Inference must run on an organization-managed
   connection (EB-001).
2. **Anthropic terms:** third-party products — including agents built on the Claude Agent SDK — may not
   use claude.ai login or its rate limits. **API-key / gateway authentication is the required path.**
   Consequence: any future Agent SDK work must set *both* `ANTHROPIC_BASE_URL` **and** an explicit
   gateway credential; setting the base URL alone leaves a saved claude.ai login as the active
   credential (compliance and billing-attribution failure).
3. **Anthropic does not support routing Claude Code to non-Claude models through a gateway.** The
   current LiteLLM upstream is OpenAI, so any Agent SDK path is blocked until Claude models are
   registered on the gateway.
4. Fee math stays exact; JSX and INDD exports stay byte-comparable; RBAC and persistence are preserved.
5. RFP files are **untrusted input**. Any design that puts tools in the model's hands must be treated as
   a prompt-injection surface.

---

## 3. Roadmap

| Phase | Scope | Depends on | Status |
|-------|-------|-----------|--------|
| **0** | Platform asks: org-managed credential, gateway routes, budgets | EB-001 | Open |
| **1** | **Option C+** — managed LLM connection *and* admin-triggered regeneration on the review page | EB-001 for production; buildable now against the current provider | **4b done**; 4a blocked on EB-001 |
| **2** | Tool layer over `fee_engine` / `relevant_selector` / exporters / repositories | Phase 1 | Optional |
| **3** | Decision: Claude Agent SDK in-app (**B2**) vs exaWizards Agent Dashboard (**A-hybrid**) | EB-001 answers | Deferred |

Phases 2 and 3 are optional. Phase 1 is the committed work.

---

## 4. Phase 1 — Option C+ (committed)

Two halves, shipped together.

### 4a. Managed LLM connection (configuration only)

Set `LLM_PROVIDER=litellm`, `LITELLM_URL=http://fpgen-ai-litellm:8080`, and a **scoped virtual key**
(not the gateway master key). No application code changes. Tracked as EB-001.

### 4b. Deterministic regeneration controls on the proposal review page

Today a proposal is generated once and can only be hand-edited. Phase 1 makes the pipeline stages
re-runnable on a saved proposal, triggered by an authorised user — not by an autonomous agent.

**Prerequisite fixes**

1. **Persist `fee_input` in the proposal payload.** `build_proposal_payload` stores
   `project / client / parsed / sections / financial / relevant_experience / markdown`; the inputs that
   produced `financial` are discarded, so a fee re-run is currently impossible.
2. **Do not swallow LLM errors on the regeneration path.** The cover-letter synthesis in
   `build_proposal_payload` is wrapped in `except Exception: pass` and gated by
   `enable_openai_synthesis`. That is acceptable for batch generation but wrong for a button: a user
   clicking *Regenerate* must see failures.

**New endpoints**

| Endpoint | Behaviour |
|---|---|
| `POST /v1/proposals/{id}/experience/reselect` | `{selected_reference_ids[], auto, limit, commit}` — re-runs `select_relevant_projects`, rebuilds `relevant_experience` and `markdown` |
| `POST /v1/proposals/{id}/sections/{key}/regenerate` | `{instruction?, commit}` — rebuilds one section from the current `parsed` / `financial` / `relevant_experience`; `cover_letter` additionally goes through `text_completion` with the optional instruction |
| `POST /v1/proposals/{id}/fee/recalculate` | `{fee_input, commit}` — re-runs `calculate_fee`, updates `financial` and the dependent sections |

`commit: false` returns a **preview** without writing. This is required: regeneration would otherwise
silently overwrite manual edits made in the review editor.

**UI — `app/static/review.html`**

- A regenerate control per section header, alongside the existing inline editor.
- A *Relevant Experience* panel: checkbox list fed by the existing `GET /v1/reference-projects`, plus an
  *Auto-select* action.
- A fee panel bound to the newly persisted `fee_input`.
- A preview dialog with **Apply / Discard**.
- Controls shown only when `GET /v1/me` reports the required role, with the **same check enforced
  server-side** (`fpgen_admin` for regeneration; `Finance` also allowed for fee recalculation).

**Implementation order**

1. Persist `fee_input`; add a targeted single-section rebuild in `proposal_builder`.
2. Add the three endpoints with role gating and `commit`/preview semantics.
3. Wire the review-page controls and preview dialog.
4. In-container verification: regeneration produces changed content, fee figures still match
   `calculate_fee`, and JSX/INDD exports are unchanged for an unmodified proposal.

**Why this shape.** It delivers the iteration capability that motivated the Agent SDK option —
reselect projects, redo the cover letter, recalculate and rebuild — without adding a Node.js runtime,
without requiring Claude models on the gateway, and without giving a model tool access while untrusted
RFP text is in its context. The LLM stays a text function; the human is the loop.

### 4c. As built (2026-08-06)

Delivered on branch `fpgen/postgres-llm-auth`. Decisions taken during implementation:

| Decision | Resolution |
|---|---|
| Legacy proposals with no `fee_input` | **No back-fill.** `fee_input` is persisted from now on; older proposals show an empty fee panel with a note. Existing rows are test data. |
| LLM unavailable or failing on a regenerate click | **Silent deterministic fallback** plus a `notice` string surfaced in the review status bar and the preview dialog. Never a hard failure. |
| Preview → Apply | **Apply commits server-side immediately.** The UI warns first if there are unsaved manual edits. The preview response is echoed back on Apply (`apply_text` / `input_patch`) so no second LLM call is made. |
| Free-text instruction on deterministic sections | The instruction is converted into an **allowlisted JSON patch of the section's source inputs in `parsed`**, type-checked server-side, then the section is re-rendered deterministically. A model can never write fee arithmetic or an unlisted field. `financial` and `relevant_experience` reject instructions outright (HTTP 400) — they have dedicated endpoints. |
| Auth | All `/v1/proposals/*` now require an authenticated user (previously completely open). Section/experience regeneration requires `fpgen_admin`; fee recalculation additionally allows `Finance`. `/v1/me` gained `can_regenerate` and `can_recalculate_fee`. |

Allowlisted instruction-editable inputs, by section (`SECTION_INPUT_FIELDS` in `services/proposal_builder.py`):
`project_understanding` → `understanding.understanding`; `methodology` → `methodology.text`;
`scope_deliverables` → `scope.scopeList`, `scope.deliverablesList`;
`schedule` → `project.duration`, `schedule.totalWeeks`, `schedule.milestones`;
`team` → `team.principal.name/title`, `team.pm.name/title`;
`assumptions_exclusions` → `assumptions.defaultText`.

**Verification** (WSL Docker, `fpgen-mvp-api:local` against `fpgen-pg`):

- `docker exec fpgen-local python /tmp/smoke.py --regen` → **27/27 pass**
  (source: `fpgen_mvp/scripts/smoke_incontainer.py --regen`). Covers `fee_input` persistence and reload,
  preview not writing to the DB, out-of-allowlist patch keys being dropped, `parsed` + section + markdown
  all updating together, instruction rejection on `financial`, unknown-section 404, auto and manual
  experience reselection, and `financial == calculate_fee(fee_input)` after recalculation.
- RBAC with `FPGEN_AUTH_ENABLED=true` and forwarded-header identities:

  | Identity | GET proposal | section regen | experience | fee | export JSX |
  |---|---|---|---|---|---|
  | anonymous | 401 | 401 | 401 | 401 | 401 |
  | `fpgen_admin` | 200 | 200 | 200 | 200 | 200 |
  | `Finance` | 200 | **403** | **403** | 200 | 200 |
  | other group | 200 | **403** | **403** | **403** | 200 |

- Browser walkthrough of `review.html`: no console errors; 9 per-section Regenerate controls; reference
  checkbox list populated; fee panel hydrated from the persisted `fee_input`; a misc-reimbursables change
  previewed 80,580.00 → 86,580.00 USD and persisted correctly to `financial`, `fee_input`, the Financial
  section and the markdown.
- An untouched proposal still exports an identical JSX bundle.

**Still outstanding for production:** 4a (`LLM_PROVIDER=litellm` + scoped virtual key) remains blocked on
EB-001. Until then every regeneration takes the deterministic path and reports the "No LLM connection is
configured" notice.

### 4d. Cross-section coherence

**The dependency edges are one-directional.** Nothing in the `team` allowlist
(`principal.name/title`, `pm.name/title`) or the `schedule` allowlist (`duration`, `totalWeeks`,
`milestones`) is an input to `calculate_fee`. Editing them therefore *cannot* invalidate the financials —
a cascade in that direction would be fabricating numbers. The cascade that genuinely exists runs the other
way and is already implemented: `fee/recalculate` rebuilds `financial` + `schedule` + `team`.

```
fee_input ──calculate_fee──> financial ──> [financial, schedule] sections
fee_input ──────────────────────────────-> [team, financial] sections
parsed.project.type/location + keywords ──select_relevant_projects──> relevant_experience section
parsed.project.name/type/location + fee.rates ────────────────────-> cover_letter
parsed.team ─> team      parsed.schedule/duration/effortByPhase ─> schedule
```

`financial`, `relevant_experience`, `team` and `schedule` are **sinks**. The only real sources are
`fee_input` and `parsed.project.*`.

#### Step 1.1 — done (2026-08-06)

Prerequisites for any cascade work; auto-selection was silently degraded before these.

1. **`select_relevant_projects` read the wrong key.** It looked up `parsed["project"]["project_type"]`
   while the parser writes `project.type`, so the +5 exact-type score never fired. Fixed to read `type`
   with `project_type` as a legacy fallback, and scoring now falls back to +3 on shared type tokens
   (free-text types such as "Master Plan" vs "master plan" rarely match on exact equality).
2. **`parsed["keywords"]` was never populated**, so the keyword-intersection score was always 0. Combined
   with (1), auto-selection effectively ranked on a location substring only — i.e. near-arbitrary order.
   `parser._derive_keywords` now emits normalized tokens from project type/name/location plus the scope
   and deliverables lists, and both sides are tokenized during scoring.
3. **Two phase vocabularies could drift** (`parsed.fee.effortByPhase` from the RFP versus
   `fee_input.roles[].hours_by_phase` from the fee panel), producing a schedule that lists an uncosted
   phase beside a costed one that is not in the plan. Rather than pollute the client-facing text,
   `phase_alignment_notice` reports the mismatch to the reviewer through the existing `notice` field on
   the schedule/financial regeneration and fee recalculation responses.

Verified: smoke suite now 33/33 (`--regen`), including that auto-select ranks a type+location+keyword
match above a non-match, and that aligned phases produce no notice.

#### Step 1.2 — provenance and staleness (proposed, not built)

The rule that makes cascade safe is **classify by provenance, not by section name**:

| Origin | On upstream change |
|---|---|
| `derived` — pure deterministic render of stored data | Auto-rebuild silently; it is a projection, nothing is lost |
| `edited` — human edit via the review editor | Never overwrite; mark stale and offer Rebuild |
| `llm` — produced by `text_completion` | Never overwrite; re-running costs money, is non-deterministic and destroys applied polish |

- Add `payload["section_state"][key] = {"origin": ..., "stale_reason": ...}`.
- Declare `SECTION_SOURCES` beside the existing `SECTION_INPUT_FIELDS` so the downstream closure is
  computed rather than hand-maintained per call site.
- Split the response: derived downstream sections stay in `changed_sections`; authored ones go to a new
  `stale_sections: [{key, reason}]`.
- Review page: amber "Out of date — fee inputs changed" badge on the section header with a one-click
  Rebuild (reusing the existing regenerate endpoint), plus "Rebuild all stale" in the top bar.
- `relevant_experience` needs `selection_mode: auto | manual`. Auto results are derived data and should
  restale when `project.type` changes; an explicit checkbox selection is a human decision and must never
  be auto-replaced. `experience/reselect` already knows which mode it ran in.

Explicitly not: auto re-running the cover letter on any upstream change (stale badge only); cascading on a
plain `PUT` section save (that is the human authoring — just set `origin = "edited"`).

#### Step 1.3 — widen the instruction allowlist (after 1.2)

Add `project.type`, `project.location`, `keywords` and phase names. These are the fields that create the
first genuine cascade, which is precisely why they must wait for 1.2 — otherwise the first `project.type`
edit silently desynchronises `relevant_experience` and `cover_letter`.

---

## 5. Phase 2 — tool layer (optional, low-regret)

Wrap the existing services as callable tools: `calculate_fee`, `select_relevant_projects`,
`build_jsx_bundle`, `export_proposal_to_indd`, and proposal read/write from `repositories.py`.

The Phase 1 endpoints are already the right granularity, so this is mostly an adapter. The same
definitions serve **both** Phase 3 branches — in-process for B2, remote (MCP over HTTP) for A-hybrid —
so the work is not wasted whichever branch is chosen, or if neither is.

---

## 6. Phase 3 — options under evaluation

### B2 — Claude Agent SDK inside the FastAPI app
The agent orchestrates; our tools do the work. Adds `claude-agent-sdk` to `requirements.txt` and
**Node.js + the Claude Code CLI** to the `Dockerfile` (the SDK spawns the CLI as a subprocess).
Blocked on Claude models being available on the gateway (constraint 3).

### A-hybrid — exaWizards AI Agent Dashboard + a remote tool server
Their dashboard holds the managed LLM connection; it supports Read / Write / Bash / WebFetch /
WebSearch and registering an external MCP server URL, so our exports and fee engine can be reached
remotely. Requires the dashboard to be deployable for our tenant and able to reach a server we host.

### Retired
- **A-pure** (dashboard only): loses exports, exact fees, the database and the REST API.
- **B1** (Agent SDK for the two existing completions): a second runtime for no added capability.

### Guardrails if either is adopted
- Allowlist our tools only; **no Bash, no WebFetch/WebSearch**.
- `strict_mcp_config`, empty `setting_sources`, non-interactive permission mode.
- `max_turns`, spend cap, request timeout.
- Validate model output against a JSON schema before it reaches the exporters.
- No `~/.claude` credentials baked into the image; credential injected from managed Secrets only.

---

## 7. Open questions for the platform team

| Question | Blocks |
|---|---|
| Scoped virtual key + org-owned upstream credential | Phase 1 production (EB-001) |
| Budget / rate limits sized for **interactive** regeneration, not one call per proposal | Phase 1 (EB-001) |
| Can Claude models be registered upstream on the gateway? | B2 entirely |
| Can `/anthropic/v1/messages` be exposed? | B2 |
| Is the Agent Dashboard deployable for our tenant, and can it reach a server we host? | A-hybrid |

---

## 8. Changelog

- 2026-08-06: **Step 1.1 coherence fixes** — `select_relevant_projects` now reads `project.type` (it read
  a key the parser never writes, so the type score never fired) and scores on shared tokens; the parser
  now derives `parsed.keywords` (previously absent, so keyword scoring was always 0); phase-vocabulary
  drift between the RFP plan and the fee inputs is reported through the `notice` field. Steps 1.2
  (provenance + staleness) and 1.3 (wider allowlist) documented in §4d, not yet built. Smoke 33/33.
- 2026-08-06: Phase 1 **4b shipped** — `fee_input` persisted in the proposal payload; three regeneration
  endpoints (`sections/{key}/regenerate`, `experience/reselect`, `fee/recalculate`) with preview/commit
  semantics and role gating; review-page regeneration controls, reference-project panel, fee panel and
  preview dialog; authentication added to all `/v1/proposals/*`. Verified in WSL Docker (27/27 smoke,
  RBAC matrix, browser walkthrough). See §4c. 4a still blocked on EB-001.
- 2026-08-05: Plan created. Option C revised to **C+** (managed LLM connection plus admin-triggered
  regeneration on the review page) and committed as Phase 1. A-pure and B1 retired; B2 and A-hybrid
  deferred to Phase 3 pending platform answers.
