# FP-GEN Development Plan

Purpose: single source of truth for what we build next in `fpgen_mvp/`, in what order, and what each
step depends on. Platform-side asks are tracked separately in `EXABASE_REQUESTS.md` (EB-xxx).

Last updated: 2026-08-05

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

| Phase | Scope | Depends on |
|-------|-------|-----------|
| **0** | Platform asks: org-managed credential, gateway routes, budgets | EB-001 |
| **1** | **Option C+** — managed LLM connection *and* admin-triggered regeneration on the review page | EB-001 for production; buildable now against the current provider |
| **2** | Tool layer over `fee_engine` / `relevant_selector` / exporters / repositories | Phase 1 |
| **3** | Decision: Claude Agent SDK in-app (**B2**) vs exaWizards Agent Dashboard (**A-hybrid**) | EB-001 answers |

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

- 2026-08-05: Plan created. Option C revised to **C+** (managed LLM connection plus admin-triggered
  regeneration on the review page) and committed as Phase 1. A-pure and B1 retired; B2 and A-hybrid
  deferred to Phase 3 pending platform answers.
