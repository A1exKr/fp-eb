# exaBase Request Backlog

Purpose: Track all requests we need to send to the exaBase platform team, so we avoid repeated ad-hoc asks during development.

Application-side roadmap and architecture decisions live in `DEVELOPMENT_PLAN.md`.

Last updated: 2026-08-05

## Status legend
- **TODO** — Not requested yet
- **SENT** — Requested, waiting response
- **DONE** — Completed by exaBase team
- **BLOCKED** — Cannot proceed until this is resolved

## Current requests

| ID | Status | Priority | Request |
|--------|--------|----------|---------|
| EB-001 | TODO | High | Replace personal LLM API key with an org-managed secret |
| EB-002 | TODO | High | Enable private image pull for the GHCR image |
| EB-003 | TODO | High | Register a GHCR pull credential (minimum-scope PAT) |
| EB-004 | TODO | High | Provide a tenant ECR repo + push/pull procedure (fallback) |
| EB-005 | TODO | Medium | Clarify the private registry policy for SideApp |
| EB-006 | TODO | Medium | Provision proper Keycloak accounts (or confirm self-manage) |

## Request details

### EB-001 — Option C+: org-managed LLM connection through LiteLLM
- **Status:** TODO · **Priority:** High
- **Owner:** You → exaBase admin
- **Selected architecture:** Option C+ — keep FP-GEN's existing OpenAI-compatible completion calls, route them through the `fpgen-ai-litellm` gateway, and add user-triggered regeneration controls on the proposal review page (see `DEVELOPMENT_PLAN.md`, Phase 1). The platform-side ask is unchanged; only the usage profile changes.
- **Request:** Replace the personal upstream LLM API key with an organization-managed credential stored only on the LiteLLM gateway, and issue FP-GEN a scoped LiteLLM virtual key.
- **Why:** FP-GEN already supports LiteLLM, so this removes the personal credential without changing the proposal pipeline, database, fee calculations, or exports.
- **Current state:**
  - The LiteLLM SideApp receives `OPENAI_API_KEY` as its upstream provider credential.
  - The FP-GEN API currently receives `LITELLM_MASTER_KEY`; production should use a restricted virtual key instead of the gateway administrator key.
  - The FP-GEN API does not need the raw upstream provider key when `LLM_PROVIDER=litellm`.
- **Required outcome:**
  - Register an organization-owned upstream model credential only on `fpgen-ai-litellm`.
  - Generate and register a scoped virtual key for the FP-GEN API, preferably as `FPGEN_LITELLM_API_KEY`; until the application variable is renamed, inject that virtual-key secret into the existing `LITELLM_MASTER_KEY` environment slot without supplying the real master/admin key.
  - Restrict the virtual key to approved model aliases, budgets, and rate limits — sized for **interactive** use, not one call per proposal (see usage profile below).
  - Confirm the internal gateway base URL and OpenAI-compatible `/v1/chat/completions` route.
  - Confirm whether the gateway can also expose Anthropic-compatible `/anthropic/v1/messages`, **and whether Claude models can be registered upstream** — both are prerequisites for a possible future Claude Agent SDK integration (Phase 3/B2). Anthropic does not support routing the Agent SDK to non-Claude models through a gateway, and forbids claude.ai-login authentication for third-party products, so that path would use the same scoped-key model as this request.
  - Provide credential ownership, rotation, revocation, audit logging, cost attribution, and incident-response procedures.
- **Usage profile (changed by Option C+):** proposal generation issues 1 JSON parse call plus 1 cover-letter call. The new review-page controls let an authorised user re-run cover-letter generation on demand, so per-proposal call volume becomes bursty and user-driven. Rate limits and spend budgets must accommodate repeated regeneration rather than a single pass. All other regeneration actions (reference-project reselection, fee recalculation, section rebuilds) are deterministic server-side code and consume no tokens.

**EB-001 — ready-to-send Option C message to ExaWizards (copy/paste):**

```text
- Request ID: EB-001
- Environment/tenant: exaBase Studio — FP-GEN workspaces (`fpgen-app`, `fpgen-admin`,
  and `fpgen-ai`)
- Component: LiteLLM Gateway (`fpgen-ai-litellm`) + managed Secrets
- Selected approach: Option C+ — retain FP-GEN's existing OpenAI-compatible completion
  calls and route them through LiteLLM. We are not requesting an Agent Dashboard or
  Agent SDK migration for this production step. FP-GEN is additionally adding
  user-triggered regeneration controls on its proposal review page; these reuse the
  same two completion calls, so no new route or protocol is required, but they make
  call volume user-driven rather than one pass per proposal.
- Current situation:
    1) The LiteLLM SideApp currently requires a personal `OPENAI_API_KEY` as its
       upstream provider credential.
    2) The FP-GEN API currently receives `LITELLM_MASTER_KEY`, which is an
       administrator credential and is broader than the application needs.
    3) With `LLM_PROVIDER=litellm`, the FP-GEN API does not need direct access to
       the upstream provider key.
- Needed outcome:
    1) Provision an organization-owned LLM provider credential and register it only
       on `fpgen-ai-litellm`. If ExaWizards cannot provide the upstream credential,
       please identify the approved organization/service-account ownership model.
    2) Issue a scoped LiteLLM virtual key for FP-GEN and register it in the managed
       secret store (suggested secret name: `FPGEN_LITELLM_API_KEY`). For compatibility
       with the current application, that secret may be injected into the existing
       `LITELLM_MASTER_KEY` environment slot, but its value must be the scoped virtual
       key—not the real LiteLLM master/admin key.
    3) Restrict the virtual key to the approved FP-GEN model alias(es), with agreed
       rate limits and spend budget. Please size these for interactive use: reviewers
       can re-run cover-letter generation on demand from the proposal review page, so
       token consumption is bursty rather than a fixed one pass per proposal.
    4) Confirm the internal LiteLLM base URL and support for the OpenAI-compatible
       `/v1/chat/completions` route used by FP-GEN.
    5) For a possible future Claude Agent SDK integration, confirm (a) whether
       `/anthropic/v1/messages` (or the equivalent Anthropic-format route) can be
       enabled, and (b) whether Anthropic Claude models can be registered as upstream
       models on the gateway. Anthropic does not support routing the Agent SDK to
       non-Claude models through a gateway, and prohibits claude.ai-login
       authentication for third-party products, so that path would authenticate with
       the same kind of scoped gateway key requested here. This is informational only
       and is not required for the current step.
    6) Confirm rotation and revocation procedures, audit-log availability, cost
       attribution, retention, and the operational owner for gateway incidents.
    7) Confirm that credential rotation can be completed through managed Secrets
       without rebuilding the FP-GEN container image.
- Desired deadline: <fill in>
- Security constraints:
    - No personal LLM credential in production.
    - The upstream provider credential must remain gateway-only.
    - FP-GEN must receive a least-privilege virtual key, not `LITELLM_MASTER_KEY`.
    - Secrets must not be stored in the canvas JSON, injected files, source control,
      application logs, or browser-visible configuration.
- Acceptance test:
    - FP-GEN can complete RFP JSON parsing and cover-letter generation through
      LiteLLM using only its scoped virtual key.
    - On-demand cover-letter regeneration from the proposal review page succeeds
      repeatedly for the same proposal without hitting the rate limit.
    - Removing `OPENAI_API_KEY` and the actual LiteLLM master/admin credential from
      the FP-GEN API does not break those workflows; only the scoped virtual key is
      available to the application.
    - Fee calculations, persistence, and document exports remain unchanged.
- Relevant deployment values:
    - `LLM_PROVIDER=litellm`
    - `LITELLM_URL=http://fpgen-ai-litellm:8080`
    - application credential: scoped virtual key supplied by ExaWizards
```

### EB-002 — Private image pull (GHCR)
- **Status:** TODO · **Priority:** High
- **Owner:** You → exaBase platform team
- **Request:** Enable private container image pull for `ghcr.io/a1exkr/fpgen-mvp-api:latest`.
- **Why:** Deployment fails with a 401 on image pull from the exaBase runtime.
- **Notes:** Likely requires cluster-level registry credentials.

### EB-003 — GHCR pull credential (PAT)
- **Status:** TODO · **Priority:** High
- **Owner:** exaBase platform team
- **Request:** If GHCR private pull is supported, register a GHCR credential using a PAT with minimum scope.
- **Why:** Needed to keep the image private while allowing exaBase to pull it.
- **Notes:** PAT should be read-only for packages and managed as a platform secret.

### EB-004 — ECR fallback route
- **Status:** TODO · **Priority:** High
- **Owner:** exaBase platform team
- **Request:** If GHCR is not supported, provide a tenant ECR repository and push/pull procedure.
- **Why:** Official docs are ECR-oriented; this is the fallback/standard route.
- **Notes:** Include URI format, account/region, and the credential handoff process.

### EB-005 — Private registry policy
- **Status:** TODO · **Priority:** Medium
- **Owner:** exaBase platform team
- **Request:** Clarify the supported private registry policy for SideApp (GHCR/ECR/others).
- **Why:** Avoid repeated back-and-forth for future services.
- **Notes:** Request a written policy and the preferred pattern.

### EB-006 — Proper Keycloak accounts
- **Status:** TODO · **Priority:** Medium
- **Owner:** You → exaBase admin
- **Request:** Provision proper Keycloak user account(s) for our team in the `fpgen` realm — or confirm we should self-manage via the Keycloak admin console.
- **Why:** We temporarily disabled the nginx OIDC gate to reach FP-GEN because only insecure seed/bootstrap accounts exist; real use needs proper SSO logins before auth is re-enabled.
- **Notes:**
  - Keycloak is self-hosted in our canvas (`fpgen-auth-keycloak`), realm `fpgen`, OIDC client `fpgen-oidc`.
  - Groups: Proposal Managers / Reviewers / Finance.
  - Roles: `fpgen_admin` / `fpgen_proposal_manager` / `fpgen_reviewer` / `fpgen_viewer`.
  - Need named account(s) with secure passwords, plus rotation of the default bootstrap admin and the seeded user shipped in the canvas import.
  - 2026-07-30: Full OIDC-gate blueprint prepared and validated at `FP-GEN_exaBase_blueprint_oidc.yaml`
    (oauth2-proxy + self-hosted Keycloak realm `fpgen` / client `fpgen-oidc` / groups mapper + nginx
    `auth_request`; api `FPGEN_AUTH_ENABLED=true`). Ready to deploy once this request delivers real
    accounts and the required secrets (KEYCLOAK_OIDC_CLIENT_SECRET, OAUTH2_COOKIE_SECRET,
    KC_BOOTSTRAP_ADMIN_PASSWORD) plus the APP_URL / AUTH_URL variables.

**EB-006 — ready-to-send message to ExaWizards (copy/paste):**

```text
- Request ID: EB-006
- Environment/tenant: exaBase Studio — workspaces fpgen-app + fpgen-auth (Keycloak realm `fpgen`)
- Component: Keycloak (Identity Provider) + oauth2-proxy + Secrets + Variables
- Current error/log: With FPGEN_AUTH_ENABLED=true and no OIDC front-door deployed, /admin
  returns {"detail":"Not authenticated"} (bare 401, no login page). Only insecure bootstrap
  Keycloak accounts currently exist.
- Needed outcome:
    1) Confirm whether ExaWizards provisions accounts or we self-manage (owner of the
       master-realm admin for fpgen-auth-keycloak).
    2) Real named users in realm `fpgen` with secure passwords, assigned to groups
       fpgen_admin / Finance (optionally Proposal Managers / Reviewers).
    3) Register secrets: KEYCLOAK_OIDC_CLIENT_SECRET, OAUTH2_COOKIE_SECRET,
       KC_BOOTSTRAP_ADMIN_PASSWORD (non-default); confirm DB_USER / DB_PASSWORD /
       OPENAI_API_KEY handling.
    4) Set variables APP_URL and AUTH_URL (each ending in '/').
    5) Confirm the fpgen-auth endpoint is publicly reachable for browser login, and that
       cross-workspace DNS is allowed (oauth2-proxy -> fpgen-auth-keycloak:8080,
       keycloak -> fpgen-app-rdb:8080).
    6) Confirm the OIDC `groups` claim is permitted (or provide the standard claim name).
- Desired deadline: <fill in>
- Security constraints: Rotate the bootstrap master-admin and delete the seeded placeholder
  user (fpgen-admin) after first boot; keep all secrets in the managed secret store; keep the
  GHCR image private (see EB-002/EB-003).
- Attachments: FP-GEN_exaBase_blueprint_oidc.yaml
```

## Consolidated cover message to ExaWizards (copy/paste)

Use this as the single message that opens all six items; attach or link the per-request
details above (EB-001 and EB-006 have their own ready-to-send bodies).

```text
Subject: exaBase Studio — platform requests for our FP-GEN application (6 items)

Hello,

We are building an internal web application ("FP-GEN") that runs on exaBase Studio.
In short: it takes an RFP document as input, uses an LLM plus deterministic
server-side logic to assemble a proposal (text sections, staffing, fees), and lets a
reviewer check and re-generate parts of it in the browser before exporting the final
document. Functionally it is a normal containerized web app: an API SideApp, a
database, a LiteLLM gateway SideApp, and (for authenticated access) Keycloak +
oauth2-proxy. No FP-GEN-specific business logic is needed on your side.

We have six platform-side requests. They fall into three groups.

A. LLM credentials (blocking production use)
   EB-001 — Today our LiteLLM gateway (`fpgen-ai-litellm`) runs on a personal LLM
   provider API key, and our API is given the LiteLLM master key. We need an
   organization-managed replacement:
     - an organization-owned LLM provider credential for the gateway,
     - an application credential our API can use to call the gateway,
     - the list of models available to us through the gateway. We are not tied to a
       specific provider or model, so please tell us which ones are approved for our
       use and what the cost implications are.
     - rotation / revocation / audit-log / cost-attribution procedures, and
       confirmation that rotation can be done via managed Secrets without rebuilding
       our image.
   Please size the rate limit for interactive use: reviewers can re-trigger
   generation on demand from the review screen, so usage is bursty rather than one
   fixed call per document.
   We are also evaluating the Claude Agent SDK. Please tell us whether Anthropic
   Claude models can be registered upstream on the gateway, and whether an
   Anthropic-format route such as /anthropic/v1/messages can be enabled.

B. Private container image distribution (blocking deployment)
   Our API image is currently hosted on GitHub Container Registry
   and the SideApp image pull fails with HTTP 401 unless the image privacy is set to public.
   From the Studio documentation we understand that (1) the SideApp Image field only
   accepts an image URI with no pull-credential option, and (2) managed Secrets are
   usable only as environment variables, so they cannot authenticate an image pull.
   We therefore ask:
     EB-002 — Can the platform pull private images at all, and if so how should we
              request it?
     EB-003 — If private GHCR pull is supported, please register a cluster-level GHCR
              pull credential (we will supply a read-only, packages-scoped PAT).
     EB-004 — If GHCR is not supported, please provide a tenant ECR repository plus
              the push/pull procedure (repository URI format, account/region, and how
              the push credentials are handed to us). We are happy to treat ECR as the
              standard route.
     EB-005 — Please share the written policy for private registries with SideApp
              (which registries are supported and which pattern you recommend), so we
              do not have to re-ask for future services.

C. Identity / SSO accounts (blocking re-enabling authentication)
   EB-006 — We have prepared, but not deployed, an OIDC front door (oauth2-proxy +
   self-hosted Keycloak, realm `fpgen`, client `fpgen-oidc`). Only insecure bootstrap
   accounts exist today, so authentication is currently disabled in our deployment.
   We need:
     - confirmation of whether ExaWizards provisions accounts for the realm or we
       self-manage it (i.e. who owns the Keycloak master-realm admin),
     - real named user accounts with secure passwords, assigned to our groups,
     - registration of the required secrets (OIDC client secret, cookie secret,
       non-default Keycloak bootstrap admin password) and the APP_URL / AUTH_URL
       variables,
     - confirmation that the auth endpoint is publicly reachable for browser login and
       that cross-workspace DNS between our components is permitted,
     - confirmation that an OIDC `groups` claim is allowed (or the standard claim name
       we should use).

Priorities: A and B are high (they block production deployment); C is medium but is
required before we can turn authentication back on.

Please let us know the expected turnaround for each group, and whether
any of these should be raised through a different channel.

Thank you,
<your name / team>
```

**Japanese version (日本語版・コピペ用):**

```text
件名: exaBase Studio — 弊社アプリケーション「FP-GEN」に関する基盤側のご依頼（6件）

お世話になっております。

弊社では exaBase Studio 上で社内向け Web アプリケーション「FP-GEN」を開発しています。
概要としては、RFP（提案依頼書）を入力すると、LLM とサーバー側の決定論的なロジックを
組み合わせて提案書（本文・体制・費用）を組み立て、レビュー担当者がブラウザ上で内容を
確認し、必要な箇所を再生成したうえで最終ドキュメントを出力する、というものです。
構成としては一般的なコンテナ型 Web アプリで、API SideApp、データベース、LiteLLM
ゲートウェイ SideApp、および認証用の Keycloak + oauth2-proxy から成ります。
FP-GEN 固有の業務ロジックを御社側でご対応いただく必要はございません。

基盤側へのご依頼は全6件で、内容は次の3グループに分かれます。

A. LLM 認証情報（本番利用のブロッカー）
   EB-001 — 現在、LiteLLM ゲートウェイ（`fpgen-ai-litellm`）は個人の LLM プロバイダー
   API キーで稼働しており、API 側には LiteLLM のマスターキーが渡っています。組織管理の
   構成への置き換えとして、以下をお願いします。
     - ゲートウェイ用の、組織所有の LLM プロバイダー認証情報
     - 弊社 API がゲートウェイを呼び出すためのアプリケーション用認証情報
     - ゲートウェイ経由で利用できるモデルの一覧。特定のプロバイダー・モデルにこだわりは
       ないため、利用を認めていただけるモデルと、そのコスト面の考慮事項をご教示ください。
     - ローテーション／失効／監査ログ／コスト按分の運用手順のご提示、および、コンテナ
       イメージを再ビルドせず managed Secrets 経由でローテーションできることの確認
   レート制限は「対話利用」を前提にご設定ください。レビュー画面から任意の
   タイミングで再生成を実行できるため、1件あたり固定回数ではなくバースト的な利用に
   なります。
   あわせて Claude Agent SDK の採用も検討しております。(a) Anthropic Claude モデルを
   ゲートウェイの上流モデルとして登録可能か、(b) /anthropic/v1/messages 相当の
   Anthropic 形式ルートを有効化できるか、ご教示ください。

B. プライベートコンテナイメージの配布（デプロイのブロッカー）
   弊社 API のイメージは現在 GitHub Container Registry に置いていますが、
   イメージを public にしない限り SideApp のイメージ取得が HTTP 401 で失敗します。
   Studio のドキュメントを確認したところ、
   (1) SideApp の Image 欄はイメージ URI のみでプルクレデンシャルの指定項目がない、
   (2) managed Secrets は Environment variables でのみ利用可能なためイメージ取得の
   認証には使えない、と理解しています。つきましては以下をお願いします。
     EB-002 — そもそも基盤側でプライベートイメージの取得は可能でしょうか。可能な場合、
              どのように申請すればよいかご教示ください。
     EB-003 — GHCR のプライベート取得に対応可能な場合、クラスタレベルの GHCR プル
              クレデンシャルのご登録をお願いします（packages 読み取り専用の最小権限
              PAT を弊社より提供します）。
     EB-004 — GHCR に対応できない場合、テナント用 ECR リポジトリと push/pull 手順
              （リポジトリ URI 形式、アカウント／リージョン、push 用認証情報の受け渡し
              方法）をご提供ください。ECR を標準ルートとすることに異存はありません。
     EB-005 — 今後のサービス追加のたびに同じ確認が発生しないよう、SideApp における
              プライベートレジストリの運用方針（対応レジストリと推奨パターン）を
              文書でご共有ください。

C. 認証・SSO アカウント（認証再有効化のブロッカー）
   EB-006 — OIDC の入口（oauth2-proxy + セルフホスト Keycloak、realm `fpgen`、
   client `fpgen-oidc`）は準備済みですが未デプロイです。現状は初期ブートストラップ
   アカウントしか存在せず安全でないため、認証を無効化した状態で運用しています。
   以下をお願いします。
     - realm のアカウントを御社にて払い出していただけるのか、弊社で自己管理するのか
       の確認（= Keycloak master realm 管理者の所有者はどちらか）
     - 実在ユーザー用の正式アカウント（安全なパスワード付き）の発行と、所定グループへの
       割り当て
     - 必要なシークレット（OIDC クライアントシークレット、Cookie シークレット、
       既定値でない Keycloak ブートストラップ管理者パスワード）の登録、および
       APP_URL / AUTH_URL 変数の設定
     - ブラウザログインのため認証エンドポイントが外部から到達可能であること、および
       各コンポーネント間のワークスペース跨ぎ DNS が許可されていることの確認
     - OIDC の `groups` クレームが利用可能かどうか（または標準として使用すべき
       クレーム名）のご教示

優先度: A と B は高（本番デプロイのブロッカー）、C は中ですが認証を再度有効化する前に
必要となります。

各グループの対応目安時期と、別の窓口・フローで起票すべきものがあればご教示ください。

以上、よろしくお願いいたします。
<担当者名／チーム名>
```

## Request template (copy/paste)

- Request ID:
- Environment/tenant:
- Component (SideApp/Function/Secret/etc):
- Current error/log:
- Needed outcome:
- Desired deadline:
- Security constraints:
- Attachments:

## Changelog
- 2026-07-10: Initial backlog created with current high-priority asks (LLM key replacement, private image pull credentials, PAT/ECR policy).
- 2026-07-16: Removed the FP-GEN nginx OIDC auth gate in the canvas import (login no longer required for the MVP); added EB-006 to provision proper Keycloak accounts before re-enabling auth.
- 2026-07-16: Reformatted the backlog for readability (compact summary table + per-request detail sections).
- 2026-07-30: Implemented app-side auth gating (home + admin pages, `/v1/me` with `can_setup`) and a
  top-right account menu; verified in-container (22/22). Prepared the full OIDC-gate deployment blueprint
  (`FP-GEN_exaBase_blueprint_oidc.yaml`); deployment blocked on EB-006 (Keycloak accounts + secrets).
- 2026-07-30: Added a copy/paste-ready EB-006 message (accounts, secrets, variables, networking) for ExaWizards.
- 2026-07-30: Expanded EB-001 with the selected Option C LiteLLM architecture, least-privilege
  virtual-key requirements, acceptance tests, and a copy/paste-ready request for ExaWizards.
- 2026-08-05: Added `DEVELOPMENT_PLAN.md` (application roadmap). Revised EB-001 to Option C+ — same
  platform ask, but budgets/rate limits must now cover user-triggered regeneration from the review
  page, and the Agent SDK follow-up question was extended to cover registering Claude models upstream.
  No new EB request needed: the Agent Dashboard (Option A-hybrid) stays unrequested until Phase 3.
- 2026-08-05: Added a consolidated cover message to ExaWizards covering all six requests (EB-001 …
  EB-006), grouped as LLM credentials / private image distribution / SSO accounts, with a
  non-detailed FP-GEN description. A Japanese version of the same message was added alongside it.
- 2026-08-05: Simplified group A of the cover message (EN + JA) — state what we need rather than
  prescribing how ExaWizards should implement it, and drop self-limiting qualifiers.
- 2026-08-05: Dropped the gateway base-URL / `/v1/chat/completions` confirmation from the cover
  message — already known and working (`LITELLM_URL=http://fpgen-ai-litellm:8080`, LiteLLM is
  OpenAI-compatible). Replaced it with an ask for the list of models we may use and their cost
  implications, since we are model/provider agnostic.
- 2026-08-05: Re-synced the Japanese cover message with the revised English one (GHCR wording now
  "fails unless the image is set to public", rate limit only, and the attachment sentence removed).
