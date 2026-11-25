# Telegram Word Document Bot

**Status**: Planning
**Created**: 2026-11-24
**Mode**: VPS deployment; existing Python workspace

## 1. Project Overview

Extend the existing Telegram polling bot into a private, paid school-roster document workflow. Users register by sharing their own Telegram contact, submit a Word template and student records, pay 5,000 UZS per generated class Word file, and retrieve only their own generated files. The bot also posts payment receipts to a dedicated Telegram review group. There is one bot administrator and one or more separately allowlisted payment-review administrators.

The existing `bot.py`, `sinf_tayyorlash.py`, `6D.docx`, JSON data, photos, and generated files are implementation inputs to preserve and migrate where applicable. Hosting is a VPS/server, not Azure. There is no separate web frontend.

## 2. Goals & Requirements

- Retain Python, `python-telegram-bot` polling, `python-docx`, and Pillow.
- Preserve the current class-roster workflow and existing content; migrate legacy JSON to SQLite with a backup and repeatable/idempotent import. Use PostgreSQL only if measured scale or deployment needs justify it.
- Require contact sharing and verify `contact.user_id == sender.id` before enabling protected user features.
- Accept a user-uploaded Word template or sample, then ask for the class. Collect each student's record in sequence: photo, name, age, gender, phone. Keep each field attached to its student record if messages arrive in a different order; write columns consistently as name, age, photo, gender, phone.
- Charge 5,000 UZS per generated Word file for a class/school roster, regardless of student count. Show the unit price before generation and record the payment/credit and generated-file ledger.
- For purchase, display `PAYMENT_CARD_NUMBER` from protected environment configuration, ask the user to send a receipt photo, and forward the receipt and review context to the configured admin review group.
- Provide approve/reject buttons. Restrict approval actions to the single configured bot admin and explicitly configured review admins. Record reviewer, action, timestamps, and reversal history; permit an authorized reversal within five minutes.
- Provide same-bot administration: the bot admin can list users and inspect each user's submitted text, images, phone, class, and generated Word files. Review admins have only the payment-review access needed for their role.
- Store files privately with ownership metadata and non-identifying generated filenames. User retrieval commands must enforce ownership on every request.
- Define explicit audit and retention/deletion settings before launch; retention duration is unresolved and must not be silently hard-coded.

## 3. Services & Architecture

**Service**: `telegram-bot` — Python background worker using existing `python-telegram-bot` long polling. No Azure resource or web API is required for the initial VPS deployment.

**Main flow**:
1. `/start` explains the 5,000 UZS per-file price and available commands. Registration requests Telegram contact sharing and checks the contact user ID against the sender.
2. The user starts a roster, uploads a `.docx` template/sample, and selects the class. A persisted conversation state collects each student's photo, name, age, gender, and phone, associating every response with the current student and retaining record order.
3. Before generation, the bot checks available paid credit and confirms the class/file charge. It fills the selected template, maps the fields to name, age, photo, gender, phone columns, stores a private output, and sends it only to its owner.
4. `/purchase` reads the payment card from `PAYMENT_CARD_NUMBER`, requests a receipt photo, and forwards it to the review group with payment/user context and approve/reject controls. Approval/rejection updates payment state and an auditable decision history. A valid reversal is recorded as a new event within the five-minute window.
5. Same-bot admin commands enforce role checks before returning user or file data. Review-group actions independently verify the callback actor against the configured reviewer allowlist.

Use explicit conversation states and durable database records rather than relying solely on in-memory `context.user_data`, so polling restarts do not lose an in-progress roster or a pending review. Keep Telegram handlers thin; put registration, workflow, payment, document, storage, and authorization operations behind testable service functions.

## 4. Data Stores & File Handling

- **SQLite** on the VPS is the initial relational store for users, verified contact metadata, roles/configuration references, roster sessions, student records, payment requests, decisions/reversals, credits/charges, generated documents, ownership, and audit events. Use transactions for payment-credit and document-charge changes.
- Keep uploaded templates, student photos, receipt images, and generated `.docx` files in a private application data directory outside any public/static path. Store opaque random identifiers and owner/type metadata in SQLite; never derive public paths or filenames from user input.
- Migrate `data/oquvchilar.json` and referenced photos/generated files with a dry-run/report, backup, idempotent import marker, and post-migration record/file counts. Preserve `data/natija`, `data/rasmlar`, `6D.docx`, and source JSON; do not delete legacy data as part of import.
- Assign legacy records to owners only where existing `kim` metadata is trustworthy. Do not expose owner-unknown legacy outputs through user retrieval; require bot-admin handling until ownership is established.
- Use a storage abstraction so moving to PostgreSQL or private object storage later does not change bot handlers. Enforce filesystem permissions, safe temporary files, and atomic output replacement.

## 5. Interfaces & Configuration

**User commands / interactions**: `/start`, contact verification, start roster/template upload, guided student entry, review/confirm generation, `/purchase`, receipt submission, and an ownership-checked command to list/retrieve the user's own generated files. Add `/cancel` and clear recovery/help prompts for incomplete sessions.

**Bot-admin interactions**: list users; inspect a selected user's submitted text, images, contact phone, class, payments, and private generated files; manage approved operational settings and retention/deletion policy. Do not expose privileged commands to ordinary users.

**Review group**: receipt message includes only the context required for review and approve/reject buttons. Callback handling checks allowlisted reviewer ID, payment state, duplicate/stale actions, and the five-minute reversal rule before changing state.

**Protected configuration**: `BOT_TOKEN`, `PAYMENT_CARD_NUMBER`, bot-admin Telegram user ID, review-admin Telegram user ID allowlist, review-group ID, and private data directory. Obtain secrets only from VPS environment/secret configuration; never put the card number, bot token, or real Telegram IDs in source, this plan, or logs. Validate required settings at startup without printing their values.

## 6. Design System & UI

**Component Library**: Telegram Bot API native chat UI (inline keyboards, reply keyboards, and concise guided prompts); no browser component library or separate frontend.

Keep the interaction vocabulary consistent, label destructive/financial actions clearly, show the 5,000 UZS per-generated-file price at the relevant confirmation point, and provide localized, actionable error/recovery messages. Admin-only controls must be visibly distinct from user actions but remain protected by server-side role checks.

## 7. Security & Privacy

- Default-deny authorization. Validate sender, chat context, ownership, and role on every command, callback, and file retrieval; Telegram button visibility is not authorization.
- Separate the one bot-admin allowlist from the review-admin allowlist. Give review admins payment-decision access only; reserve full user-content inspection for the bot admin.
- Keep stored files private and use opaque filenames. Restrict VPS filesystem/database permissions, avoid public links, and clean temporary files after use.
- Never log phone numbers, receipt contents, card number, bot token, submitted text, or file contents. Log event IDs, actor IDs as needed for audit, action/result, and timestamps with access controls and defined retention.
- Prevent duplicate payment decisions, credit grants, charges, and document generation with database transactions and idempotency keys. Do not trust receipt images or user-provided payment claims as proof without an authorized review decision.
- Record retention/deletion policy as a pre-launch decision. Implement auditable deletion for personal content and generated files while retaining only records required by the approved audit/payment policy.

## 8. Deployment & Operations

Deploy as a single long-running Python polling worker on the chosen VPS, supervised by the platform service manager with restart-on-failure. Run one polling instance initially to avoid competing Telegram pollers and duplicate processing. Use a dedicated OS account, restricted data-directory permissions, encrypted VPS disks/backups where available, and protected environment configuration.

Add health/startup validation, structured redacted logs, database backups, migration/restore instructions, and an operational runbook for token/config rotation, bot restart, failed document generation, pending review recovery, and data restoration. Pin/test dependency versions and document Python/runtime requirements. No Azure hosting or Azure datastore is part of the initial deployment.

## 9. Implementation Sequence & Acceptance Criteria

1. Inspect and preserve current behavior; add tests around the existing template/document generation and class parsing before refactoring.
2. Add SQLite schema, transaction boundaries, repeatable legacy JSON/file migration, and migration verification.
3. Implement contact ownership verification, persisted registration/conversation flow, template upload, ordered student records, and owner-scoped file storage.
4. Implement payment requests, protected card display, review-group approve/reject callbacks, role checks, audit history, five-minute reversal, and per-file charge/credit accounting.
5. Add same-bot admin commands, private own-file retrieval, configurable audit/retention/deletion behavior, and redacted operational logging.
6. Deploy on a VPS and validate polling restart recovery, backups, restore, and review-group configuration.

**Acceptance criteria**:
- A shared contact is accepted only when its Telegram user ID matches the sender; unverified users cannot use protected workflows.
- A roster created from an uploaded template retains the user's student order and maps name, age, photo, gender, and phone into the correct columns even when messages arrive in a different field order.
- One generated roster file costs exactly 5,000 UZS, independent of student count; duplicate callbacks/retries cannot double-grant credit or double-charge.
- Only the configured bot admin can inspect full user content; only configured bot/review admins can approve/reject; an authorized decision can be reversed only within five minutes and every transition is audited.
- Users can retrieve only files whose ownership metadata matches their Telegram account.
- The migration preserves legacy source data, imports repeatably, reports unresolved ownership, and never makes owner-unknown files available to users.
- `PAYMENT_CARD_NUMBER` is read from protected environment configuration and absent from source and logs.
- Retention and deletion settings are explicit, auditable, and approved before production launch.

## 10. Assumptions & Open Decisions

- VPS provider, operating-system image, backup destination, and operational owner remain to be selected.
- The exact retention duration and legally required payment/audit record period remain undecided; settle them before launch.
- Confirm supported template/table variants and behavior for templates whose columns cannot be mapped reliably; reject ambiguous templates rather than silently placing data in the wrong cells.
- Confirm payment amount entry/verification UX. Reviewers must see the submitted amount and explicitly grant only the corresponding whole number of 5,000 UZS file credits; fractional credit is not allowed.
- Confirm handling for a payment amount that is not an exact multiple of 5,000 UZS, refunds, and chargebacks before production.
- Keep SQLite for initial VPS use; reconsider PostgreSQL only after assessing concurrent workload, backup/availability requirements, or measured scale.
