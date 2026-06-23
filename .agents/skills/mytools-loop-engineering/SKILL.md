---
name: mytools-loop-engineering
description: Use for C:\ToolOfAjun MyTools development, bugfix, review, release-prep, installer packaging, or refactor tasks that should follow a durable engineering loop: inspect, plan, implement, validate, repair up to 5 times, review, and deposit reusable lessons into the right project artifact.
---

# MyTools Loop Engineering

## Purpose

Use this skill to run C:\ToolOfAjun tasks as a closed engineering loop instead of a one-pass edit. Keep the work grounded in the real repository, verify with project commands, and turn reusable lessons into future project assets.

## Workflow

1. Inspect the real state.
   - Read `AGENTS.md`, then use `docs\智能助手记忆索引.md` to choose the smallest relevant set of `docs\*.md`, `docs\规划\*.md`, `docs\场景驱动开发\*.md`, or nested `AGENTS.md`.
   - Check `git status --short --branch` before editing.
   - Read the target source files before proposing or changing code.
   - Treat existing dirty files as user work unless proven otherwise; do not revert them.

2. Make a minimal plan.
   - State the task goal, likely files, verification command, and any known scope boundary.
   - If the task touches真实使用流程 or the user uses `开qg`, follow `docs\场景驱动开发\场景驱动开发范式.md`.
   - If the task changes visible controls, sync `docs\场景驱动开发\控件交互逻辑说明.md` plus the existing `docs\功能说明.md`, `docs\程序逻辑.md`, and `docs\开发记录.txt` rules when applicable.
   - Decide the multi-agent role split before implementation. Micro tasks such as one-line text fixes, read-only answers, and trivial config checks may stay single-agent. Non-trivial code, rule, release, cross-file, or cross-module tasks default to multi-agent execution, with at least one read-only `检查/协调/验收` agent. Add execution agents only when disjoint ownership improves elapsed time or review quality; optimize for task throughput, not cost. The primary agent owns final integration, conflict resolution, and delivery.

3. Implement narrowly.
   - Prefer existing WPF views, MVVM ViewModels, Services, DPAPI/config/logging helpers, startup-performance patterns, and installer scripts.
   - Do not introduce heavyweight NuGet packages, .NET Core/high-version runtime dependencies, or alternate UI frameworks without explicit user approval.
   - Do not modify unrelated dirty files.

4. Validate through the shared entrypoint.
   - Run `powershell -ExecutionPolicy Bypass -File scripts\codex-eval.ps1 -Quick` for ordinary small tasks that need a fast sanity gate.
   - Run `powershell -ExecutionPolicy Bypass -File scripts\codex-eval.ps1` when a broad release-level check is appropriate.
   - For targeted work, pass one or more scopes: `-Build`, `-Installer`.
   - If the user asks to rebuild or repack the installer, treat it as the full installer pipeline, not just a loose EXE copy.
   - If a faster command is more appropriate, explain why and run the narrowest reliable command.

5. Repair loop.
   - If validation fails, read the exact failing log, identify the likely root cause, make the smallest safe fix, and rerun the relevant validation.
   - Count each validation-fix cycle after the first failure as one repair attempt.
   - Stop after 5 failed repair attempts if the issue is still unresolved. Report the evidence, exact commands, key logs, attempted fixes, remaining gap, and what user input or external state is needed.

6. Review before finalizing.
   - Inspect the diff for unrelated changes, leftover debug code, unused imports, broken references, missing docs, startup-performance regressions, and packaging risks.
   - Check whether each module or feature touched by the change has its current application logic recorded in the right living document (`docs\程序逻辑.md`, `docs\功能说明.md`, and `docs\场景驱动开发\控件交互逻辑说明.md` for visible controls). If behavior changed and the logic is not recorded, treat it as a review finding to fix before completion.
   - For security-sensitive, DPAPI/config, SQL, startup, installer, or large cross-module changes, do an explicit risk pass before calling the task done.
   - For any non-trivial code or rule change, spawn the project review agent `mytools-reviewer` from `.codex\agents\mytools-reviewer.toml`, wait for its result, and address every blocking finding before finalizing.
   - The review agent must be read-only: pass the task goal, relevant diff, validation commands/results, and known scope boundaries; ask it to report blocking issues, suggestions, confirmation needs, and a final `记忆候选` line (per `docs\LoopEngineering记忆沉淀.md` §五).
   - If the current Codex surface cannot spawn custom agents, run the same `mytools-reviewer` instructions as a separate read-only review pass and state that fallback in the completion report.

7. Deposit reusable lessons.
   - First merge the reviewer's `记忆候选` line with your own conclusions, then apply the three deposit questions and routing table in `docs\LoopEngineering记忆沉淀.md`; if there is no reusable lesson, do not force a memory update.
   - Add or update tests when the lesson is executable behavior.
   - Add or update `scripts\codex-eval.ps1` when the lesson is a repeatable validation gate.
   - Add or update `AGENTS.md` only for stable project-wide rules, hard boundaries, validation expectations, or required doc-sync rules.
   - Add or update `.agents\skills\*` when the lesson is a reusable workflow.
   - Add or update `docs\*.md`, `docs\规划\*.md`, or `docs\场景驱动开发\*.md` when the lesson is project knowledge, module behavior, runbook detail, or user-facing control logic.
   - For repeated failures, record the failure signature only in the routed durable location: symptom, root cause, minimal fix, validation command, and deposit location.
   - Do not create progress-summary Markdown files unless the root `AGENTS.md` explicitly requires a planning document for the task type or the user asks for one.

## Completion Report

End with:

- What changed.
- What validation ran and whether it passed.
- Any remaining risk, blocked item, or skipped validation.
- Any reusable lesson deposited, or a concise memory candidate if the user should confirm before adding a durable rule. If the three deposit questions all come back negative, state "无新增可沉淀记忆" instead of forcing a summary.
