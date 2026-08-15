---
name: superpowers-agentic-workflow
description: Superpowers software development methodology skill for coding agents - structured spec, TDD verification, DRY & YAGNI, subagent execution.
---

# Superpowers Agentic Workflow Skill

This skill enforces the **Obra Superpowers** software engineering methodology for coding agents:

## Core Principles

1. **Spec & Brainstorm First**: Never write production code before clarifying user intent, architecture constraints, and data contracts.
2. **Implementation Planning**: Create explicit `implementation_plan.md` artifacts breaking work into digestible components.
3. **Red-Green TDD Verification**: Verify syntax and runtime validity using node AST checks (`scratch/check_syntax.js`) before deploying or declaring completion.
4. **YAGNI & DRY**: Keep abstractions lean, reuse helper functions (`getSvgIcon`, `formatDisplayDate`, `escapeHtml`, `runGAS`), avoid dead code paths.
5. **Persistent Memory Retention**: Keep state synced with `NEXUS_GRAPH.md` and design tokens in `DESIGN.md`.
