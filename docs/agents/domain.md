# Domain docs

## Layout

This repository uses a single domain context:

- `GLOSSARY.md` at the repository root defines domain vocabulary.
- `docs/adr/` holds architecture decision records.

## Before exploring

Read `GLOSSARY.md` and any ADRs relevant to the work, when present.
If these documents are absent, proceed silently. `/domain-modeling`,
used by `/grill-with-docs` and `/improve-codebase-architecture`, creates
them as terms and decisions are resolved.

## Use the glossary's vocabulary

Use established glossary terms in issue titles, proposals, hypotheses,
and test names. If a needed concept is missing, verify it belongs to
the domain and note the gap for `/domain-modeling`.

## Surface ADR conflicts

If a proposal contradicts an existing ADR, identify the ADR and explain
why the decision merits reopening.
