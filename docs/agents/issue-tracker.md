# Issue tracker: GitHub

Issues and specs live in GitHub Issues for `naufalahmadfauz/PGN-Sawala`.
Use the `gh` CLI from this clone. To select the repository explicitly,
pass `--repo naufalahmadfauz/PGN-Sawala`.

## Operations

- Publish a spec or ticket: `gh issue create --title "..." --body-file <file>`.
- Read a ticket: `gh issue view <number> --json number,title,body,labels,comments,state,assignees,url`.
- List issues: `gh issue list --state open --json number,title,body,labels,comments,assignees`. Add label/state filters as needed.
- Comment: `gh issue comment <number> --body-file <file>`.
- Apply/remove labels: `gh issue edit <number> --add-label "..."` or `--remove-label "..."`.
- Close: `gh issue close <number> --comment "..."`.

For multiline text, use a body file or `--body-file -` with a quoted heredoc.
When a skill says "publish to the issue tracker", create a GitHub issue.
When it says "fetch the relevant ticket", read the issue including its comments and labels.

## Pull requests as a triage surface

**PRs as a request surface: no.**

GitHub shares a number space for issues and PRs. When a bare reference
is ambiguous, try `gh pr view <number>` and fall back to `gh issue view <number>`.

## Task graphs and wayfinding

- Map: one issue labelled `wayfinder:map`, containing Notes, Decisions-so-far, and Fog.
- Children: link tickets as native GitHub sub-issues. If unavailable, use a task list in the map and `Part of #<map>` in each child.
- Wayfinder ticket types: `wayfinder:research`, `wayfinder:prototype`, `wayfinder:grilling`, or `wayfinder:task`.
- Blocking: use native GitHub issue dependencies through `gh api` or supported `gh issue edit` relationship flags. If unavailable, put `Blocked by: #<number>, ...` at the top of the ticket.
- Ready frontier: open, unassigned children whose blockers are all closed; select the first in map order.
- Claim: `gh issue edit <number> --add-assignee @me` as the session's first tracker write.
- Resolve: comment with the answer, close the ticket, then add a summary and ticket link to the map's Decisions-so-far.
