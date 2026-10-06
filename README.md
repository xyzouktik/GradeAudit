# GradeAudit

A small Google Apps Script that reads a Google Doc, finds every grading entry
(person + rating 1–10 + who graded them), keeps only the entries **you** graded,
and emails you the list.

Names are never hardcoded — they are scanned straight from the document.

## How it works

- **Tables** with a header row, e.g. `Person | Rating | Graded by`. The script
  auto-detects which column is the name, the rating and the grader.
- **Inline lines**, e.g. `Alice - 8 - graded by john@example.com` or
  `Bob: 7/10 (by John Doe)`.
- An entry is only reported if its "graded by" value matches you (your email
  addresses, your contacts' names, and any aliases you list).

## Setup

1. Go to https://script.google.com and create a **New project**.
2. Paste `GradeAudit.gs` over the default `Code.gs`.
3. Set `DOC_URL_OR_ID` (and `EMAIL` if you want the mail somewhere else) at the
   top of the file. Optionally add `MY_ALIASES`.
4. Run `emailGradedList` once and accept the permission prompt.

If the script is bound to the document, it uses the active document and no URL
is needed. A **Grade Checker → Email list graded by me** menu item is added on
open.

## Configuration

| Variable | Purpose |
| --- | --- |
| `DOC_URL_OR_ID` | Google Doc URL or ID to scan. Leave blank to use the bound document. |
| `EMAIL` | Override the recipient. Defaults to you. |
| `MIN_RATING` / `MAX_RATING` | Accepted rating range (default 1–10). |
| `MY_ALIASES` | Extra emails/names that count as "you". |
