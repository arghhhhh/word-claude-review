# Claude Review for Microsoft Word

A Word add-in that adds a **Review Document** button. One click sends the open document to Claude for a proofread and legal review. The suggestions come back inside the document as tracked changes and margin comments, ready to accept or reject like any other redline.

It was built for litigation documents (pleadings, motions, letters) and catches:

- **Grammar, spelling, punctuation and typos**: "it's" vs "its", "compliant" for "complaint", "statue" for "statute"
- **Misphrasings and unclear sentences**
- **Content and consistency errors**: party names and roles (Plaintiff vs. Defendant, singular vs. plural), case numbers, dates and amounts, defined terms, paragraph numbering gaps, wrong cross-references, leftover template text, missing certificate of service or signature block
- **Citation problems**, such as a wrong rule number for the proposition stated

## The buttons

<img width="473" height="102" alt="claude-doc-review" src="https://github.com/user-attachments/assets/cb5725fb-8a6a-4657-ab5f-5d115a240696" />

Everything is in a **Claude** group on Word's **Home** tab:

| Button | What it does |
|---|---|
| **Review Document** | Sends the document to Claude and applies the suggestions as tracked changes + comments by "Claude Review". |
| **Accept** / **Reject** | Accepts or rejects the whole Claude suggestion at the cursor in one click, removes its comment, and jumps to the next one. |

Word's own **right-click → Accept / Reject** (and the Review tab's Accept / Reject) also resolve a whole Claude suggestion in one click. Normally Word treats a replaced word as two separate changes, a deletion and an insertion. On anyone else's tracked changes, these commands behave exactly as usual.

## Requirements

- Windows with Microsoft Word (Microsoft 365 / Word 2016 or newer, desktop version)
- [Node.js](https://nodejs.org) 18 or newer
- [Claude Code](https://claude.com/claude-code) installed and signed in (`claude` works in a terminal). Either the official installer or `npm install -g @anthropic-ai/claude-code` works. No API key is needed; it uses your Claude Code login. The add-in runs the `claude` program directly with its own settings, so shell aliases or profile shortcuts you may have set up for `claude` are not used and not needed.

## Install

1. Put this folder somewhere permanent (for example `C:\Users\<you>\ClaudeReview`). The add-in runs the scripts in this folder, so don't move or delete it after installing.
2. **Close Word completely.**
3. In PowerShell, from this folder:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\build-addin.ps1
   ```

4. Open Word. The **Claude** group appears at the left of the Home tab.

The installer builds `ClaudeReview.dotm` into `%APPDATA%\Microsoft\Word\STARTUP`, so Word loads it automatically. It fills in this folder's location and Node's location while building. During the build it briefly turns on Word's "Trust access to the VBA project object model" setting, and turns it off again when done.

**Re-run the installer** (with Word closed) after editing `ClaudeReview.bas` or moving this folder. Edits to `review-prompt.md` or `claude-review.mjs` take effect on the next click, with no reinstall.

## Uninstall

Close Word and delete `%APPDATA%\Microsoft\Word\STARTUP\ClaudeReview.dotm`.

## Using it

1. Open the document and click **Home → Claude → Review Document**, then confirm.
2. A small window shows a timer while Claude reads the document (usually under a minute). Word is busy until it finishes.
3. Suggestions appear as tracked changes, each with a comment explaining why, plus a summary comment at the top with an overall assessment.
4. Work through them with the **Accept / Reject** buttons or right-click. **One Ctrl+Z** undoes the whole review. Nothing is saved automatically.

It reviews the body, headers, footers, text boxes, footnotes and endnotes. Word doesn't allow comments inside headers and footers, so the explanations for those edits go into the summary comment.

If the document already has tracked changes, you'll get a warning first, because Claude would also see the deleted text. Accepting or rejecting them first gives a cleaner review.

## Customizing the review

`review-prompt.md` holds Claude's instructions: what to check, what to leave alone (quoted statutes, cases and contracts; accepted legal phrasing like "COMES NOW"), and the output format.

- It's set up for **Florida** practice (Florida citation style, Fla. R. Civ. P., Fla. Stat.). If you practice elsewhere, change those references to your state's.
- Edit the plain-English parts freely to add your own style rules or things to skip.
- Keep the **"Output format"** section as is, because the scripts depend on it.

To use a different Claude model, change `MODEL` at the top of `claude-review.mjs` (default `opus`).

## Privacy

Review Document sends the document's text to Anthropic through Claude Code, the same as pasting it into Claude. Make sure your firm is comfortable with that for client documents. Claude runs with all tools turned off (it can only read the text it's given and reply), and the temporary text files are deleted after each review.

## How it works

```
Word (ClaudeReview.bas)                    claude-review.mjs                    Claude Code
  export paragraphs ──[n] text──────────▶  build prompt from review-prompt.md ─▶ claude -p --tools ""
                                           match each "find" to the real text ◀─ JSON edits + summary
  apply as tracked changes + comments ◀──  shrink edits to the changed words
                                           (TSV, last-to-first order)
```

- **ClaudeReview.bas**: the Review / Accept / Reject macros, and the takeover of Word's built-in Accept/Reject commands (`ReviewAcceptChange` and friends, wired up in the ribbon XML inside `build-addin.ps1`).
- **claude-review.mjs**: runs Claude Code with no tools, checks every suggested edit against the document text (it tolerates curly vs. straight quotes), and trims each edit so only the changed words are redlined. Pure insertions or deletions (DEFENDANTS → DEFENDANT'S) become a single tracked change.
- **review-prompt.md**: Claude's review instructions.
- **build-addin.ps1**: builds the `.dotm` (macro + ribbon) and installs it. `-Target <path>` builds somewhere else, for testing.

## Tests

`tests/test-accept-reject.ps1` checks the Accept/Reject logic in a hidden copy of Word, without touching your open documents. It builds a test document with Claude suggestions plus one tracked change by another author, then checks that each click resolves exactly one suggestion and leaves the other author's change alone.

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\test-accept-reject.ps1
```

## Troubleshooting

- **"Compile error in hidden module"**: `ClaudeReview.bas` has a VBA error. Module-level `Private x As ...` lines must come before the first `Sub`/`Function`.
- **"The review failed: Could not run Claude"**: open a terminal and run `claude` to check that it's installed and signed in.
- **Buttons missing**: check that `ClaudeReview.dotm` is in `%APPDATA%\Microsoft\Word\STARTUP`, and that Word → File → Options → Add-ins → Manage: Templates lists it.
- **A suggestion couldn't be placed**: if Claude quotes text that isn't exactly in the document, that suggestion is listed in the summary comment instead of being applied.
