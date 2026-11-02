---
name: docx-three-way-merge
description: Three-way merge for Word (.docx) files with a conflict check. Use it when edits were made on the wrong copy or an old version and must be carried over to the right one, including added, deleted and changed comments. The user supplies three files: the original copy, the edited copy, and the correct file that should receive the edits. Also use it for 改錯檔案、改錯版本、把修改同步到正確的母檔、合併兩個 Word 版本、把註解同步過去, even if the user never says "merge" or "three-way".
compatibility: Python 3 with lxml. Word is optional, for checking the result.
---

# Three-way merge for Word files

Carry the changes between two versions of a document over to a third version, and stop to ask wherever the same words were changed on both sides. This is the same idea as a git merge, for files git cannot merge.

## The three files

Identify these before doing anything. Confirm with the user if the file names leave any doubt, because swapping two of them replays the edits backwards.

| Name | Meaning | Example wording |
|---|---|---|
| `BASE` | The copy before the edits | "改錯母檔", "the wrong file before I edited it" |
| `MODIFIED` | The same copy after the edits | "改後的改錯母檔", "after my edits" |
| `TARGET` | The correct file that should receive them | "正確的母檔", "the real one" |

Only the difference between `BASE` and `MODIFIED` moves. Anything else that differs in `TARGET` is the user's newer work and stays. A difference that exists only between `BASE` and `TARGET` is not a change to carry over, even if it looks like a mistake; mention it instead.

## Workflow

### 1. Check

```bash
python scripts/docx_merge.py check BASE MODIFIED TARGET --report report.json
```

Exit code 0 means clean. Exit code 2 means there are conflicts or manual items. Nothing is written in this step.

Read `report.json`:

- `edit_list`: changes that will be applied, as `base` and `new`
- `comments_removed`, `comments_added`, `comments_already_gone`
- `both_made_same_change`: both sides already agree, so nothing to do
- `conflicts`: the same words changed differently on both sides
- `manual`: things the script will not attempt, such as a new image
- `warnings`: changes outside its scope, such as headers and tables

### 2. Show the user the conflicts

Show each conflict with its context and the base, modified and target text, then let the user decide. Do not choose silently, even when one option looks obviously right: these are the user's words, and the point of the check is that they get to see each collision.

The object character `￼` in the report stands for an image or equation. When one side has `≥ 40 mm` as text and the other has `￼`, the same content probably exists in two forms, which is worth saying when you show that conflict.

If the user has already said how to settle a type of conflict, apply that and list what you decided.

### 3. Merge

Write the decisions to `resolutions.json`, keyed by conflict id:

```json
{
  "P35-1": "use_modified",
  "P35-2": "keep_target",
  "P35-4": {"text": "the wording the user chose"}
}
```

```bash
python scripts/docx_merge.py merge BASE MODIFIED TARGET -o OUT.docx --resolutions resolutions.json --report merged.json
```

The script refuses to write while a conflict or manual item is open, refuses to overwrite an input file, and needs `--force` to replace an existing output. It re-reads its output and discards it if an untouched paragraph changed or the comment markers do not balance. `--keep-target-on-conflict` keeps the target's wording in every conflict, so use it only when the user asked for that.

### 4. Verify and report

Open the output in Word if it is available, and render it to PDF to look at the changed pages. Then tell the user:

- what was applied, in plain language, and which comments were added or removed
- every conflict and how it was settled
- every `manual` item and `warning`, since these were not synced
- a change that looks like a typo in the user's edits: apply it as written and point it out, rather than fixing it

## What it does not sync

The script handles paragraph text and comments in the document body. It reports these rather than syncing them, so handle them by hand or tell the user:

- paragraph and character formatting changes
- new images, equations, tables or table rows
- headers, footers, footnotes and endnotes
- tracked changes in `MODIFIED`, which it treats as already accepted

## When the formatting looks wrong afterwards

The merge keeps the formatting `TARGET` already had, so strange fonts or line spacing usually come from `TARGET` itself. A file that passed through Google Docs is a common cause. It often changes fonts and line rules, turns equations into images, removes roman page numbers, and replaces automatic headers with fixed text. Compare `styles.xml` and the paragraph properties with a good copy to confirm this before changing anything, and treat that repair as a separate step that the user asks for.
