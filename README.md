# SWordBook

**A WordPad for books.** A small, distraction-free word processor for drafting
novels, stories and e-books.

SWordBook doesn't try to be Word. It gives you the handful of things a book
draft needs, chapter titles, dialogue, scene breaks and a navigator, plus
something most word processors don't have: **three ways to see your
manuscript**, from printed pages to one endless sheet. It saves plain `.docx` files, so when the draft is ready you can
keep polishing it in Word, Google Docs or LibreOffice Writer.

> **Status: alpha (v0.1.0).** It works and saves standard `.docx` files.
> While you write, SWordBook keeps an automatic backup of your unsaved
> changes and offers to recover them if it ever closes unexpectedly. That
> backup is not a version history, though: it is deleted when you save. As
> with any early release, keep your own copies of manuscripts that matter.

## ★ Three views, one manuscript

The view mode is SWordBook's signature feature. The same book can be shown
in three ways, and you can switch at any moment in **Settings → View**.
Switching never touches the text or its formatting, so there is nothing to
undo and nothing to lose.

| View | What you see | Best for |
| --- | --- | --- |
| **Pages** | Separate sheets with the paper size and margins you choose, exactly as they will print | Checking length, layout and page breaks |
| **No pages** | One continuous sheet with no cuts between pages, at the same width as the printed page | Writing in flow, without page breaks interrupting a scene |
| **Mixed** | A continuous sheet that splits only where *you* put a page break (Ctrl+Enter). Each sheet is as long as its content | Writing chapter by chapter: each chapter is its own sheet, however long it is |

**Mixed view** is the one you won't find elsewhere. Put a page break before
each chapter and the manuscript turns into a stack of chapters, each on its
own sheet: long chapters are long sheets, short chapters are short ones, and
nothing is ever cut mid-paragraph. You still see where the book is divided
without the constant interruptions of fixed pages.

All three views lay out lines at the real page width, so line lengths always
match the printed result. The view you choose is saved in the `.docx`, and it
reopens the same way. Word and Writer open No pages and Mixed documents in
their web layout.

Zoom goes from 50 % to 200 % (Ctrl + mouse wheel, or the slider in the
status bar), with *Fit to width* and *Whole page*.

## Why SWordBook

- **A view for every stage.** Pages to review, No pages to write in flow,
  Mixed to work chapter by chapter.
- **Made for drafts.** Open it and write. Book styles are one click (or one
  shortcut) away, and there is nothing else in the way.
- **Book styles built in.** Each style matches what e-book readers expect,
  so headings show up in the table of contents.
- **Your files stay yours.** Everything is a standard `.docx`. There is no
  account, cloud or proprietary format.
- **Portable.** No installation. On Windows and Linux the whole program,
  settings included, fits in one folder you can carry on a USB drive.
- **It won't lose your work.** While there are unsaved changes, SWordBook
  keeps a backup copy, and if it closes unexpectedly it offers to recover it.

## Features

**Styles**

| Style | Shortcut | Use |
| --- | --- | --- |
| Body | Ctrl+1 | Normal text |
| Title | Ctrl+2 | Chapter title |
| Subtitle | Ctrl+3 | Chapter subtitle |
| Section | Ctrl+4 | Section heading inside a chapter |
| Thoughts | Ctrl+5 | A character's inner voice (italics) |
| Dialogue | Ctrl+6 | Starts a line with the dialogue dash (—) |
| Quote | Ctrl+7 | Letters, epigraphs, quoted texts |
| Scene break | Ctrl+8 | Centered `* * *` separator |

You can rename styles and change their font, size, spacing and dialogue
symbol in `settings.ini`.

**Writing and formatting**
- Bold, italic, underline, strikethrough, subscript and superscript, plus
  font color and highlight.
- Alignment, indents, line and paragraph spacing, letter spacing.
- Numbered and bulleted lists.
- Images, centered between paragraphs: drag to move them, use the handles to
  resize them.
- Page breaks with Ctrl+Enter (in Pages and Mixed views). Backspace at the
  start of the new page removes the break.
- Page size and margins (Letter, A4, A5, 6 × 9 in and other book sizes, or
  custom).

**Navigating a long book**
- **Navigator** (F5): a tree of titles, subtitles and sections. Click one to
  jump to it.
- **Collapsible headings:** hide a chapter's text with the triangle next to
  its title.
- **Find and replace** (Ctrl+F / Ctrl+H), with all matches highlighted.
- **Word count** in the status bar, also for the selected text.

**Reviewing**
- **Comments** (Ctrl+Alt+M) on any text, each with its own color, in a side
  panel (F6). They are saved as real Word comments.

**Interface languages:** English, Spanish, Portuguese, French, German,
Italian, Chinese and Japanese (Settings → Language).

## What it is not

SWordBook is for writing the draft, not for final typesetting. Tables,
headers and footers, footnotes, tracked changes and multiple columns are
outside its scope. For those, open your `.docx` in a full word processor.
When it opens a `.docx` from another program, it keeps the text and its
formatting, but drops what it doesn't support.

## Download

Portable builds are on the [Releases](../../releases) page. No installation
is needed:

| System | File | How to start it |
| --- | --- | --- |
| Windows 10/11 (64-bit) | `SWordBook-v…-windows-x64.zip` | Unzip and run `SWordBook.exe` |
| Linux (64-bit) | `SWordBook-v…-linux-x64.tar.gz` | Extract and run `./SWordBook` |
| macOS (Apple Silicon) | `SWordBook-v…-macos-arm64.zip` | Unzip, move `SWordBook.app` to Applications and open it |

**Windows:** SmartScreen may warn that the app is from an unknown publisher.
Click **More info → Run anyway**.

**macOS:** the app is not signed by Apple. The first time, right-click
`SWordBook.app` → **Open** → **Open**. If macOS says it is damaged, run
`xattr -cr /Applications/SWordBook.app` in Terminal.

**Linux:** if it doesn't start, install `libxcb-cursor0`
(`sudo apt install libxcb-cursor0`).

## Settings

On Windows and Linux, settings are stored in `settings.ini`, next to the
program. On macOS, they are stored in the app's data folder. The file is
created on first launch with every option, and you can edit it by hand while
the program is closed: default page size and margins, view, zoom, backup
timing, and each style's name, font, size, color and spacing.

## Run from source

Requires Python 3.9 or newer.

```
pip install -r requirements.txt
python swordbook.py
```

## Build a portable package

```
python build.py
```

The package for the current system is created in `release/`. When a release
is published, GitHub Actions builds the Windows, Linux and macOS packages and
attaches them to it (see `.github/workflows/build.yml`).

## License

MIT — © DinoSoft.1998
