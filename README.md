**A WordPad for book drafts.** A small, distraction-free word processor for drafting novels, stories and e-books.

SWordBook isn't trying to be Word. It offers the essential tools for book drafting: chapter titles, dialogue, scene breaks, a navigator, and comments. It also gives you **three ways to work on your manuscript**: paginated pages (like Word), one continuous sheet (like WordPad), or a combined view (SWordBook). You can also customize page lengths to organize your manuscript with one chapter per sheet.

It saves standard `.docx` files, so when your draft is finished, you can give it the final formatting in Word, Google Docs, or LibreOffice Writer.

> **Status: alpha (v0.1.0).** It works and saves standard `.docx` files.
> While you write, SWordBook keeps an automatic backup of your unsaved
> changes and offers to recover them if it ever closes unexpectedly. That
> backup is not a version history, though: it is deleted when you save. As
> with any early release, keep your own copies of manuscripts that matter.

## ★ Three Views, One Manuscript

The view mode is SWordBook's signature feature. The same book can be displayed in three different ways, and you can switch between them at any time from **Settings → View**.

| View | What you see | Best for |
| --- | --- | --- |
| **Pages** | Separate pages with your chosen paper size and margins, exactly as they will print | Checking length, layout and page breaks |
| **No Pages** | One continuous sheet with no page breaks, at the same width as the printed page | Writing naturally without page breaks interrupting a scene |
| **Mixed** | A continuous sheet that splits only where you insert a page break (Ctrl+Enter). Each sheet is as long as its content requires | Writing chapter by chapter, with each chapter on its own sheet regardless of length |

**Mixed view** gives you a different way to organize your manuscript. Insert a page break before each chapter, and your manuscript becomes a collection of chapters, each on its own sheet: long chapters get long sheets, short chapters get short sheets. Nothing gets cut off in the middle of a paragraph. You can still see where the book is divided without the constant interruptions of fixed-length pages.

All three views wrap lines according to the actual page width, so line lengths always match the printed result. Your chosen view is saved in the `.docx` file, so the document reopens the same way. Word and Writer open documents created in No Pages and Mixed views using their web layout.

Zoom ranges from 50% to 200% (Ctrl + mouse wheel or the slider in the status bar), with *Fit to Width* and *Whole Page* options.

## Why SWordBook?

- **Three views, one manuscript.** Switch between Pages, No Pages and Mixed to review your layout, write without interruptions, or organize your draft chapter by chapter. Your manuscript adapts to the way you write.
- **Built for book drafts.** Chapter titles, dialogue, thoughts, scene breaks and section headings are ready to use. Apply them with a click or a keyboard shortcut instead of formatting every element by hand.
- **Your chapters, your way.** In Mixed view, each chapter can occupy its own sheet, regardless of length. Keep your manuscript organized without forcing every chapter into fixed-length pages.
- **A simple editor with a purpose.** Get the straightforward writing experience of a lightweight text editor, with the book-specific tools that general-purpose editors aren't designed around.
- **Navigate your manuscript.** Jump between chapters and sections, collapse headings, search and replace text, track your word count, and add comments without losing your place.
- **Your files stay yours.** Work with standard `.docx` files and continue editing your manuscript in Microsoft Word, Google Docs or LibreOffice Writer whenever you need more advanced formatting.
- **Portable and self-contained.** No installation or account required. On Windows and Linux, carry the program and its settings in a single folder, even on a USB drive.
- **Extra protection for your writing.** SWordBook keeps a backup while you have unsaved changes and offers recovery if the program closes unexpectedly.

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
| Quote | Ctrl+7 | Letters, epigraphs and quoted texts |
| Scene Break | Ctrl+8 | Centered `* * *` separator |

You can rename styles and change their font, size, spacing and dialogue symbol in `settings.ini`.

**Writing and Formatting**

- Bold, italic, underline, strikethrough, subscript and superscript, plus font color and highlighting.
- Text alignment, indents, line spacing, paragraph spacing and letter spacing.
- Numbered and bulleted lists.
- Images centered between paragraphs: drag to move them and use the handles to resize them.
- Page breaks with Ctrl+Enter (in Pages and Mixed views). Press Backspace at the beginning of the new page to remove the break.
- Page size and margins (Letter, A4, A5, 6 × 9 inches and other book sizes, or custom dimensions).

**Navigating Long Manuscripts**

- **Navigator (F5):** A tree of titles, subtitles and sections. Click any item to jump directly to it.
- **Collapsible headings:** Hide a chapter's text using the triangle next to its title.
- **Find and Replace (Ctrl+F / Ctrl+H):** Search your manuscript with all matches highlighted.
- **Word count:** Displayed in the status bar, including the word count of selected text.

**Reviewing**

- **Comments (Ctrl+Alt+M):** Add comments to any text, each with its own color, in a side panel (F6). Comments are saved as actual Word comments.

**Interface languages:** English, Spanish, Portuguese, French, German, Italian, Chinese and Japanese (Settings → Language).

## What It Isn't

SWordBook is designed for writing drafts, not for final typesetting. Tables, headers and footers, footnotes, tracked changes and multiple columns are outside its scope. For these tasks, open your `.docx` file in a full-featured word processor.

When opening a `.docx` file created in another program, SWordBook preserves the text and its formatting but discards elements it doesn't support.

## Download

Portable builds are available on the [Releases](https://chatgpt.com/releases) page. No installation is required.

| System | File | How to launch |
| --- | --- | --- |
| Windows 10/11 (64-bit) | `SWordBook-v…-windows-x64.zip` | Extract the archive and run `SWordBook.exe` |
| Linux (64-bit) | `SWordBook-v…-linux-x64.tar.gz` | Extract the archive and run `./SWordBook` |
| macOS (Apple Silicon) | `SWordBook-v…-macos-arm64.zip` | Extract the archive, move `SWordBook.app` to Applications, and open it |

**Windows:** SmartScreen may warn you that the app is from an unknown publisher. Click **More info → Run anyway**.

**macOS:** The app is not signed by Apple. The first time you open it, right-click `SWordBook.app` → **Open** → **Open**. If macOS reports that the app is damaged, run `xattr -cr /Applications/SWordBook.app` in Terminal.

**Linux:** If the app doesn't start, install `libxcb-cursor0` using `sudo apt install libxcb-cursor0`.

## Settings

On Windows and Linux, settings are stored in `settings.ini`, next to the program. On macOS, they are stored in the app's data folder.

The file is created on first launch and includes every available option. You can edit it manually while the program is closed to change the default page size and margins, view mode, zoom, backup intervals, and each style's name, font, size, color and spacing.

## Run from Source

Requires Python 3.9 or newer.

```bash
pip install -r requirements.txt
python swordbook.py
```

## Build a Portable Package

```bash
python build.py
```

The package for the current system is created in the `release/` folder. When a new release is published, GitHub Actions builds the Windows, Linux and macOS packages and attaches them to the release (see `.github/workflows/build.yml`).

## License
MIT — © DinoSoft.1998
