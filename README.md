# SWordBook

A word processor for writing books and e-books, saved as `.docx` (opens in
Word, Google Docs and LibreOffice Writer; accepted by Amazon KDP).

- Book styles: Body, Title, Subtitle, Section, Thoughts, Dialogue, Quote, Scene break
- Pages, No pages and Mixed views, zoom, collapsible headings, navigator
- Numbered and bulleted lists, images, comments, find and replace
- Automatic backup and recovery
- Interface in English, Spanish, Portuguese, French, German, Italian, Chinese and Japanese

## Download

Portable builds (no installation) are on the
[Releases](../../releases) page:

| System | File |
| --- | --- |
| Windows 10/11 (64-bit) | `SWordBook-v…-windows-x64.zip` — unzip and run `SWordBook.exe` |
| Linux (64-bit) | `SWordBook-v…-linux-x64.tar.gz` — extract and run `./SWordBook` |

On Linux, if it doesn't start, install `libxcb-cursor0`
(`sudo apt install libxcb-cursor0`).

Settings are kept in `settings.ini` next to the program, so the folder can
travel on a USB drive.

## Run from source

```
pip install -r requirements.txt
python swordbook.py
```

## Build a portable package

```
python build.py
```

The package for the current system is created in `release/`.

## License

MIT — © DinoSoft.1998
