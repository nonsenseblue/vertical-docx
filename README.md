# vertical-docx

Markdown to vertical-writing docx converter for Japanese manuscripts.

Converts `.md` files into publication-ready `.docx` with vertical text layout (right-to-left columns), full-width character conversion, and configurable grid settings.

## Install

```bash
pip install python-docx lxml
```

## Usage

```bash
# Basic — A4 portrait, 40 chars x 30 lines, 10.5pt
python vertical_docx.py novel.md --author "Author Name"

# With synopsis prepended
python vertical_docx.py novel.md --author "Author Name" --synopsis synopsis.md

# 20 chars x 20 lines, 12pt, wider margins
python vertical_docx.py novel.md --chars 20 --lines 20 --font-size 12 \
  --margin-top 35 --margin-left 30 --margin-right 30

# Landscape, 30 chars x 40 lines
python vertical_docx.py novel.md --landscape --chars 30 --lines 40
```

## Options

| Option | Default | Description |
|--------|---------|-------------|
| `input` | (required) | Input markdown file |
| `-o, --output` | `<input>.docx` | Output docx path |
| `--title` | First `#` heading | Document title |
| `--author` | — | Author name on title page |
| `--synopsis` | — | Synopsis markdown (prepended) |
| `--font` | `Yu Mincho` | Font name |
| `--font-size` | `10.5` | Font size in pt |
| `--chars` | `40` | Characters per line |
| `--lines` | `30` | Lines per page |
| `--landscape` | off | Landscape orientation |
| `--margin-top` | `30` | Top margin (mm) |
| `--margin-bottom` | `25` | Bottom margin (mm) |
| `--margin-left` | `25` | Left margin (mm) |
| `--margin-right` | `25` | Right margin (mm) |
| `--line-pitch` | auto | docGrid linePitch override |
| `--char-space` | `0` | docGrid charSpace override |
| `--no-page-numbers` | off | Disable page numbers |

## Markdown support

- `# Heading` — title page (first occurrence) or chapter break
- `## Heading` — chapter page
- `### Heading` — section header
- `> text` — blockquote
- `---` / `-----` — scene break (renders as centered `*`)
- Code blocks (`` ``` ``) — preserved as-is
- `**bold**` and `` `code` `` markers are stripped

## How it works

The script uses `python-docx` to generate a Word document with:

- **Vertical text direction** (`tbRl`) — text flows top-to-bottom, columns right-to-left
- **Full-width conversion** — ASCII alphanumerics and spaces are converted to their full-width equivalents for proper vertical rendering
- **Character grid** — `linesAndChars` grid type with configurable pitch for precise layout control

## License

MIT
