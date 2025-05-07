# pdf2pptx

Convert PDF documents to PowerPoint presentations using Python.

## Features

- Fully preserves the original PDF’s layout
- Preserves nearly all original PDF hyperlinks (internal page links are not preserved)
- Supports custom bitmap DPI settings or direct use of vector graphics to ensure high-quality page rendering

## Installation

```bash
pip install -r requirements.txt
```

\[Optional\]
Install Inkscape to enable vector graphics support.

## Usage

```bash
python pdf2pptx.py [-h] [-svg] [--dpi DPI] [--aspect-ratio ASPECT_RATIO] [--inkscape-path INKSCAPE_PATH] input [output]
```

### positional arguments:
- `input`: The input pdf file
- `output`: The output PowerPoint file

### options:
- `-h, --help`: show this help message and exit
- `-svg`: Use SVG for rendering instead of PNG. This is experimental and depends on inkscape.
- -`-dpi DPI`: The DPI to render the pdf, default is 600. When using SVG, this is ignored.
- `--aspect-ratio ASPECT_RATIO`: The aspect ratio of the slides like "16:9" for example, default is automatically derived from the pdf
- `--inkscape-path INKSCAPE_PATH`: Path to the inkscape executable. This is only used when --svg is specified.

## License
This project is licensed under the GPL-3.0 License - see the [LICENSE](LICENSE) file for details.