# PPTX to Markdown Batch Converter

A Python script to convert PowerPoint presentations (.pptx) to Markdown format with optional image extraction.

## Features

- ✅ Batch convert multiple PPTX files at once
- ✅ Extract text content, tables, and data from slides
- ✅ Optionally extract and save embedded images
- ✅ Automatic metadata headers (filename, conversion date)
- ✅ Progress tracking and summary statistics
- ✅ Support for single file or directory conversion

## Requirements

- Python 3.6+
- `markitdown` library with PPTX support

## Installation

Install the required dependency:

```bash
pip install "markitdown[pptx]" --break-system-packages
```

## Usage

### Basic Usage - Convert All PPTX Files

Convert all PPTX files in the default upload directory:

```bash
python pptx_to_markdown.py
```

This will:
- Look for PPTX files in `~/pptx`
- Convert them to Markdown
- Save output to `~/pptx`

### Convert Files from a Specific Directory

```bash
python pptx_to_markdown.py -i /path/to/presentations
```

### Specify Output Directory

```bash
python pptx_to_markdown.py -o /path/to/output
```

### Convert a Single File

```bash
python pptx_to_markdown.py -f presentation.pptx
```

### Extract and Keep Images

By default, images are extracted for processing but not saved. To keep them:

```bash
python pptx_to_markdown.py --keep-images
```

This will create a separate folder for each presentation's images (e.g., `presentation_images/`).

### Combined Options

```bash
python pptx_to_markdown.py -i ~/Documents/Reports -o ~/Markdown -keep-images
```

## Command Line Options

| Option | Short | Description |
|--------|-------|-------------|
| `--input-dir` | `-i` | Input directory containing PPTX files (default: `/mnt/user-data/uploads`) |
| `--output-dir` | `-o` | Output directory for markdown files (default: `/mnt/user-data/outputs`) |
| `--file` | `-f` | Convert a single PPTX file instead of batch processing |
| `--keep-images` | - | Extract and keep images from presentations |
| `--help` | `-h` | Show help message |

## Output Format

Each converted file will include:

1. **Metadata Header**
   - Original filename
   - Conversion timestamp

2. **Extracted Content**
   - All slide text content
   - Tables (formatted as Markdown tables)
   - Bullet points and lists
   - Notes and comments

3. **Images** (if `--keep-images` is used)
   - Saved to `[filename]_images/` directory
   - Referenced in the markdown file

## Example Output

```
$ python pptx_to_markdown.py

🔍 Found 3 PPTX file(s) to convert
📁 Input directory: /mnt/user-data/uploads
📁 Output directory: /mnt/user-data/outputs
🖼️  Keep images: False

📄 Processing: Q4_Report.pptx
  ✅ Created: Q4_Report.md
  📊 Size: 45,231 bytes

📄 Processing: Budget_2026.pptx
  ✅ Created: Budget_2026.md
  📊 Size: 32,108 bytes

📄 Processing: Team_Update.pptx
  ✅ Created: Team_Update.md
  📊 Size: 18,456 bytes

============================================================
📊 CONVERSION SUMMARY
============================================================
✅ Successful: 3
❌ Failed: 0
📁 Output location: /mnt/user-data/outputs
============================================================
```

## What Gets Converted

### ✅ Successfully Converted
- Slide titles and content
- Bullet points and numbered lists
- Tables (converted to Markdown table format)
- Text boxes and annotations
- Speaker notes
- Basic text formatting

### ⚠️ Limitations
- Complex animations are not preserved
- Custom fonts may not be referenced
- Advanced slide layouts are simplified
- Some embedded objects may not convert perfectly
- Charts are converted to data tables where possible

## Use Cases

### Project Knowledge Base
Convert technical reports, CTO updates, and project status presentations to searchable markdown for Project Knowledge in Claude.

### Documentation
Transform presentation content into documentation that can be version controlled, searched, and referenced easily.

### Archive & Reference
Create text-based archives of presentations that are more accessible than binary PPTX files.

### Data Extraction
Extract tabular data from slides for further analysis or reporting.


## Troubleshooting

### "markitdown is not installed"
```bash
pip install "markitdown[pptx]" --break-system-packages
```

### "No .pptx files found"
Check that:
- You're pointing to the correct input directory
- Files have the `.pptx` extension (not `.ppt`)
- You have read permissions for the directory

### Conversion fails for specific file
Some heavily customised or corrupted PPTX files may fail. Try:
- Opening and re-saving the file in PowerPoint
- Using the `-f` option to convert just that file and see the error message

## Advanced Usage

### Process files in the current directory
```bash
python pptx_to_markdown.py -i .
```

### Convert and keep images with custom paths
```bash
python pptx_to_markdown.py \
  -i ~/Presentations/2025 \
  -o ~/Documents/Markdown \
  --keep-images
```

### Scripting / Automation
```bash
#!/bin/bash
# Convert all presentations in multiple directories

for dir in Q1 Q2 Q3 Q4; do
  echo "Converting $dir reports..."
  python pptx_to_markdown.py -i ~/Reports/$dir -o ~/Markdown/$dir
done
```

## Output Structure

```
outputs/
├── Presentation_1.md
├── Presentation_1_images/    # (if --keep-images used)
│   ├── image1.png
│   ├── image2.png
│   └── image3.png
├── Presentation_2.md
├── Presentation_2_images/
│   └── image1.png
└── ...
```

## Support

For issues or questions about the conversion script, check:
1. That all dependencies are installed correctly
2. Input files are valid PPTX format
3. You have write permissions to the output directory

