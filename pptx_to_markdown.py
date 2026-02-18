#!/usr/bin/env python3
"""
PPTX to Markdown Batch Converter
Converts PowerPoint presentations to markdown format with image extraction.
"""

import os
import sys
import subprocess
import zipfile
import shutil
from pathlib import Path
import argparse


def extract_images_from_pptx(pptx_path, output_dir):
    """Extract images from PPTX file."""
    images = []
    temp_extract = output_dir / "temp_extract"

    try:
        # Unzip the PPTX
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_extract)

        # Find all images
        media_dir = temp_extract / "ppt" / "media"
        if media_dir.exists():
            for img_file in sorted(media_dir.iterdir()):
                if img_file.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif']:
                    images.append(img_file)

        return images, temp_extract
    except Exception as e:
        print(f"  Warning: Could not extract images: {e}")
        return [], None


def analyze_image(image_path):
    """
    Extract text from an image using OCR (EasyOCR with pytesseract fallback).
    Returns a markdown-formatted string of extracted content, or empty string.
    """
    # Try to preprocess with Pillow
    try:
        from PIL import Image as PILImage
        img = PILImage.open(image_path).convert("RGB")
        # Scale up small images for better OCR accuracy
        w, h = img.size
        if w < 1000:
            scale = 1000 / w
            img = img.resize((int(w * scale), int(h * scale)), PILImage.LANCZOS)
        preprocessed_path = str(image_path) + "_ocr_tmp.png"
        img.save(preprocessed_path)
    except Exception:
        preprocessed_path = str(image_path)

    lines = []

    # Try EasyOCR first
    try:
        import easyocr
        reader = easyocr.Reader(['en'], gpu=False, verbose=False)
        results = reader.readtext(preprocessed_path, detail=0)
        lines = [line.strip() for line in results if line.strip()]
    except ImportError:
        pass
    except Exception as e:
        print(f"    EasyOCR error on {image_path.name}: {e}")

    # Fallback to pytesseract if EasyOCR found nothing or isn't available
    if not lines:
        try:
            import pytesseract
            from PIL import Image as PILImage
            img = PILImage.open(preprocessed_path)
            raw = pytesseract.image_to_string(img)
            lines = [l.strip() for l in raw.splitlines() if l.strip()]
        except ImportError:
            pass
        except Exception as e:
            print(f"    pytesseract error on {image_path.name}: {e}")

    # Clean up temp file
    try:
        if preprocessed_path != str(image_path) and os.path.exists(preprocessed_path):
            os.remove(preprocessed_path)
    except Exception:
        pass

    if not lines:
        return ""

    # Format as a blockquote section
    formatted = "\n".join(f"> {line}" for line in lines)
    return formatted


def _check_ocr_available():
    """Check which OCR backends are available, warn if none."""
    has_easyocr = False
    has_pytesseract = False

    try:
        import easyocr  # noqa: F401
        has_easyocr = True
    except ImportError:
        pass

    try:
        import pytesseract  # noqa: F401
        pytesseract.get_tesseract_version()
        has_pytesseract = True
    except Exception:
        pass

    if not has_easyocr and not has_pytesseract:
        print("  Warning: No OCR backend available. Image content will not be extracted.")
        print("  Install with: pip install easyocr  (or: pip install pytesseract + brew install tesseract)")
        return False
    return True


def convert_pptx_to_markdown(pptx_path, output_dir, keep_images=False, image_ocr=True):
    """Convert a single PPTX file to markdown."""
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)

    # Create output filename
    md_filename = pptx_path.stem + ".md"
    md_path = output_dir / md_filename

    print(f"\nProcessing: {pptx_path.name}")

    # Extract images first
    images, temp_extract = extract_images_from_pptx(pptx_path, output_dir)

    if images and keep_images:
        # Create images subdirectory
        img_dir = output_dir / f"{pptx_path.stem}_images"
        img_dir.mkdir(exist_ok=True)

        print(f"  Found {len(images)} images, copying to {img_dir.name}/")

        # Copy images to output
        for img in images:
            dest = img_dir / img.name
            shutil.copy2(img, dest)

    # Convert using markitdown
    try:
        result = subprocess.run(
            ['python3', '-m', 'markitdown', str(pptx_path)],
            capture_output=True,
            text=True,
            check=True
        )

        markdown_content = result.stdout

        # Add metadata header
        header = f"""# {pptx_path.stem}

**Source File:** {pptx_path.name}
**Converted:** {subprocess.run(['date', '+%Y-%m-%d %H:%M:%S'], capture_output=True, text=True).stdout.strip()}

---

"""

        # Analyze images with OCR and append sections
        image_sections = ""
        if image_ocr and images:
            print(f"  Running OCR on {len(images)} image(s)...")
            for img in images:
                print(f"    Analyzing {img.name}...")
                ocr_text = analyze_image(img)
                if ocr_text:
                    image_sections += f"\n## Image: {img.name}\n\n> **[Extracted image content]**\n{ocr_text}\n"
                else:
                    image_sections += f"\n## Image: {img.name}\n\n> *No text content detected*\n"

        # Write the markdown file
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(header + markdown_content)
            if image_sections:
                f.write("\n---\n\n## Embedded Images\n" + image_sections)

        print(f"  Created: {md_filename}")
        print(f"  Size: {md_path.stat().st_size:,} bytes")

        return True

    except subprocess.CalledProcessError as e:
        print(f"  Error converting: {e}")
        print(f"  stderr: {e.stderr}")
        return False

    finally:
        # Cleanup temp directory
        if temp_extract and temp_extract.exists():
            shutil.rmtree(temp_extract, ignore_errors=True)


def batch_convert(input_dir, output_dir, keep_images=False, image_ocr=True):
    """Convert all PPTX files in a directory."""
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)

    # Find all PPTX files
    pptx_files = list(input_dir.glob("*.pptx"))
    print(f"\nFiles to convert: ")
    print(*pptx_files, sep=' - ')

    if not pptx_files:
        print(f"No .pptx files found in {input_dir}")
        return

    print(f"\nFound {len(pptx_files)} PPTX file(s) to convert")
    print(f"Input directory: {input_dir}")
    print(f"Output directory: {output_dir}")
    print(f"Keep images: {keep_images}")
    print(f"Image OCR: {image_ocr}")

    if image_ocr:
        _check_ocr_available()

    # Convert each file
    successful = 0
    failed = 0

    for pptx_file in pptx_files:
        if convert_pptx_to_markdown(pptx_file, output_dir, keep_images, image_ocr):
            successful += 1
        else:
            failed += 1

    # Summary
    print("\n" + "="*60)
    print("CONVERSION SUMMARY")
    print("="*60)
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    print(f"Output location: {output_dir.absolute()}")
    print("="*60 + "\n")


def main():
    parser = argparse.ArgumentParser(
        description="Batch convert PPTX files to Markdown format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert all PPTX files in current directory
  %(prog)s

  # Convert files from specific directory
  %(prog)s -i /path/to/presentations

  # Specify output directory
  %(prog)s -o /path/to/output

  # Keep extracted images
  %(prog)s --keep-images

  # Convert single file
  %(prog)s -f presentation.pptx

  # Skip OCR on embedded images
  %(prog)s --no-image-ocr
        """
    )

    parser.add_argument(
        '-i', '--input-dir',
        default=str(Path.home() / "pptx"),
        help='Input directory containing PPTX files (default: ~/pptx)'
    )

    parser.add_argument(
        '-o', '--output-dir',
        default=str(Path.home() / "pptx"),
        help='Output directory for markdown files (default: ~/pptx)'
    )

    parser.add_argument(
        '-f', '--file',
        help='Convert a single PPTX file instead of batch processing'
    )

    parser.add_argument(
        '--keep-images',
        action='store_true',
        help='Extract and keep images from presentations'
    )

    parser.add_argument(
        '--no-image-ocr',
        action='store_true',
        help='Skip OCR analysis of embedded images (faster, text-only output)'
    )

    args = parser.parse_args()

    # Check if markitdown is available
    try:
        subprocess.run(
            ['python3', '-m', 'markitdown', '--help'],
            capture_output=True,
            check=True
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("Error: markitdown is not installed")
        print("\nPlease install it with:")
        print("  pip install markitdown[pptx] --break-system-packages")
        sys.exit(1)

    image_ocr = not args.no_image_ocr

    # Warn early if OCR requested but no backend found
    if image_ocr:
        _check_ocr_available()

    # Convert single file or batch
    if args.file:
        output_dir = Path(args.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        if convert_pptx_to_markdown(args.file, output_dir, args.keep_images, image_ocr):
            print(f"\nSuccessfully converted {args.file}")
        else:
            print(f"\nFailed to convert {args.file}")
            sys.exit(1)
    else:
        batch_convert(args.input_dir, args.output_dir, args.keep_images, image_ocr)


if __name__ == "__main__":
    main()
