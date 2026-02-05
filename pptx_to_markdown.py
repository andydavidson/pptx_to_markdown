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
        print(f"  ⚠️  Warning: Could not extract images: {e}")
        return [], None


def convert_pptx_to_markdown(pptx_path, output_dir, keep_images=False):
    """Convert a single PPTX file to markdown."""
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    
    # Create output filename
    md_filename = pptx_path.stem + ".md"
    md_path = output_dir / md_filename
    
    print(f"\n📄 Processing: {pptx_path.name}")
    
    # Extract images first
    images, temp_extract = extract_images_from_pptx(pptx_path, output_dir)
    
    if images and keep_images:
        # Create images subdirectory
        img_dir = output_dir / f"{pptx_path.stem}_images"
        img_dir.mkdir(exist_ok=True)
        
        print(f"  📸 Found {len(images)} images, copying to {img_dir.name}/")
        
        # Copy images to output
        for img in images:
            dest = img_dir / img.name
            shutil.copy2(img, dest)
    
    # Convert using markitdown
    try:
        result = subprocess.run(
            ['python', '-m', 'markitdown', str(pptx_path)],
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
        
        # Write the markdown file
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(header + markdown_content)
        
        print(f"  ✅ Created: {md_filename}")
        print(f"  📊 Size: {md_path.stat().st_size:,} bytes")
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"  ❌ Error converting: {e}")
        print(f"  stderr: {e.stderr}")
        return False
    
    finally:
        # Cleanup temp directory
        if temp_extract and temp_extract.exists():
            shutil.rmtree(temp_extract, ignore_errors=True)


def batch_convert(input_dir, output_dir, keep_images=False):
    """Convert all PPTX files in a directory."""
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    
    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Find all PPTX files
    pptx_files = list(input_dir.glob("*.pptx"))
    
    if not pptx_files:
        print(f"❌ No .pptx files found in {input_dir}")
        return
    
    print(f"\n🔍 Found {len(pptx_files)} PPTX file(s) to convert")
    print(f"📁 Input directory: {input_dir}")
    print(f"📁 Output directory: {output_dir}")
    print(f"🖼️  Keep images: {keep_images}")
    
    # Convert each file
    successful = 0
    failed = 0
    
    for pptx_file in pptx_files:
        if convert_pptx_to_markdown(pptx_file, output_dir, keep_images):
            successful += 1
        else:
            failed += 1
    
    # Summary
    print("\n" + "="*60)
    print("📊 CONVERSION SUMMARY")
    print("="*60)
    print(f"✅ Successful: {successful}")
    print(f"❌ Failed: {failed}")
    print(f"📁 Output location: {output_dir.absolute()}")
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
        """
    )
    
    parser.add_argument(
        '-i', '--input-dir',
        default='/mnt/user-data/uploads',
        help='Input directory containing PPTX files (default: /mnt/user-data/uploads)'
    )
    
    parser.add_argument(
        '-o', '--output-dir',
        default='/mnt/user-data/outputs',
        help='Output directory for markdown files (default: /mnt/user-data/outputs)'
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
    
    args = parser.parse_args()
    
    # Check if markitdown is available
    try:
        subprocess.run(
            ['python', '-m', 'markitdown', '--help'],
            capture_output=True,
            check=True
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("❌ Error: markitdown is not installed")
        print("\nPlease install it with:")
        print("  pip install markitdown[pptx] --break-system-packages")
        sys.exit(1)
    
    # Convert single file or batch
    if args.file:
        output_dir = Path(args.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        if convert_pptx_to_markdown(args.file, output_dir, args.keep_images):
            print(f"\n✅ Successfully converted {args.file}")
        else:
            print(f"\n❌ Failed to convert {args.file}")
            sys.exit(1)
    else:
        batch_convert(args.input_dir, args.output_dir, args.keep_images)


if __name__ == "__main__":
    main()
