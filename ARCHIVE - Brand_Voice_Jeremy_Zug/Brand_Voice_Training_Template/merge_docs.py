#!/usr/bin/env python3

import os
from datetime import datetime

def read_file_content(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read().strip()

def generate_toc(sections):
    toc = ["## Table of Contents\n"]
    for i, section in enumerate(sections, 1):
        # Convert the filename to a section title
        title = section[3:].split('.')[0].replace('_', ' ').title()
        # Create the markdown link
        link = title.lower().replace(' ', '-')
        toc.append(f"{i}. [{title}](#{link})")
    return '\n'.join(toc)

def main():
    # Directory containing the markdown files
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Output file path
    output_file = os.path.join(base_dir, "Generated_Brand_Voice_Master.md")
    
    # List of files to merge in order
    files_to_merge = [
        "00_MAIN.md",
        "01_executive_summary.md",
        "02_brand_foundation.md",
        "03_brand_assets_and_history.md",
        "04_content_generation_guidelines.md",
        "05_linguistic_patterns.md",
        "06_audience_engagement.md",
        "07_advanced_storytelling.md",
        "08_brand_language_parameters.md",
        "09_quality_control_system.md",
        "10_implementation_guidelines.md",
        "11_SYSTEM_SUMMARY.md"
    ]
    
    # Generate header
    header = f"""# Brand Voice Documentation
Master Document | Version 1.0
Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"""
    
    # Generate table of contents
    toc = generate_toc(files_to_merge)
    
    # Start building the document
    merged_content = [header, toc, "\n---"]
    
    # Process each file
    for filename in files_to_merge:
        file_path = os.path.join(base_dir, filename)
        if os.path.exists(file_path):
            content = read_file_content(file_path)
            # Remove any existing H1 headers (# Title) as we'll structure these differently
            lines = content.split('\n')
            filtered_lines = []
            for line in lines:
                # Skip H1 headers and empty lines at the start
                if line.strip().startswith('# ') or (not filtered_lines and not line.strip()):
                    continue
                filtered_lines.append(line)
            
            # Add section header
            section_title = filename[3:].split('.')[0].replace('_', ' ').title()
            merged_content.append(f"\n# {section_title}\n")
            merged_content.append('\n'.join(filtered_lines))
            merged_content.append('\n---')
    
    # Write the merged content
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(merged_content))
    
    print(f"Successfully created {output_file}")

if __name__ == "__main__":
    main()
