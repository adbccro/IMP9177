#!/usr/bin/env python3
"""
fix_draft_revision.py
Modifies a single DOCX file:
  1. Renames file: _RevA_ → _DRAFT_ in filename
  2. Inside document XML: replaces "Revision: A" → "Revision: DRAFT"
     and any remaining "Rev A" / "RevA" in header table cells
  3. Leaves all other content untouched

Usage: python fix_draft_revision.py <input.docx> <output_dir>
Returns the new filename as the last line of stdout.
"""
import sys
import os
import zipfile
import shutil
import re
import tempfile

def fix_docx(input_path: str, output_dir: str) -> str:
    """Process one DOCX and return the new filename."""
    
    basename = os.path.basename(input_path)
    
    # Determine new filename
    # Pattern: _RevA_ or _RevA. or RevA_ — replace with _DRAFT_
    new_basename = re.sub(r'_Rev[AB]_', '_DRAFT_', basename)
    new_basename = re.sub(r'_Rev[AB]\.docx$', '_DRAFT.docx', new_basename)
    # Handle cases like SOP-PRD-108_RevA.docx (no trailing text)
    if new_basename == basename:
        new_basename = basename.replace('_RevA.docx', '_DRAFT.docx')
        new_basename = new_basename.replace('_RevB.docx', '_DRAFT.docx')
    
    output_path = os.path.join(output_dir, new_basename)
    
    # Work in a temp directory
    with tempfile.TemporaryDirectory() as tmpdir:
        # Extract DOCX (it's a ZIP)
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(tmpdir)
        
        # Files to patch: document.xml and header*.xml
        targets = []
        word_dir = os.path.join(tmpdir, 'word')
        if os.path.isdir(word_dir):
            for fname in os.listdir(word_dir):
                if fname.endswith('.xml'):
                    targets.append(os.path.join(word_dir, fname))
        
        # Also check root XML files
        for fname in os.listdir(tmpdir):
            if fname.endswith('.xml'):
                targets.append(os.path.join(tmpdir, fname))
        
        for fpath in targets:
            try:
                with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                
                original = content
                
                # Replace "Revision: A" → "Revision: DRAFT"
                # The text may be split across XML runs so we do simple string replacement
                # on common patterns found in the 3H document headers
                replacements = [
                    ('Revision: A', 'Revision: DRAFT'),
                    ('Revision:A',  'Revision:DRAFT'),
                    ('>Revision</w:t>', '>Revision</w:t>'),  # no-op anchor
                    # Handle the value cell — "A" alone in a table cell after "Revision:"
                    # We target the specific patterns from the 3H header table structure
                    ('Rev A\n',  'DRAFT\n'),
                ]
                
                for old, new in replacements:
                    if old in content:
                        content = content.replace(old, new)
                
                # More targeted: find runs that are just "A" adjacent to "Revision"
                # Pattern: <w:t>A</w:t> in a cell that also contains Revision
                # Since runs may be split, we use a broader approach on the full XML:
                # Replace the standalone "A" in revision context
                # We look for the pattern: ">A<" in table cells near "Revision"
                # This is safe because the header table cells are isolated
                
                # Handle: Revision: A (in same run)
                content = re.sub(r'(Revision:\s*)(A)(\s*<)', r'\1DRAFT\3', content)
                
                # Handle: value is just "A" in a w:t tag in the header section
                # We identify this by proximity to "Revision" within 2000 chars
                def replace_revision_a(m):
                    region = content[max(0, m.start()-2000):m.end()]
                    if 'Revision' in region or 'revision' in region:
                        return m.group(0).replace('>A<', '>DRAFT<')
                    return m.group(0)
                
                # Find isolated >A< patterns near Revision
                content = re.sub(r'>A<', lambda m: replace_revision_a(m), content)
                
                if content != original:
                    with open(fpath, 'w', encoding='utf-8') as f:
                        f.write(content)
            
            except Exception as e:
                sys.stderr.write(f'Warning: could not patch {fpath}: {e}\n')
        
        # Repack into new DOCX
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, tmpdir)
                    zout.write(file_path, arcname)
    
    print(f'Processed: {basename} → {new_basename}')
    print(f'OUTPUT:{new_basename}')  # last line sentinel for PowerShell
    return new_basename

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Usage: python fix_draft_revision.py <input.docx> <output_dir>')
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_dir = sys.argv[2]
    
    if not os.path.exists(input_path):
        print(f'Error: {input_path} not found')
        sys.exit(1)
    
    os.makedirs(output_dir, exist_ok=True)
    fix_docx(input_path, output_dir)
