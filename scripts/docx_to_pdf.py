"""
Convert one or more .docx files to PDF using Microsoft Word.

Usage:
    python docx_to_pdf.py <file1.docx> [file2.docx ...]
    python docx_to_pdf.py <directory>  (converts all .docx in the directory)

Output PDFs are saved alongside the .docx files with the same name.
"""

import sys
import os
from docx2pdf import convert


def main():
    if len(sys.argv) < 2:
        print("Usage: python docx_to_pdf.py <file.docx> [file2.docx ...]")
        sys.exit(1)

    for path in sys.argv[1:]:
        path = os.path.abspath(path)

        if os.path.isdir(path):
            # Convert all .docx in directory
            for f in os.listdir(path):
                if f.endswith('.docx') and not f.startswith('~$'):
                    docx_path = os.path.join(path, f)
                    pdf_path = docx_path.rsplit('.', 1)[0] + '.pdf'
                    print(f'Converting: {f}')
                    convert(docx_path, pdf_path)
                    print(f'  -> {os.path.basename(pdf_path)}')
        elif os.path.isfile(path) and path.endswith('.docx'):
            pdf_path = path.rsplit('.', 1)[0] + '.pdf'
            print(f'Converting: {os.path.basename(path)}')
            convert(path, pdf_path)
            print(f'  -> {os.path.basename(pdf_path)}')
        else:
            print(f'Skipping: {path} (not a .docx file or directory)')


if __name__ == '__main__':
    main()
