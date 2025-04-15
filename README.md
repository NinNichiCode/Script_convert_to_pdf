# Office to PDF Converter Scripts

This repository provides a set of simple command-line scripts for converting Office documents (Word, Excel, PowerPoint, and images) into PDF using PowerShell on Windows.

## Included Scripts

| Script       | Description                          |
|--------------|--------------------------------------|
| `w_pdf`      | Convert Word documents (.doc/.docx) to PDF |
| `xls_pdf`    | Convert Excel files (.xls/.xlsx) to PDF |
| `ppt_pdf`    | Convert PowerPoint files (.ppt/.pptx) to PDF |
| `png_pdf`    | Convert image files (.png/.jpg/.bmp) to PDF |
| `excel2pdf.ps1` | PowerShell script used internally by `xls_pdf` |

##  Requirements

- Windows OS
- Git Bash with PowerShell access
- Microsoft Office (Word, Excel, PowerPoint) installed
- For image-to-PDF conversion (png_pdf, jpg_pdf): ImageMagick (required)


##  How to Use

### Step 1: Clone this repository

git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name

### Step 2: Make scripts executable (if needed)

chmod +x *.sh

### Step 3: Convert files
Convert Excel to PDF:
./xls_pdf "file1.xlsx"
./xls_pdf "file1.xlsx" "file2.xls"  # multiple files

Convert Word to PDF:
./w_pdf "document.docx"

Convert PowerPoint to PDF:
./ppt_pdf "slides.pptx"

Convert PNG or JPG images to PDF:
./png_pdf "image1.png" "image2.jpg"
## Output PDF files will be saved in the same folder as the input files.


To make the commands globally accessible from any location in Git Bash, you can add the folder containing these scripts to your system PATH.

Once added, you can simply run commands like xls_pdf, w_pdf, or ppt_pdf from any directory — as long as the target file (e.g. .xlsx, .docx, etc.) exists in the current directory.

This makes converting files to PDF much easier and faster.
