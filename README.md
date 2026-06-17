# PDF Extraction Pipeline with AI Builder

A complete solution for converting documents and images to PDF, then using Power Automate with AI Builder to automatically extract text and structured data.

## Overview

This solution consists of three main components:

1. **Word-to-PDF Converter** - PowerShell script that converts Word documents (.docx, .doc) to PDF
2. **Image-to-PDF Converter** - PowerShell script that converts images (photos, scans) to PDF
3. **Power Automate Flow** - Intelligent document processing that extracts text using AI Builder

### Use Cases

- **Expense Processing**: Convert receipt photos to PDF, extract amounts and dates
- **Document Management**: Batch convert documents and scans to standardized PDF format
- **Data Extraction**: Automatically extract key information from invoices, forms, and contracts
- **Compliance**: Maintain a centralized repository of processed documents

---

## Demo Video

Watch the complete solution in action:

**[▶️ Watch Demo Video](https://github.com/sposaro/ai-powered-document-extractor/raw/main/Text%20Extraction%20Demo.mp4)** (50 MB)

The demo shows:
- Converting Word documents to PDF
- Converting cell phone photos to PDF
- Importing the Power Automate flow
- Configuring the PDF folder path
- Running AI Builder text extraction
- Reviewing extracted results

---

## Prerequisites

### For Local PDF Conversion

- **Microsoft Word** (for .docx/.doc conversion)
- **ImageMagick** (for image conversion)
  - Install: `winget install ImageMagick.ImageMagick`
  - Or download from: https://imagemagick.org/script/download.php#windows

### For Power Automate and AI

- **Microsoft 365 Account** or **Power Automate Cloud** subscription
- **AI Builder License** (included with most M365 plans)
- Access to Power Automate at https://make.powerautomate.com

---

## Folder Structure

```
TextExtractor/
├── Input Word Files/           # Source Word documents
│   ├── *.docx
│   └── *.doc
├── Input Images/               # Source images (cell phone photos, scans)
│   └── *.jpg, *.png, *.bmp, etc.
├── Output PDF Files/           # Generated PDFs (input for Power Automate)
│   └── *.pdf
├── Convert-WordFilesToPdf.ps1  # Word conversion script
├── Convert-ImagesToPdf.ps1     # Image conversion script
├── PowerAutomateFlow-ExtractPDFs.zip  # Exported Power Automate flow
├── README.md                   # This file
└── NARRATION_SCRIPT.md        # Screen recording guide
```

---

## Component 1: Word-to-PDF Converter

### Features

- Converts `.docx`, `.docx`, `.docm`, `.dot`, `.dotx`, `.dotm` files
- Validates legacy `.doc` files to prevent conversion hangs
- Supports recursive folder processing
- Can skip existing PDFs or overwrite them
- Detailed error reporting

### Usage

```powershell
# Basic usage - convert current folder
.\Convert-WordFilesToPdf.ps1 -InputDirectory ".\Input Word Files\"

# With output folder
.\Convert-WordFilesToPdf.ps1 `
  -InputDirectory ".\Input Word Files\" `
  -OutputDirectory ".\Output PDF Files\" `
  -Recurse `
  -Overwrite

# Parameters
-InputDirectory   (Required) Folder with Word files
-OutputDirectory  (Optional) Where to save PDFs
-Recurse          (Optional) Search subfolders
-Overwrite        (Optional) Replace existing PDFs
```

### Output

```
Converting: C:\Users\...\Alpine Expansion Kickoff Agenda.docx
Converting: C:\Users\...\Blue Yonder Onboarding Guide Sample.docx
...
Done.
Converted: 7
Skipped:   0
Failed:    0
```

---

## Component 2: Image-to-PDF Converter

### Features

- Converts `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, `.tiff` files
- Maintains image quality in PDF output
- Supports recursive folder processing
- Can skip existing PDFs or overwrite them
- Perfect for cell phone photos and scanned documents

### Dependencies

**Important**: This script requires ImageMagick to be installed.

Installation:
```powershell
winget install ImageMagick.ImageMagick
```

Or download from: https://imagemagick.org/script/download.php#windows

### Usage

```powershell
# Basic usage - convert current folder
.\Convert-ImagesToPdf.ps1 -InputDirectory ".\Input Images\"

# With output folder
.\Convert-ImagesToPdf.ps1 `
  -InputDirectory ".\Input Images\" `
  -OutputDirectory ".\Output PDF Files\" `
  -Recurse `
  -Overwrite

# Parameters
-InputDirectory   (Required) Folder with image files
-OutputDirectory  (Optional) Where to save PDFs
-Recurse          (Optional) Search subfolders
-Overwrite        (Optional) Replace existing PDFs
```

### Output

```
Converting: C:\Users\...\Contoso Invoice 12847.jpg
Converting: C:\Users\...\Fabrikam Receipt 2026-06-17.jpg
...
Done.
Converted: 5
Skipped:   0
Failed:    0
```

---

## Component 3: Power Automate Flow

### Overview

The Power Automate flow automates document processing using AI Builder:

1. **List PDF Files** - Gets all PDFs from the Output PDF Files folder
2. **Extract Text** - Uses AI Builder to analyze each PDF and extract text
3. **Store Results** - Saves extracted data in a variable for further processing

### Setup Instructions

#### Step 1: Import the Flow

1. Go to https://make.powerautomate.com
2. Click **My flows** → **Import** → **Import package**
3. Select `PowerAutomateFlow-ExtractPDFs.zip`
4. Configure required connections (OneDrive, SharePoint, etc.)
5. Click **Import**

#### Step 2: Configure the PDF Folder Path

1. Open the imported flow to edit it
2. Find the **"Get files in folder"** action
3. Set the **Location** to your Output PDF Files path:
   ```
   C:\Users\franksp\Documents\Clawpilot\TextExtractor\Output PDF Files\
   ```

#### Step 3: Review AI Builder Configuration

The flow includes an **"Extract information from PDFs"** action that:
- Processes each PDF document
- Uses AI to identify and extract text
- Stores results in the `ExtractedText` variable
- Handles errors gracefully

#### Step 4: Test the Flow

1. Click **Save** to save changes
2. Click **Test** → **Manually**
3. Click **Test** button to run the flow
4. Watch each step execute in real-time
5. Check the **ExtractedText** variable output

### Output Example

```json
{
  "ExtractedText": "CONTOSO INC.\nInvoice #12847\nDate: 2026-06-15\n..."
}
```

---

## Next Steps: Extending the Solution

The demo shows text extraction. Here are common extensions:

### Option 1: Save to Excel
Add a Power Automate action to write extracted data to an Excel file on OneDrive or SharePoint.

### Option 2: Store in Database
Connect to SQL Server or Microsoft Dataverse to create structured records from extracted data.

### Option 3: Send Notifications
Email or Teams message the extracted data to team members for review.

### Option 4: Trigger Workflows
Pass extracted data to other automated processes or business applications.

---

## Troubleshooting

### Word Conversion Issues

**Problem**: "Cannot convert the '...' value of type 'psobject' to type 'Object'"
- **Solution**: Ensure the script uses `[string]` and `[int]` casts for COM parameters
- The provided script has this fix applied

**Problem**: Script appears to hang on .doc files
- **Cause**: File is not a valid Word binary document
- **Solution**: The script automatically detects and skips invalid .doc files
- Check warnings in output for details

### Image Conversion Issues

**Problem**: "ImageMagick is not installed or not in PATH"
- **Solution**: Install ImageMagick using:
  ```powershell
  winget install ImageMagick.ImageMagick
  ```

**Problem**: Permission denied when writing output
- **Solution**: Ensure Output PDF Files folder exists and is writable
- Create manually if needed: `New-Item -ItemType Directory -Path ".\Output PDF Files\"`

### Power Automate Issues

**Problem**: Flow cannot access the PDF folder
- **Solution**: Verify the path in "Get files in folder" action is correct
- Ensure the path uses backslashes: `C:\Users\...\Output PDF Files\`

**Problem**: AI Builder extraction returns empty results
- **Solution**: Check that PDFs are readable and contain extractable text
- Some scanned documents may need OCR preprocessing

---

## Performance Notes

- Word conversion: ~2-5 seconds per document
- Image conversion: ~1-3 seconds per image
- AI extraction: ~5-10 seconds per document (depends on complexity)
- Batch of 12 PDFs: ~2-3 minutes end-to-end

---

## File Support

### Word Conversion
- ✅ .docx (Word 2007+)
- ✅ .docm (Word with macros)
- ✅ .dot (Word template)
- ✅ .dotx (Word 2007+ template)
- ✅ .dotm (Word template with macros)
- ❌ .doc (legacy - must be valid OLE format)

### Image Conversion
- ✅ .jpg / .jpeg
- ✅ .png
- ✅ .bmp
- ✅ .gif
- ✅ .tiff

### AI Extraction
- ✅ Any PDF created by this solution
- ✅ Standard business documents (invoices, receipts, forms)
- ✅ Scanned documents with clear text
- ⚠️ Handwritten text (limited OCR support)

---

## License & Support

For issues or feature requests, contact the solution owner.

---

## Quick Start

```powershell
# 1. Convert Word documents
.\Convert-WordFilesToPdf.ps1 `
  -InputDirectory ".\Input Word Files\" `
  -OutputDirectory ".\Output PDF Files\" `
  -Recurse -Overwrite

# 2. Convert images
.\Convert-ImagesToPdf.ps1 `
  -InputDirectory ".\Input Images\" `
  -OutputDirectory ".\Output PDF Files\" `
  -Overwrite

# 3. Verify output
Get-ChildItem ".\Output PDF Files\" -Filter "*.pdf" | Measure-Object

# 4. Import and test Power Automate flow at https://make.powerautomate.com
```

---

**Last Updated**: June 17, 2026
