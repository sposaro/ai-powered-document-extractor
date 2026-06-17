# PDF Extraction Pipeline - Screen Recording Narration Script

## Opening (10 seconds)
"Hi, I'm going to walk you through a complete PDF extraction pipeline that combines PowerShell conversion scripts with Power Automate and AI Builder for intelligent document processing.

This solution handles two common scenarios: converting Word documents and cell phone photos of receipts into PDFs, then extracting text automatically using AI."

---

## Section 1: Run the Conversions (2 minutes)

### Intro (15 seconds)
"Let's start by converting our source files to PDF. We have two types of input:
- Word documents in the Input Word Files folder
- Images of receipts and documents in the Input Images folder

To convert everything, I'll run the demo script which automates both conversions in sequence."

### Action Sequence
1. **Open PowerShell** (show terminal)
   Narration: "I'm opening PowerShell in the TextExtractor folder where all our scripts are located."

2. **Run the demo**
   Command: `.\Demo-PDFExtractionPipeline.ps1`
   Narration: "Now I'll execute the demo script. This will convert all Word documents and images to PDF automatically."

3. **Show Word conversion output** (30 seconds)
   Narration: "The script first processes all Word documents. As you can see, it's converting 7 .docx files from our Input Word Files folder. It also detects and skips invalid .doc files, which prevents conversion errors. This is a safety feature built into the script."

4. **Show image conversion output** (30 seconds)
   Narration: "Next, the script converts all the images - these are cell phone photos of invoices and receipts - into PDF format using ImageMagick. We see all 5 images converted successfully."

5. **Show PDF output summary** (30 seconds)
   Narration: "The script completes by showing us a summary: all 12 PDFs have been generated and are ready in the Output PDF Files folder. This includes both the Word document conversions and the image conversions."

---

## Section 2: Power Automate Setup (3 minutes)

### Intro (15 seconds)
"Now that we have our PDFs ready, let's set up the Power Automate flow that will use AI Builder to extract text from these documents automatically."

### Step 1: Import the Flow (1 minute)
Narration:
"STEP 1: Import the Power Automate Flow

1. Navigate to make.powerautomate.com and sign in with your Microsoft account
2. Click 'My flows' in the left navigation
3. Select 'Import' and then 'Import package'
4. Choose the PowerAutomateFlow-ExtractPDFs.zip file from the TextExtractor folder
5. Complete the import wizard - it will ask you to configure any connections to cloud services
6. Once imported, the flow will appear in your flows list"

[PAUSE - allow user to follow along]

### Step 2: Configure the Flow (1 minute)
Narration:
"STEP 2: Configure the Flow Path

1. Open the imported flow to edit it
2. Look for the first action called 'Get files in folder'
3. Click on it to expand it
4. In the Location field, paste the path to our Output PDF Files folder:

   C:\Users\franksp\Documents\Clawpilot\TextExtractor\Output PDF Files\

5. This tells the flow where to find all the PDFs we just generated"

[PAUSE - allow configuration]

### Step 3: Review the AI Builder Action (45 seconds)
Narration:
"STEP 3: Review the AI Builder Step

Notice the next action in the flow: 'Extract information from PDFs'. This is where the AI magic happens.

This action:
- Takes each PDF file from the folder
- Uses AI Builder's form processor to analyze the document
- Automatically extracts text and structured data
- Stores the results in a variable called 'ExtractedText'

We're not saving this data anywhere yet - the extraction happens and the results are shown, but storing it in Excel or a database is something you can extend after seeing it work."

---

## Section 3: Run the Flow (2 minutes)

### Intro (15 seconds)
"Now let's run the flow and see it process all our documents in real time."

### Execute Flow (1.5 minutes)
Narration:
"1. Click the 'Save' button to ensure all changes are saved
2. Click 'Test' in the top right
3. Select 'Manually' to trigger the flow
4. Click 'Test' again to start execution

Watch as the flow processes each PDF:
- It lists all files in the Output PDF Files folder
- For each PDF, it runs the AI Builder extraction
- You can see in real-time as documents are being analyzed
- The extracted text appears in the 'ExtractedText' variable

[Watch execution]

Once the flow completes successfully, you'll see green checkmarks next to each step, indicating successful processing."

---

## Section 4: Results and Next Steps (1 minute)

### Show Results (30 seconds)
Narration:
"The flow has successfully extracted text from all 12 documents using AI Builder. The extracted data is stored in the variable for immediate use.

You can see the flow ran without errors - all steps completed successfully."

### Extensibility (30 seconds)
Narration:
"This is where you can extend the solution:

- **Save to Excel**: Add a step to write extracted data to an Excel file on OneDrive or SharePoint
- **Store in a database**: Connect to SQL Server or Dataverse to create records
- **Send notifications**: Email the extracted data to team members
- **Trigger other processes**: Pass the extracted data to other workflows or systems

The foundation is here - the PDF conversion and AI extraction are automated. The storage and downstream actions are entirely customizable to your needs."

---

## Closing (15 seconds)
"And that's the complete PDF extraction pipeline! From raw documents and cell phone photos to intelligent text extraction with AI - all automated and ready for your organization. Thanks for watching!"

---

## Recording Tips:
- Record at 1080p or higher
- Use a clear, steady pace for narration
- Pause after each major step to let viewers follow along
- Highlight important elements on screen (use arrow cursor or zoom)
- Show any error handling or validation messages
- Total recording time: approximately 6-7 minutes
