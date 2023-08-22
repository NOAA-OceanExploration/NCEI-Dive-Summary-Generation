# NCEI-Dive-Summary-Generation

This project provides a Google Apps Script to automate the creation of dive summary reports. Follow the instructions below to set up and use the script.

## How to Use

### Open Dive Summary Report

1. Scroll down to the “Sample Collected” header and delete everything between "Samples collected" and "Scientists Involved."

### Setting Up the Google Appscript Project

1. Open your Google Appscript Project.
2. If you haven't already, copy the code from `sample_table_generation.txt`. Note: If you have added this code to a template or source document previously, you can skip this step.
3. Within the Dive Summary, navigate to Extensions > Appscript.
4. Delete the “My function” code so you have a blank workspace.
5. Paste the copied code into the main code editor window.
6. Click on the save button.

### Preparing the Dive Record Spreadsheet

1. Upload the `EX####_DIVE##_DiveReport.xlsx` to Dive Summary Reports Excel to GoogleSheet Folder.
2. Convert the spreadsheet to a Google Sheet:
   - Open the spreadsheet.
   - Go to File > Save as Google Sheets.
3. Once the document is saved as a Google Sheet, open it. Look at the URL and make a note of the unique ID. Unique ID is between the `d/[UniqueID]/edit`.

### Editing the Code

1. In the Dive Summary, go to Extensions > Appscript.
2. Find line 25 in the code editor.
3. Replace the existing ID with the unique ID you noted from the Google Sheet's URL.
4. Save the project.
5. Refresh the Dive Summary Document. This will automatically close the AppScript window.

### Inserting the Sample Tables

1. In your document, locate the section between the headers “Sample Tables” and “Scientists Involved”.
2. Insert `$SAMPLES_TABLES$` between these headers.

### Generating Tables from Spreadsheet Data

1. Navigate to Table Generation > Create Tables from Spreadsheet Data.

### Formatting

1. Locate the Niskin Sampling Summary in your document.
2. Highlight the text “Niskin Sampling Summary” and assign it as “heading 2” : (or just copy and paste the text below)

"Niskin Sampling Summary"


## Code

The code for the automation script is provided in the `sample_table_generation.txt` file. It includes functions for creating tables from spreadsheet data, setting default fonts, inserting summary text, and more.

## Contributing

Feel free to fork this project, submit pull requests, or report issues if you have any suggestions or find any bugs.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.
