'Need a list of all sources we need to make files for
DONE - Create non-duplicate list of all sources within Validation Form


'Loop to be completed for each source.
DONE - Prompt User for client naming abbreviation and save input to variable 'ex. MEDC or NBRO_FL

DONE - Create new folder in users documents section
    Title folder UserNamingInput + " " + PCST Files

For each source

Open a new excel file

  Assign current source FULL name to a variable

  Create short name of source FULL name
    (example - urn:cerner:coding:codingsystem:codeset:72 -> codeset:72)
    'Short name is determined by all text following the ":"
    'We can not use FULL name for everything because the FULL name will not fit on the tabs. The length is too long so we need a short name as reference.

  Create variable and store code system ID = urn:cerner:coding:codingsystem:nomenclature.source_vocab:PTCARE 'This is for easy lookup and maintenance
  Create variable and store code system id = urn:cerner:coding:codingsystem:codeset:72 'This is for easy lookup and maintenance

  Create new tabs with the following names
      a. Index Sheet
      b. Unmapped Codes
      c. Clinical Documentation
      d. Health Maintenance Summary
      e. codeset:72
      f. Nomenclature - Patient Care

   Copy the following tabs to the new file
      a. Unmapped Codes
      b. Clinical Documentation
      c. Health Maintenance Summary 'Will need to create way to standardize format of this sheet. Currently there are multiple "tables" or groups of information in the same sheet. Look into creating dynamic ranges.

  Filter The following tabs to display the current source
      a. Unmapped Codes
      b. Clinical Documentation

'Remove duplicates from main data sheets

       Navigate to Unmapped Codes Sheet
         Format All Cells As tables
         Filter to Remove Blank lines
         Remove Duplicates 'Use Raw Code AND Raw Display columns

       Navigate to Clinical Documentation Sheet
         Format All Cells As tables
         Filter to Remove Blank lines
         Remove Duplicates 'Use EventCode AND EventDisplay columns

       Navigate to Health Maintenance Summary Sheet
         Convert Automatic Satisfiers with missing mappings rows to Table
         Filter to Remove Blank lines
         Remove Duplicates 'Use EVENT_CD AND EVENT_CD_DISP columns


'Populate Nomenclature - Patient Care Sheet
   For Nomenclature - Patient Care

       Navigate to Clinical Documentation Sheet
         Filter to remove blank lines
         Copy all results to the Nomenclature - Patient Care sheet. 'Make sure to align columns. It is likely each column will need to be copied individually.

       Navigate to unmapped codes sheet.
         Filter to only display code system id which contains PTCARE
         Filter to remove blank lines 'Check by code system id column
         Copy results to end of Nomenclature - Patient Cwoare sheet 'Notice results will already be on the sheet. You will need to determine next free blank row and start there. NOTICE - The columns will not directly match up. Expect to paste each column individually to make sure it goes in the correct spot.

       Navigate back to unmapped Codes Sheet
         Remove Code System ID Filter

'Populate codeset:72 sheet
   For codeset:72

       Navigate to Unmapped Codes Sheet
         Filter sheet to codeset:72 'Use full name
         Filter to remove blanks 'Filter by code system id column
           'Use IF THEN statement to negate error incase there are no blanks.
         Copy results to the codeset:72 sheet. 'Again columns may not line up perflectly. Copy columns one at a time to ensure correct placement.

       Navigate to Unmapped Codes Sheet
         Remove code system id Filter

'Taylor what else needs to get filtered out of clinical documentation besides "Free Text"?
       Navigate to Clinical Documentation Sheet
         Filter ControlType column to remove "Free Text" 'will need IF THEN statement to negate error incase free text does not exist.
         Filter to EventCode column to hide blanks 'Will need IF THEN statement to negate error incase there are no blanks.
         Copy results to the codeset:72 sheet. 'Columns will NOT line up. Copy columns one at a time. There will be results on the page, determine first blank line and begin entry at that point.

       Navigate to Clinical Documentation Sheet
         Remove filter from EventCode column

       Navigate to Health Maintenance Summary Sheet
         Create named range for AUTOMATIC satisfiers with missing mappings 'needed to identify correct rows go to the correct place
         Filter sheet to remove any blank lines
         Copy results to codeSet:72 sheet. 'Columns will NOT line up. Copy columns one at a time. There will be results on the page, determine first blank line and begin entry at that point.

       Navigate to Health Maintenance Summary Sheet
         Remove any filters


'Populate the remaining sheets
   For each Code System ID that IS NOT codeset 72 OR PTCARE OR on the Blacklist Then 'Codeset 72 and PTCare have different rules and are grouped with data from multiple tabs.

     Create new sheet 'Use Short Name ex. (codeset:72)

       Navigate to Unmapped Codes Sheet
         Filter data by current Code system ID 'Full name
         Copy results to the correct sheet 'Should be the short name of the current code system id
           'Columns may not match up perfectly. Copy each column one at a time to ensure correct location.
       Navigate to Unmapped Code sheet
         Remove Code system ID Filter

     Next Code 'Next code system ID not already laid out and "Sheeted"

'Delete Extra sheets
  Delete Unmapped Codes Sheet
  Delete Clinical Documentation Sheet
  Delete Health Maintenane Summary Sheet

'Save File After all codes populated
  Save excel file within the folder created
  Name file UserNamingInput + " " + Source + " " + PCST

'This source is finished Repeat all of the above with the next Source
Next source
