# video file manager - Information
This readme covers the information contained in the "video file manage assistant" python script

## Introduction
Ripping/organizing files is a bit of a pain so I made this to help organize and move files. This markdown file is a work in progress but at least I wanted to cover some of the key points before uploading to github.

In summary, there's two primary functions:
- Export video runtime and file properties
- After exporting, users can use a spreadsheet or CSV editor to designate updated names and/or locations
- Copy or Move existing files to the new location(s)

## Future Versions
No future planned versions

## Known Issues / Bugs / Open Items
No current known issues

# Dependencies and Licensing
## Dependencies
The below are non-standard python library dependencies that will have to be installed to modify or run the python script (also also to build an executable).
- cv2
- pandas
- openpyxl (pandas open excel required)

# Script Functions
## Export Video Information
The "Export Properties" button will ask users if they would like the export the properties of a single file, or all files in a directory. After making their choice, the script will next ask for a folder to save the output. After selecting the output location, the srcript will compile the file name(s), full path(s), and runtime(s) for the associated file(s) selected, or those that are present in the directory (and sub-directories) selected.

This file can then be modified and used to generate a new file path for migrating files.

### Export Properties File
When parsing a file or directory, the following information is generated for the user:
- File_Name
    - The file name and extension found in the directory (and its sub-directories) selected by the user
- File_Path
    - Full file path of the files found in the directory (and its sub-directories) selected by the user
- Good_Estimate
    - True/False field indicating if the runtime estimate was a success
    - Based on a confidence score through the MKV processing
- Runtime_Seconds
    - Total runtime in seconds
- Runtime_HH:MM:SS
    - Total runtime formatted in Hours:Minutes:Seconds
- New_File_Path
    - Blank at the property export, used as a template for users to migrate the listed files

## Copy/Move files
The "Update Files" button will ask for a input file used as a template to copy an existing file to a new location
- File templates are accepted in CSV, XLS, or XLSX formats
- Script expects two fields:
	- "File_Path" -> This is the old file path
	- "New_File_Path" -> This is the new file path - where the file at "old file path" will be moved to

Users are able to select two options when migrating files:
- Move
    - This is a simple move that takes the file input path and moves the file to the output file path
    - Additionally, if a new name is given (via the provided full path) that file is also renmaed at the new location
- Copy
    - This does not modify the input file path and creates a copy to the new file path
    - Additionally, if a new name is given (via the provided full path) that file is also renmaed at the new location
    - Users can also select "Delete old files". Selecting this option will remove the files at the old file path after they are copied to the new location(s).

Some things to consider when choosing options:
- If selecting "copy" one input file can be copied to one or multiple output locations. The "copy" is done before the delete
- If selecting "move" one input file may only go to one destination, as it will no longer exist at the input file path after the move.

Clicking on "update files" will perform the uptades in the selected input sheet. Before performing the updates, an error check will be performed to ensure that common issues like duplicate outputs, skipped files, etc. are not present in the template. Depending on the severity of the issue, users may be able to proceed with the operation. See the "Error Check" section of this guide for the various checks performed.

### Input Template
Included in the GIT repository is a xlsx file that demonstrates how users could define updated directories and outputs using formulas.

When updating the files using the tool, two fields are required, listed below. Any additional fields are ignored
Source file path = 'File_Path'
Destination file path = 'New_File_Path'

### Error Check
Before performing updates, the following error checks are performed. Depending on the severity of the issue, users may be able to proceed with the operation.
- All modes: Unable to find input file
    - Error, cannot proceed
    - Identifies records where the "File_Path" field is blank or not a valid file path.
- All modes: Output file is blank
    - Warning, record will be skipped
    - Identifies records where the "New_File_Path" field is blank
- All modes: Output file already exists
    - Error, cannot proceed
    - Identifies records where the file at the "New_File_Path" already exists
- All modes: Duplicate output files
    - Error, cannot proceed
    - Error, Identifies duplicate entries in the "New_File_Path" field
- Move Operation: Multiple input files
    - Error, cannot proceed
    - If the update is a "move" and not a "copy", identifies duplicate entries in the "File_Path" field.
    - Because the file is being moved, after the first entry is updated, any subsequent entries back to the original file path will not be able to be found.

# Copyright and licensing information
## GNU GPLv3
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/>.
