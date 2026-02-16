# video file manager - Information
This readme covers the information contained in the "video file manage assistant" python script

## Introduction
Ripping/organizing files is a bit of a pain so I made this to help organize and move files. This markdown file is a work in progress but at least I wanted to cover some of the key points before uploading to github.

In summary, there's two primary functions:
- Export video runtime and file properties
- Copy/Move existing files to new location

# Dependencies
## Dependencies
TBD list library dependencies and non-standard inclusions (its in the main file as well)

# Script Functions
## Export Video Information
The "Export Properties" button will ask users if they would like the export the properties of a single file, or all files in a directory. After making their choice, the script will next ask for a folder to save the output. After selecting the output location, the srcript will compile the file name, full path, and runtime for the associated file(s).

### Export Type
TBD list the single file vs directory differences and how the program handles it.

### Export Properties File
TBD list the various fields and what they all are

## Copy/Move files
The "Update Files" button will ask for a input file used as a template to copy an existing file to a new location
- File templates are accepted in CSV, XLS, or XLSX formats
- Script expects two fields:
	- "File_Path" -> This is the old file path
	- "New_File_Path" -> This is the new file path - where the file at "old file path" will be moved to
	
When selecting "Update files" users will be asked to find their input template containing the new vs old files. An error check will be performed to ensure that common issues like duplicate outputs, skipped files, etc. are not present in the template. Depending on the severity of the issue, users may be able to proceed with the operation.

### Input Template
TBD list the format for the input template (Required fields)

### Error Check
TBD list the various error conditions in the code and what/how it handles things

# Future Versions
TBD list any known improvements/future versions "TODO" information

# Copyright and licensing information
## GNU GPLv3
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/>.

## Additional disclaimer
The information included is provided "as is", without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose and non-infringement. In no event shall the authors or copyright holders be liable for any claim, damages or other liability, whether in an action of contract, tort or otherwise, arising from, out of or in connection with the software or the use or other dealings in the software or tools included in this repository.
