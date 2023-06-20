# MetadataPlus V.1.0 - Chris Nevin @ NCCGroup

A tool to extract metadata from Microsoft Office files that includes new locations not checked in other tools.

## Example Usage

### Run on compatible documents in the directory it is in:

`MetaDataPlus.exe`

### Specify input folder:

`MetaDataPlus.exe -i=c:\Docs`

### Run on every file in folder (not just those known to work):

`MetaDataPlus.exe -a`

### Extract images to Media folder (for manual EXIF examingation), and embedded documents to Embed folder (to include manually in later metadata analysis):

`MetaDataPlus.exe -m -e`

### Include user defined search string:

`MetaDataPlus.exe -s=apikey`

### View help file:

`MetaDataPlus.exe -h`
