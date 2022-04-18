# Word IEEE Referencing
Fixes to Microsoft Word's IEEE bibliography entries to match the IEEE referencing standard according to the University of Sheffield referencing guide.

The currently supported source types are:

- Blog
- Book - Print
- Book - Electronic
- Book - Section/Chapter
- Book - Section/Chapter (Electronic)
- Journal Article - Print
- Journal Article - Electronic
- Podcast
- Standard 
- Standard - Electronic
- Technical Report/Paper
- Technical Report/Paper (Electronic)
- Web Page

## Installation
**NOTE: I did this for Microsoft Word on Windows; I don't know if Word for MacOS uses the same files, and if so, the directories they're stored in. If you manage to get this working on MacOS, please contact me with the file locations so I can add instructions to this guide.**

1. Download the files by clicking 'Code' and then 'Download ZIP' and extract them.
2. Open Notepad as administrator (search "notepad" from the taskbar, right click on it and select 'Run as administrator')
3. Select File -> Open `BIBFORM.XML` file found in `C:\Program Files (x86)\Microsoft Office\root\Office16\1033\Bibliography` (show 'All files'). I recommend that you make a copy of the original, renaming it something like `BACKUP_BIBFORM.XML` or move it elsewhere, in case you need to revert back.
4. Select the entire contents of the file (shortcut ctrl+A) and delete it.
5. Open the downloaded `BIBFORM.XML` in a separate Notepad instance, copy its entire contents and paste it in the original, newly cleared file **(it is important to replace the code inside the original file rather than replacing the whole file because Word will see the new one as an _imposter_)**
6. Select File -> Save
7. Copy `IEEE_UoS.XSL` to `C:\Users\<username>\AppData\Roaming\Microsoft\Bibliography\Style` and `C:\Program Files (x86)\Microsoft Office\root\Office16\Bibliography\Style` directories. You should see the rest of the .XSL template files for other referencing standards in these folders.
8. (Re)start Word and you should see **IEEE - UoS** appear in your bibliography styles in Word.

## Usage

- Select **IEEE - UoS** as your referencing style under the References tab.
- Add a new source
- Select the source type
- Enter the information into important fields - the necessary information for the source type is selected automatically
- If you're referencing an online source, use only URL **or** doi.
- Click OK.

## Source Type Requests or Discrepancies from the Guide

I've made templates for the most commonly used source types; if you want to reference a source that isn't yet supported, contact me at ltaylor19@sheffield.ac.uk and I'll make a template.

Likewise, if you notice an error in any of the bibliography entries that this generates, let me know at the above address and I'll fix it.

## Roughly How it Works (if you're interested in adding source types)
When creating a new source in Word, you must select a source type, this is chosen from a list of source types kept in `BIBFORM.XML`. New source types can be creating by adding `<Source>` Elements. These source types are universal across different referencing standards.

Each referencing standard has a templates file which tells Word how to format the bibliography entries, i.e. `IEEE_UoS.XSL`. These templates determine the information displayed in the bibliography entries, the order, style etc.

