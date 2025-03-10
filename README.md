## PURPOSE

Exports the currently selected Excel worksheet as CSV or tab-delimited text,
without changing the filename or format of the currently open Excel workbook.

Dates are exported as "the one true date format", YYYY-MM-DD, and numbers are
exported at full precision, rather than display precision.  Tab-delimited
output does not contain any CSV-related escaping, which is both unnecessary
and usually corrupting, that Excel Save As and Python csv.writer produce.
<B>Tabs are escaped as \\\\\\t in tab-delimited output, but are left as-is in
CSV output (you have been warned)</B>.  Embedded end of line characters
(CR, LF, CRLF) are escaped as \\\\\\r, \\\\\\n, \\\\\\r\\\\\n.  \\\\\\ was
chosen as the escape string, instead of \\, due to \\ and \\\\ appearing in
Windows paths.

<BR>



## BACKGROUND

Excel has the annoying behavior of, when you save an individual worksheet to a
CSV/text file, it:

    A) renames the worksheet to the new file name, truncated if it is too long
    B) changes the currently active file, and file format, to the newly saved
       text file

So, I then need to go and rename the worksheet back to its original name,
then re-SaveAs the file back onto the original workbook as an .xlsx file again,
before I continue working.  This has always been an annoyance.  Recently, I
saved a worksheet to text, went to lunch, came back, and kept working, using
the Save button to continue saving over the file (forgetting that it had been
changed to a text file before lunch).  Since my Remote Desktop bar covers up
the document name, I didn't notice that it was .txt, either.  A few hours
later, I closed Excel, re-opened the file, and discovered that all my
formatting and formulas had been lost.

Determined not to have this happen to me again, internet searches led me to
decide that I could avoid the issue entirely by writing my own text export
function that wouldn't trigger any of the problematic behavior.  I could also
fix some other annoyances while I was at it, such as stopping Excel from
applying CSV escaping to tab-delimited output, as well as fix more destructive
behavior, such as truncating the number of significant digits due to various
cell formatting (Excel silently formats larger scientific notation values to 3
digits on text import, causing you to lose significant figures if you forget
to Format -> General before saving to text).

Thanks to nixda's example code from
https://superuser.com/questions/978228/how-to-export-not-save-as-to-another-format
for giving me a good starting point to extend upon.  The initial posted
question also summarizes the problem nicely:

    "This is a disaster waiting to happen."

<BR>



## METHOD

The subroutine opens an output file requestor, reads through the currently
active Excel worksheet one row at a time, passes each row to Visual Basic
(VBA) as a 2D variant array, escapes each cell as necessary, joins the cells
together into a line, then writes the lines to the selected filename.  It is,
sadly, a good bit slower than Save As on large files, but it is well worth not
worrying about, working around, and correcting for all the various ways that
Excel can corrupt your data when saving to a text file.

Microsoft has generally done a poor job of porting Excel to Mac over the years,
especially with regards to VBA, resulting in Mac versions of Excel typically
missing many useful functions that PC VBA programers rely on.  For example, on
PC, GetSaveAsFilename() can apply a filter to the allowed output formats
(.csv, .txt), but attempting to do so on Mac will abort with an error message,
so Mac users are stuck with scrolling through the whole list of file formats to
select *.txt or *.csv every time they export a file to text.  Even more
problematic for us, Mac Excel was stuck on VBA5 for ~20 years, which meant that
it was missing important string processing functions, such as InStrRev(),
StrReverse(), Join(), and Replace().  Surprisingly, current Office 365 for Mac
finally supports these functions!  This is good, but it also means that I can't
just check for Mac vs. PC when declaring the backwards compatability
functions.  I've supplied optimized backwards compatability functions from
http://www.xbeat.net/vbspeed/, which are declared if VBA6 isn't supported.
Hopefully, this will catch and fix all cases of these functions not
existing....

<BR>


## UTF-8 Unicode support

The only way for Excel to normally auto-detect UTF-8 is to include the UTF-8
BOM in a .csv file.  The BOM must be there, and the filename must end in .csv,
otherwise it will default to Windows-1252.  Excel will also ignore any
arguments you may pass to various VBA functions to open files with any code
page other than Windows-1252 if it A) ends in .csv and B) does not contain a
UTF-8 BOM.  Basically, if a file ends in .csv, Excel thinks it knows better
than you and will ignore most attempts to tell it otherwise.  And, if a file
ends in .txt and has a UTF-8 BOM, Excel will also think it knows better than
what is in the actual file, and continue to import it as default Windows-1252.
Sigh.

ExportAsText exports all text files as UTF-8 without a BOM.  The UTF-8 standard
actually discourages the use of the UTF-8 BOM, probably because many software
packages and functions do not support BOMs, and inclusion of a BOM will corrupt
the file (the first few characters will be non-printable "garbage" that is
difficult to notice due to them being non-printable characters).  I cannot tell
you how many times I have had to deal with UTF-8 BOMs introducing downstream
first field corruption due to lack of UTF-8 BOM awareness.  And the vast
majority of offending BOM-corrupted files are generated by Microsoft software,
such as Excel and Access, since the BOMs are required in order for the same
Microsoft software to read the files back in correctly.  Thus, we have a
problem.  By not outputting UTF-8 BOMs, ExportAsText writes files that Excel
cannot read back in properly without UTF-8 --> Windows-1252 "mojibake"
corruption.

<BR>


## OpenAsUTF8

A function/macro, OpenAsUTF8, has been included to solve all of the above
UTF-8 import problems.  It will import a text file as UTF-8, whether or not it
has a UTF-8 BOM.  Files ending in .csv must be opened with a custom function,
due to OpenText() ignoring most, if not all, of the various options you pass
it when a file ends in .csv.  The custom function uses QueryTable() to load
the data in via a database query, which actually honors all the options that
you set for it.  Thank you EEM of Stackoverflow!  Files not ending in .csv
used to be imported using the standard OpenText() function with the
appropriate arguments.  However, our organization's version 2402 Build
17328.20708 appears to have introduced a new bug into OpenText(), which
results in silent corruption of input fields and/or crashes.  Due to the new
OpenText() bug, *all* text files are opened with the same custom function now,
regardless of file extension.

The new QueryTable-based function needs to set ActiveWindow.Caption to the
name of the file, so that it displays properly in the title bar and task bar
(rather than Book#), and so that ExportAsText can detect the proper default
file name to export.  Unfortunately, setting ActiveWindow.Caption disables the
normal caption behavior of Excel's native SaveAs, so that it no longer changes
the caption to that of the newly saved file.  So, I added additional support to
ExportAsText() for .xlsx files, so that an OpenAsUTF8 imported text file can be
saved as an Excel workbook and restore proper default caption behavior by
setting ActiveWindow.Caption = False afterwards (False, not "", Microsoft's
documentation for Window.Caption is incorrect).  Workarounds for workarounds
for workarounds....

Macro buttons can be added for OpenAsUTF8 by following the same procedures
detailed below, substituting OpenAsUTF8 for ExportAsText where appropriate.

<b><i>!! WARNING -- it is hardcoded to open as UTF-8, even if another BOM is
present !!</b></i>  I may add smarter support for other BOMs in the future.

If you already have existing files with mojibake corruption, I highly recommend
the Python ftfy package for recovering the proper Unicode characters:
(https://pypi.org/project/ftfy/).

<BR>



## INSTALLATION

### Install and activate Add-in on PC:

   1) Open ExportAsText.xlsm
   2) File -> Save As -> More Options -> Save as type: Excel Add-in (*.xlam)
   3) click the Save button
   4) File -> Options (at the very bottom of screen) -> Add-ins
   5) click the Go button next to Manage: Excel Add-ins, near the bottom
   6) check the box to the left of Export as Text
   7) click the OK button to close the Add-ins window


#### De-activate Add-in on PC (so that you can install a newer version)

   1) File -> Options (at the very bottom of screen) -> Add-ins
   2) click the Go button next to Manage: Excel Add-ins, near the bottom
   3) un-check the box to the left of Export as Text
   4) click the OK button to close the Add-ins window
   5) [follow installation instructions above to update with new version]

<BR>


### Install and activate Add-in on Mac:

   1) Open ExportAsText.xlsm
   2) File -> Save As -> More Options -> Save as type: Excel Add-in (*.xlam)
   3) click the Save button
   4) Tools -> Excel Add-ins...
   5) click the Browse button in the Add-ins window
   6) navigate to where you just saved ExportAsText.xlam,
      which should default to the same location as ExportAsText.xlsm
   7) select the ExportAsText.xlam file
   8) click the Open button
   9) check the box to the left of Export as Text
  10) click the OK button to close the Add-ins window


#### De-activate Add-in on Mac (so that you can install a newer version)

   1) Tools -> Excel Add-ins...
   2) un-check the box to the left of Export as Text
   3) click the OK button to close the Add-ins window
   4) [follow installation instructions above to update with new version]

<BR>



### Add button to the Quick Access Toolbar

   1) File -> Options (very bottom of the window) -> Quick Access Toolbar
   2) [right panel] Customize Quick Access Toolbar: For all documents (default)
   3) [left  panel] Choose commands from: Macros
   4) select ExportAsText.xlam!ExportAsText.Exp...
   5) click the "Add > >" button to add it to the Quick Access Toolbar
   6) select the ExportAsText.xlam that you just added
   7) click the Modify... button
   8) choose whichever Symbol: icon you like --
      I like the outline of a piece of paper with its top-right corner folded,
      with a diagonally down arrow in the bottom-right corner
      (10th icon from the left in my version of Excel)
   9) click the OK button to close the Modify Button window
  10) click and hold to drag the newly added "button" to where you want it
      relative to the other existing "buttons" on the toolbar.
      I left it at the end, after Undo and Redo.
  11) click the OK button to close the Quick Access Toolbar window

<BR>


### Add "Export as Text" button to the Ribbon

   1) File -> Options (very bottom of the window) -> Customize Ribbon
   2) [right panel] Customize the Ribbon: Main Tabs
   3) expand the [+] box to the left of Home if it isn't already [-] expanded
   4) select Home under Main Tabs
   5) click the New Group button
   6) select the newly created New Group (Custom) if it isn't selected
   7) click the Rename button
   8) type "Export" in the Display name: box where it says "New Group";
      select any Symbol: icon, it doesn't matter, it won't be visible
   9) click the OK button to close the Rename window
  10) click and hold to drag the Export group up/down to where you want it
      relative to the other existing Home groups.
      I put it after the Clipboard and Font groups.
  11) select the Export (Custom) group you just created and/or moved around
  12) [left panel] Choose commands from: Macros
  13) select ExportAsText.xlam!ExportAsText.Exp...
  14) click the "Add > >" button to add it to the selected Export group
  15) select the ExportAsText.xlam that you just added to the Export group
  16) click the Rename button
  17) type "Export as Text" so that the name will word-wrap on spaces
  18) choose whichever Symbol: icon you like --
      I like the outline of a piece of paper with its top-right corner folded,
      with a diagonally down arrow in the bottom-right corner
      (10th icon from the left in my version of Excel)
  19) click the OK button to close the Rename window
  20) click the OK button to close the Customize the Ribbon window

<BR>


## Usage:

Click on either the newly added Ribbon or Quick Access Toolbar buttons
to run the VBA script and export the current worksheet to either CSV or
tab-delimited text.  Mac users will need to select .csv or .txt from
amongst all possible output types in the pop-up output file name requestor,
since there is no simple way to limit it to only .csv or .txt on Mac.

Sometimes, Visual Basic will throw some obscure error message related to
converting data between types.  So far, whenever this happens, I have been
able to copy / paste values into a new spreadsheet and then export from the
new spreadsheet with no problems.  You may also need to exit / restart Excel,
I don't remember exactly what I needed to do the last time this happened.

<BR>



## AUTHORS

<pre>
Eric A. Welsh (Eric.Welsh@moffitt.org)

See comments above/within vbspeed functions for their respective authors:

  Main author:
    Eric A. Welsh (Eric.Welsh@moffitt.org)

  Original UTF-8 encoding function:
    Tim Perkin

  Original CSV QueryTable function:
    EEM of Stackoverflow

  InStrRev , StrReverse:
    Donald Lessau  (donald@xbeat.net)

  Join:
    Guido Beckmann (G.beckmann@NikoCity.de)
    Keith Matzen   (kmatzen@ispchannel.com)

  Replace:
    Jost Schwider  (jost@schwider.de)
</pre>

<BR>



## License and Copyright

Copyright (C) 2023, Eric A. Welsh (Eric.Welsh@moffitt.org)<BR>
Licensed under the zlib license:

<pre>
This software is provided 'as-is', without any express or implied
warranty. In no event will the authors be held liable for any damages
arising from the use of this software.

Permission is granted to anyone to use this software for any purpose,
including commercial applications, and to alter it and redistribute it
freely, subject to the following restrictions:

1. The origin of this software must not be misrepresented; you must not
   claim that you wrote the original software. If you use this software
   in a product, an acknowledgment in the product documentation would be
   appreciated but is not required.
2. Altered source versions must be plainly marked as such, and must not be
   misrepresented as being the original software.
3. This notice may not be removed or altered from any source distribution.
</pre>
