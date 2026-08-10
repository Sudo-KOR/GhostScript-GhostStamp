# GhostScript-GhostStamp
This is .VBS file which allows setting text onto a PDFs for Variable Data Printing (VDP) set up, using GhostScript, which needs to be installed before running.

Data is driven from a .CSV file which places each row of text onto the input PDF, so with 4 lines of data you would get 4 PDFs in the output folder.

## How-to
Just copy the main .VBS file to your computer, then you need three files for it to work, make sure they are named the following:

 - input.pdf (your PDF template, with artwork)
 - data.csv (single column, with each variable data per line)
 - font.ttf (the custom font you want to stamp the data with)

Run the .VBS file and it'll create an output folder with each stamped PDF file.

## NOTEs:
This is set to put text central, so if you need text moved else where you'll need to alter the postscript code under the "makePDFs" sub. Colour can also be changed here in the postscript at "setfont 0.0 0.0 0.0 0.0 setcmykcolor" where each 0.0 is Cyan, Magenta, Yellow and Black.

The font name is used by GhostScript for setting the text, so a key part of this script is to extract the font name from the .TTF file, so it must have a valid PostScript name.

This script will pickup the latest version of GhostScript you have installed.
