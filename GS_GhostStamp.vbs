'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'|-------------------------------------------------------------|'
'|                                                             |'
'|                         GHOST STAMP                         |'
'|                                                             |'
'|   Add text to PDF based on the contents of the .CSV input,  |'
'|   using a custom .ttf font.                                 |'
'|                                                             |'
'|                       By KORInc, 2026                       |'
'|                            V1.0.3                           |'
'|                                                             |'
'|                       REQUIRED FILES:                       |'
'|   - data.cvs (one column for one store name per line)       |'
'|   - input.pdf (template that gets stamped with the text)    |'
'|   - font.ttf (desired font for the text)                    |'
'|                                                             |'
'|-------------------------------------------------------------|'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'On Error resume Next
Option Explicit

Dim thisDir, flagFileLoc, pathToOutp, pathToInput, pathToData, pathToFont, errorList, gsProg, splitText

' get the directory of this script
thisDir = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

' flag file name
flagFileLoc = thisDir & "\\AddText_Flag"	

' folder locations
pathToOutp = thisDir & "\\output\\"

' pdf input file
pathToInput = thisDir & "\\input.pdf"

' data location
pathToData = thisDir & "\\data.csv"

' font name
pathToFont = thisDir & "\\font.ttf"


Dim pathToGS, exeToGS
' path to ghostscript executable, calculated on run
gsProg = "C:\Program Files\gs\"
pathToGS = ""
exeToGS = "gswin64c.exe"
' text for using to split folder (folder cannot contain this text)
splitText = "<:>"

'---------------------------------------------------------------
'---------------------- CUSTOM FUNCTIONS -----------------------
'---------------------------------------------------------------

Dim fontPSName
fontPSName = ""
Function GetPostScriptFontName(ttfPath)
    Dim st, bytes
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1 ' binary
    st.Open
    st.LoadFromFile ttfPath
    bytes = st.Read
    st.Close

    Dim numTables, i, dirBase, nameTableOffset
    numTables = UInt16BE(bytes, 5)
    dirBase = 13
    nameTableOffset = -1

    For i = 0 To numTables - 1
        Dim recPos, tagStr
        recPos = dirBase + i * 16
        tagStr = ChrB2(bytes, recPos) & ChrB2(bytes, recPos+1) & ChrB2(bytes, recPos+2) & ChrB2(bytes, recPos+3)
        If tagStr = "name" Then
            nameTableOffset = UInt32BE(bytes, recPos + 8) + 1
            Exit For
        End If
    Next

    If nameTableOffset = -1 Then
        Err.Raise vbObjectError + 1, "GetPostScriptFontName", _
            "No 'name' table found in font file: " & ttfPath
    End If

    Dim count, stringOffset, recBase
    count = UInt16BE(bytes, nameTableOffset + 2)
    stringOffset = UInt16BE(bytes, nameTableOffset + 4)
    recBase = nameTableOffset + 6

    Dim psName
    psName = ""

    For i = 0 To count - 1
        Dim rp, platformID, nameID, len_, off_, strPos, s
        rp = recBase + i * 12
        platformID = UInt16BE(bytes, rp)
        nameID     = UInt16BE(bytes, rp + 6)
        len_       = UInt16BE(bytes, rp + 8)
        off_       = UInt16BE(bytes, rp + 10)
        strPos = nameTableOffset + stringOffset + off_

        If nameID = 6 Then
            s = DecodeNameString(bytes, strPos, len_, platformID)
            If s <> "" Then
                psName = s
                Exit For
            End If
        End If
    Next

    If psName = "" Then
        Err.Raise vbObjectError + 2, "GetPostScriptFontName", _
            "No PostScript name (ID 6) found in font file: " & ttfPath
    End If

    GetPostScriptFontName = psName
End Function

Function DecodeNameString(bytes, pos, length, platformID)
    Dim result, i, hi, lo
    result = ""
    If platformID = 1 Then
        ' Macintosh platform: 1 byte per character
        For i = 0 To length - 1
            result = result & ChrW(AscB(MidB(bytes, pos + i, 1)))
        Next
    Else
        ' Windows/Unicode platform: 2 bytes per character, big-endian
        For i = 0 To length - 1 Step 2
            hi = AscB(MidB(bytes, pos + i, 1))
            lo = AscB(MidB(bytes, pos + i + 1, 1))
            result = result & ChrW(hi * 256 + lo)
        Next
    End If
    DecodeNameString = result
End Function

Function ChrB2(bytes, pos)
    ChrB2 = Chr(AscB(MidB(bytes, pos, 1)))
End Function

Function UInt16BE(bytes, pos)
    UInt16BE = AscB(MidB(bytes, pos, 1)) * 256 + AscB(MidB(bytes, pos + 1, 1))
End Function

Function UInt32BE(bytes, pos)
    UInt32BE = CDbl(AscB(MidB(bytes, pos, 1))) * 16777216 + _
               AscB(MidB(bytes, pos + 1, 1)) * 65536 + _
               AscB(MidB(bytes, pos + 2, 1)) * 256 + _
               AscB(MidB(bytes, pos + 3, 1))
End Function


Function getFolderList(folderspec)
	Dim fs, f, f1, fc, s
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set f = fs.GetFolder(folderspec)
	Set fc = f.SubFolders
	For Each f1 in fc
		s = s & f1.name
		s = s & splitText
	Next
	getFolderList = s
End Function


Sub makePDFs(pdfStamp, pdfTextAdd)
	Dim objMakeShell, fontSize, wordLength

	' reduce font size based on the characters in the sring
	wordLength = Len(pdfTextAdd)
	fontSize = 80'90 - (wordLength * 1.5)

	' get PDF and add the data centered on it
	Set objMakeShell = WScript.CreateObject ("WScript.Shell")
	objMakeShell.Run("%comspec% /c cd " & chr(34) & pathToGS & chr(34) & " && " & exeToGS & " -dBATCH -dNOPAUSE -sDEVICE=pdfwrite -sFONTPATH=" & chr(34) & thisDir & "\\" & chr(34) & " -sOutputFile=" & chr(34) & pdfStamp & chr(34) & " -c " & chr(34) & "<< /EndPage { 2 eq { pop false } { gsave " & chr(34) & "/" & fontPSName & chr(34) & " findfont " & fontSize & " scalefont setfont 0.0 0.0 0.0 0.0 setcmykcolor (" & pdfTextAdd & ") dup stringwidth pop 2 div currentpagedevice /PageSize get 0 get 2 div exch sub 700 moveto show grestore true } ifelse } bind >> setpagedevice" & chr(34) & " -f " & chr(34) & pathToInput & chr(34))
	Set objMakeShell = Nothing
	
End Sub


Sub handlePDF(fileToUse, text2Add)
	Dim format2Add, curWait
	' remove special characters
	format2Add = regExp.Replace(text2Add, "")

	' make stamp .PDF with the data
	makePDFs pathToOutp&"stamp_"&format2Add&".pdf", text2Add
	WScript.Sleep 500
	' wait for ghost script to finish
	do while (true)
		Dim objWMIGhost, colGhostList
		Set objWMIGhost = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		Set colGhostList = objWMIGhost.ExecQuery ("Select * from Win32_Process Where Name = '"& exeToGS &"'")
		If (colGhostList.Count = 0) Then
			' no instances found exit loop
			Set objWMIGhost = Nothing
			Set colGhostList = Nothing
			exit do
		End if
		' quit loop and ghostscript if waiting too long
		if (curWait > 5) Then
			Dim strComputer, objWMIService, colProcessList, objProcess
			strComputer = "."
			' close any open instances of ghostscript
			Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			Set colProcessList = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '"& exeToGS &"'")
			if (colProcessList.Count > 0) then
				For Each objProcess in colProcessList
					objProcess.Terminate()
				Next
			End if
			set objWMIService = nothing
			set colProcessList = nothing
			' add error to list
			errorList = errorList & text2Add & vbNewline
			' stop the actions on this file
			exit sub
		else
			curWait=curWait+1
		end if
		WScript.Sleep 500
	loop
	
End Sub


'---------------------------------------------------------------
'---------------------- START MAIN SCRIPT ----------------------
'---------------------------------------------------------------

Dim flagFSO, makeFSO, makeFile, dirFSO, objTempF, objOutpF, regExp, objFSO, objTextFile, delFSO, delFile, txtSplit, txtUpper
Dim gsFolderList, folderArr, folderSingle, checkProgVer, curProgVer, GSVersion, gsi

' check to see if the flag file is present
Set flagFSO = CreateObject("Scripting.FileSystemObject")
If (flagFSO.FileExists(flagFileLoc)) Then
	' flag file found stop the attempted run of script (already running)
	msgbox "SCRIPT ALREADY RUNNING!!" + vbNewline + vbNewline + "If you think this is an error, use ctrl, alt + del and check for wscript.exe in processes." + vbNewline + vbNewline + "If present the script is already running, if not then remove the flag file located in the same folder as the script and the script will start as normal."
	WScript.Quit
elseif NOT(flagFSO.FileExists(pathToData)) Then
	' check data file is present
	msgbox "NO DATA FILE FOUND!!" + vbNewline + vbNewline + "Please add a data file in this folder." + vbNewline + vbNewline + "The data file needs to be a .CSV (data.csv) in this folder and one string per line."
	WScript.Quit
elseif NOT(flagFSO.FileExists(pathToInput)) Then
	' check input pdf file is present
	msgbox "NO PDF FILE FOUND!!" + vbNewline + vbNewline + "Please add the input .PDF file to this folder." + vbNewline + vbNewline + "The input file needs to be a .PDF and the data will get printed on the center."
	WScript.Quit
elseif NOT(flagFSO.FileExists(pathToFont)) Then
	' check font file is present
	msgbox "NO FONT FILE FOUND!!" + vbNewline + vbNewline + "Please add the font file to this folder." + vbNewline + vbNewline + "The font file needs to be a .TTF and in this folder."
	WScript.Quit
else
	' no flag file make it
	Set makeFSO = CreateObject("Scripting.FileSystemObject")
	Set makeFile = makeFSO.CreateTextFile(flagFileLoc, True)
	makeFile.Close
end if

' set up output folder
Set dirFSO = CreateObject("Scripting.FileSystemObject")
If Not dirFSO.FolderExists(pathToOutp) Then
	Set objOutpF = dirFSO.CreateFolder(pathToOutp)
End If

' set up regex for removing non valid file characters
Set regExp = New RegExp
regExp.IgnoreCase = True
regExp.Global = True
regExp.Pattern = "[^a-z0-9!@]"

errorList = ""

' get fontname from ttf file
fontPSName = getPostScriptFontName(pathToFont)

' get a list of folders in this folder
gsFolderList = getfolderlist(gsProg)
' check sub folders present
if not IsEmpty(gsFolderList) then
	' remove split chars at end of string
	gsFolderList = LEFT(gsFolderList, (LEN(gsFolderList)-3))
	' split into each folder name
	folderArr = split(gsFolderList, splitText)
	curProgVer = 0
	
	' loop each gs folder
	For Each folderSingle In folderArr
		checkProgVer = RIGHT(folderSingle, (LEN(folderSingle)-2)) ' remove GS from folder name
		checkProgVer = CInt(REPLACE(checkProgVer, ".", "")) ' remove any decimals
		if (checkProgVer > curProgVer) then ' check version higher
			curProgVer = checkProgVer ' set higher version to check
			GSVersion = gsi ' set lookup index
		End if
		gsi = gsi + 1
	Next
	
	' set version path for ghostscript
	pathToGS = gsProg & folderArr(GSVersion) & "\bin\"
	
Else
	msgbox "Ghost Script not found, closing..."
	' quit the script
	WScript.Quit
end if

' loop through the data file for text to add to PDF, one PDF per line
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(pathToData, 1)
' loop to end of csv file
do while not (objTextFile.AtEndOfStream)
	' remove any brackets in the data
	txtSplit = Split(objTextFile.ReadLine,"(")
	' only use row if data present
	if (uBound(txtSplit) => 0) then
		txtUpper = UCase(txtSplit(0))
		' create the new stamp from data and add to pdf template
		handlePDF pathToInput, txtUpper
		WScript.Sleep 1000
	end if
Loop
objTextFile.close


' remove flag file
if (errorList = "") then
	msgbox "Ghost Stamp Complete." & vbNewline & vbNewline & "No Errors"
else
	msgbox "Ghost Stamp Complete." & vbNewline & vbNewline & "Errors:" & vbNewline & errorList
end if
Set delFSO = CreateObject("Scripting.FileSystemObject")
delFile = delFSO.DeleteFile(flagFileLoc, True)
' quit the script
WScript.Quit

'---------------------------------------------------------------
'----------------------- END OF SCRIPT -------------------------
'---------------------------------------------------------------