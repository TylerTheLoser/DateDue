'the purpose of this script is to format the date after Date Due:, which is currently formatted as YYYYMMDD, to MM/DD/YYYY
'example: Date Date:20180605 should be Date Due:06/05/2018
'DO NOT USE THIS SCRIPT IN ANY NETWORK DRIVES. THIS IS INTENDED TO BE USED IN A SINGLE FOLDER ON THE USERS DESKTOP/DOCUMENTS
'written by tyler lagasse
'selectfolder function written by rob van der woude
Set fso = CreateObject("Scripting.FileSystemObject")

Dim strPath

strPath = SelectFolder( "" )

If strPath = vbNull Then
	WScript.Echo "Script Cancelled - No files have been modified" 'if the user cancels the open folder dialog
Else
	WScript.Echo "Selected Folder: """ & strPath & """" 'prompt that tells you the folder you selected
End If

Function SelectFolder( myStartFolder )

	Dim objFolder, objItem, objShell
	Dim objFolderItems
	
	On Error Resume Next
	SelectFolder = vbNull
	
	Set objShell = CreateObject( "Shell.Application" )
	Set objFolder = objShell.BrowseForFolder( 0, "Please select the .dat file location folder", 0, myStartFolder)
	set objFolderItems = objFolder.Items
	If IsObject( objFolder ) Then SelectFolder = objFolder.Self.Path
	
	Set objFolder = Nothing
	Set objShell = Nothing
	
	On Error Goto 0
	
End Function

Set re = New RegExp
'start looking for a line in the file that contains the string "Date Due:"
'followed by storing the first four digits (the year), the next two digits (the month), and the last two digits(the day)
re.Pattern = "(\nDate Due:)(\d{4})(\d{2})(\d{2})"
re.Global = True
re.IgnoreCase = True
For Each f in fso.GetFolder(strPath).Files 'change the path where your .dat files are 
	If LCase(fso.GetExtensionName(f.Name)) = "dat" Then 'if it is a .dat file
		text = f.OpenAsTextStream.ReadAll 'reading the text file
		'takes the stored line, and replaces it using the stored variables noted above, as well as adding those pesky slashes
		f.OpenAsTextStream(2).Write re.Replace(text, "$1$3/$4/$2")
		count = count + 1 'Count each file modified
	End If
Next

WScript.Echo count & " " & "files have been modified" 'Tells the number of files that have been modified