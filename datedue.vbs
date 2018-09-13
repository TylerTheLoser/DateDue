Set fso = CreateObject("Scripting.FileSystemObject")

'the purpose of this script is to format the date after Date Due:, which is currently formatted as YYYYMMDD, to MM/DD/YYYY
'example: Date Date:20180605 should be Date Due:06/05/2018
'created by tyler lagasse
Set re = New RegExp
'start looking for a line in the file that contains the string "Date Due:"
'followed by storing the first four digits (the year), the next two digits (the month), and the last two digits(the day)
re.Pattern = "(\nDate Due:)(\d{4})(\d{2})(\d{2})"
re.Global = True
re.IgnoreCase = True

For Each f in fso.GetFolder("C:\Users\tgl\Desktop\TestFolder\").Files 'change the path where your .dat files are 
	If LCase(fso.GetExtensionName(f.Name)) = "dat" Then 'if it is a .dat file
		text = f.OpenAsTextStream.ReadAll 'reading the text file
		'takes the stored line, and replaces it using the stored variables noted above, as well as adding those pesky slashes
		f.OpenAsTextStream(2).Write re.Replace(text, "$1$3/$4/$2")
	End If
Next