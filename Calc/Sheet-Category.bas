' Macros/Basic/Calc/Sheets
' Libreoffice - Spreadsheet_Documents
'
' https://wiki.documentfoundation.org/Macros/Basic/Calc/Sheets

Sub Main

	Dim Doc As Object
	Dim Sheet As Object

	Doc = ThisComponent

	' All sheets
	For Each Sheet In Doc.Sheets
	    MsgBox Sheet.Name
	Next
	
	' All Names
	Sheet = ThisComponent.Sheets
	MsgBox Join(Sheet.ElementNames, Chr(13))
	
	'Get by name
	Sheet = ThisComponent.Sheets
	Sheet = Sheet.getByName("Hack - 0")


	Sheet = ThisComponent.Sheets
	Msgbox Sheet.Count
	
End Sub
