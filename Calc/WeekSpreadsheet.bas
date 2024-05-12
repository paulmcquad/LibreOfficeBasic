REM  *****  BASIC  *****

' Libreoffice - Spreadsheet_Documents
' 
' https://wiki.documentfoundation.org/Documentation/BASIC_Guide#Spreadsheet_Documents

Sub Main

	Dim Doc As Object
	Dim Sheet As Object
	Dim Cell As Object
	
	Doc = ThisComponent
	Sheet = Doc.Sheets(0)
	
	Cell = Sheet.getCellByPosition(0, 0)
	Cell.String = "Monday"

	Cell = Sheet.getCellByPosition(1, 0)
	Cell.String = "Tuesday"
	
	Cell = Sheet.getCellByPosition(2, 0)
	Cell.String = "Wednesday"

	Cell = Sheet.getCellByPosition(3, 0)
	Cell.String = "Thursday"
	
	Cell = Sheet.getCellByPosition(4, 0)
	Cell.String = "Friday"

	Cell = Sheet.getCellByPosition(5, 0)
	Cell.String = "Saturday"
	
	Cell = Sheet.getCellByPosition(6, 0)
	Cell.String = "Sunday"

End Sub
