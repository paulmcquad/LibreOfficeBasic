REM  *****  BASIC  *****

' Libreoffice - Spreadsheet_Documents
' 
' https://wiki.documentfoundation.org/Documentation/BASIC_Guide#Spreadsheet_Documents

Sub Main

	Dim Doc As Object
	Dim Sheet As Object

	Doc = ThisComponent
	Sheet = Doc.createInstance("com.sun.star.sheet.Spreadsheet")
  	
  	
  	' Doc.Sheets.insertByName("Hack - 1", Sheet)

		' Creating Sheets  	
  		If Doc.Sheets.hasByName("Sheet1") Then
  		Sheet = Doc.Sheets(0)
		Sheet.Name = "Hack - 0"
		sheet.TabColor = RGB(0, 255, 0)

	  	Doc.Sheets.insertNewByName("Hack - 1", 1)	
	  	Doc.Sheets.insertNewByName("Hack - 2", 2)
	  	Doc.Sheets.insertNewByName("Hack - 3", 3)
	  	End If	
  	'Sheet = Doc.Sheets(0)
	'Sheet.Name = "Hack"
		
End Sub
