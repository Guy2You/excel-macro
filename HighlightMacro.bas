Sub highlightData()

	Dim sheetToHighlight As WorkSheet
	Dim sheetWithSearchValues As WorkSheet
	Set sheetToHighlight = Sheets("")
	Set sheetWithSearchValues = Sheets("")

	Dim searchValuesColumn As Integer
	searchValuesColumn = -1
	Const searchValuesStartRow As Integer = -1
	Const searchValuesEndRow As Integer  = -1

	Dim highlightValuesColumn As Integer
	Dim highlightRowStart As Integer
	Dim highlightRowEnd As Integer
	Dim highlightColumnStart As Integer
	Dim highlightColumnEnd As Integer
	highlightValuesColumn = -1
	highlightRowStart = -1
	highlightRowEnd = -1
	highlightColumnStart = -1
	highlightColumnEnd = -1

	'SET VARIABLES ABOVE


	'UNHIGHLIGHT ALL ROWS IN SHEET TO HIGHLIGHT
	Dim range As Range
	Set range = sheetToHighlight.Range(sheetToHighlight.Cells(highlightRowStart, highlightColumnStart), sheetToHighlight.Cells(highlightRowEnd, highlightColumnEnd))
	With range.Interior
		.Pattern = xlNone
		.TintAndShade = 0
		.PatternTintAndShade = 0
	End With
	

	'GET LIST OF SEARCH VALUES
	Dim searchValues(searchValuesStartRow To searchValuesEndRow) As String
	For row = searchValuesStartRow To searchValuesEndRow
		searchValues(row - searchValuesStartRow) = sheetWithSearchValues.Cells(row, searchValuesColumn)
	Next row

	'ITERATE OVER LIST OF SEARCH VALUES AND FIND MATCHING ROWS IN HIGHLIGHT SHEET THEN HIGHLIGHT ROWS
	For row = highlightRowStart To highlightRowEnd
		If IsInArray(sheetToHighlight.Cells(row, highlightValuesColumn), searchValues) Then
			Set range = sheetToHighlight.Range(sheetToHighlight.Cells(row, highlightColumnStart), sheetToHighlight.Cells(row, highlightColumnEnd))
			With range.Interior
				.Pattern = xlSolid
				.PatternColorIndex = xlAutomatic
				.Color = 65535
				.TintAndShade = 0
				.PatternTintAndShade = 0
			End With
		End If
	Next row

End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
