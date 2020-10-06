Sub copyRelevantData()

    Dim sourceSheet As Worksheet
    Dim destSheet As Worksheet
    Set sourceSheet = Sheets("sourceSheetName")
    Set destSheet = Sheets("destSheetName")
    
    Dim sourceColumnSearchValue As Integer
    Dim destRowStartSearchValue As Integer
    Dim destRowEndSearchValue As Integer
    Dim destColumnSearchValue As Integer
    destRowStartSearchValue = 2
    destRowEndSearchValue = 200
    sourceColumnSearchValue = 2
    destColumnSearchValue = 2
        
    Dim sourceColumnStart As Integer
    Dim sourceColumnEnd As Integer
    sourceColumnStart = 1
    sourceColumnEnd = 3
        
    Dim destColumn As Integer
    destColumn = 1
    
    'SET ALL VARIABLES ABOVE
    
    Dim copiedRange As Range
    Dim searchRange As Range
    Dim pasteRange As Range
    Set searchRange = destSheet.Range(destSheet.Cells(destRowStartSearchValue, destColumnSearchValue), destSheet.Cells(destRowEndSearchValue, destRowEndSearchValue))
    
    Dim rowNum As Integer
    For rowNum = 2 To 200
        
        Dim searchFor As String
        searchFor = sourceSheet.Cells(rowNum, sourceColumnSearchValue)
        Set copiedRange = sourceSheet.Range(sourceSheet.Cells(rowNum, sourceColumnStart), sourceSheet.Cells(rowNum, sourceColumnEnd))
 
        If Trim(searchFor) <> "" Then
            Set pasteRange = searchRange.Find(what:=searchFor, _
                                              After:=copiedRange.Cells(copiedRange.Cells.Count), _
                                              LookIn:=xlValues, _
                                              LookAt:=xlWhole, _
                                              SearchOrder:=xlByRows, _
                                              SearchDirection:=xlNext, _
                                              MatchCase:=False)
            If Not pasteRange Is Nothing Then
		sourceSheet.Select
                copiedRange.Select
                Selection.Copy
                
		destSheet.Select
                Range(destSheet.Cells(pasteRange.Row, destColumn), destSheet.Cells(pasteRange.Row, destColumn)).Select
                Selection.PasteSpecial
                
            Else
                'maybe put n/a in the left col
            End If
        End If
    Next rowNum
End Sub
