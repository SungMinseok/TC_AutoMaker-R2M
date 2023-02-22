

Sub 행묶기() '무기이름 등등에 사용
    Dim lastRow As Long
    Dim currRow As Long
    Dim mergeRange As Range
    Dim n, i As Integer
    
    Application.DisplayAlerts = False
    n = 14
    
    ' Find the last row of data in column A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row in column A
    For currRow = 1 To lastRow Step n
        ' Create a range to merge
        Set mergeRange = Range(Cells(currRow, 1), Cells(currRow + n - 1, 1))
        
        ' Merge the cells in the range
        mergeRange.Merge
        
        ' Format the merged cell
        mergeRange.Font.Bold = True
        mergeRange.Font.Size = 9
    Next currRow
    
    ' Sort the data by column A
    'Range("A1:A" & lastRow).Sort key1:=Range("A1")
    'Range("A1:A" & lastRow).Sort key1:=Range("A1"), order1:=xlAscending, Header:=xlYes
    
    Application.DisplayAlerts = True
End Sub

Sub 도감명색칠() '도감명에 색칠하기
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row ' get the last row with a value in column B
    
    For i = 1 To lastRow ' loop through each row from row 1 to the last row with a value in column B
        If Cells(i, "A").Value = "도감명" Then ' if the value in column A of the current row is "book name"
            ' set the font color of columns A and B in the current row to gray
            'Range(Cells(i, "A"), Cells(i, "B")).Font.Color = RGB(128, 128, 128)
            Range(Cells(i, "A"), Cells(i, "B")).Interior.Color = RGB(192, 192, 192)
        End If
    Next i
End Sub
Sub 아이디추출하기()
    Dim lastRow As Long
    Dim i As Long
    Dim numString As String
    Dim pipePos As Integer
    Dim num As String
    
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row ' get the last row with a value in column B
    
    For i = 1 To lastRow ' loop through each row from row 1 to the last row with a value in column B
        numString = Cells(i, "B").Value ' get the value in column B of the current row
        pipePos = InStr(numString, "|") ' find the position of the "|" character in the value
        If pipePos > 0 Then ' if "|" is found
            num = Mid(numString, pipePos + 1) ' get the substring after the "|"
            Cells(i, "C").Value = num ' write the extracted number to column C of the current row
        End If
    Next i
End Sub
Sub fillColors()
    Dim lastRow As Long
    Dim currRow As Long
    Dim setCount As Long
    
    ' Find the last row of data in column A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row in column A
    For currRow = 1 To lastRow
        ' Check if this is the start of a new set of rows
        If currRow Mod 14 = 1 Then
            ' Increment the set count
            setCount = setCount + 1
            
            ' Fill the 1st, 8th, and 14th rows of the set with colors
            Cells(currRow, 1).Resize(1, 1).Interior.Color = RGB(192, 192, 192) ' Light gray
            Cells(currRow + 7, 1).Resize(1, 1).Interior.Color = RGB(155, 194, 230) ' Light blue
            Cells(currRow + 13, 1).Resize(1, 1).Interior.Color = RGB(183, 225, 205) ' Light green
        End If
    Next currRow
End Sub




Sub highlightPVP() '능력치중 PVP포함 뒤졲 빨간색칠
    Dim lastRow As Long
    Dim currRow As Long
    Dim cellText As String
    Dim pvpPos As Long
    Dim digit As String
    Dim digitPos As Long
    
    ' Find the last row of data in column A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row in column A
    For currRow = 1 To lastRow
        ' Get the text in the current cell
        cellText = Cells(currRow, 1).Value
        
        ' Check if [PVP] is in the cell
        pvpPos = InStr(1, cellText, "PVP")
        If pvpPos > 0 Then
            ' Change the font color of the text after [PVP] to red
            Cells(currRow, 1).Characters(pvpPos - 1, Len(cellText)).Font.Color = RGB(255, 0, 0)
            Cells(currRow, 1).Characters(pvpPos - 1, Len(cellText)).Font.Bold = True
        End If
        
        
        
        
    Next currRow
End Sub


Sub boldSlashesInCells()
    Dim lastRow As Long
    Dim currRow As Long
    Dim cellValue As String
    Dim slashPos As Long
    
    ' Find the last row of data in column A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row in column A
    For currRow = 1 To lastRow
        ' Get the value of the current cell
        cellValue = Cells(currRow, 1).Value
        
        ' Check if the cell contains "/"
        If InStr(1, cellValue, "/") > 0 Then
            ' Loop through each slash in the cell
            For slashPos = 1 To Len(cellValue)
                ' Check if the current character is a slash
                If Mid(cellValue, slashPos, 1) = "/" Then
                    ' Set the font of the slash to bold
                    Cells(currRow, 1).Characters(slashPos, 1).Font.Bold = True
                    Cells(currRow, 1).Characters(slashPos, 1).Font.Color = RGB(255, 0, 0)
                End If
            Next slashPos
        End If
    Next currRow
End Sub


Sub boldAndColorNumbersInCells()
    Dim lastRow As Long
    Dim currRow As Long
    Dim cellValue As String
    Dim digit As String
    Dim digitPos As Long
    
    ' Find the last row of data in column A
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row in column A
    For currRow = 1 To lastRow
        ' Get the value of the current cell
        cellValue = Cells(currRow, 1).Value
        
        ' Loop through each character in the cell
        For digitPos = 1 To Len(cellValue)
            ' Check if the current character is a digit
            If IsNumeric(Mid(cellValue, digitPos, 1)) Then
                ' Get the digit
                digit = Mid(cellValue, digitPos, 1)
                
                ' Set the font of the digit to bold and red
                Cells(currRow, 1).Characters(digitPos, 1).Font.Bold = True
                Cells(currRow, 1).Characters(digitPos, 1).Font.Color = RGB(0, 0, 255)
            End If
        Next digitPos
    Next currRow
End Sub


Function SortValues(str As String, rules As String) As String
    Dim arr As Variant, i As Long, j As Long, ruleArr As Variant
    arr = Split(str, "/")
    ruleArr = Split(rules, ",")
    
    For i = 0 To UBound(ruleArr)
        For j = 0 To UBound(arr)
            If Left(arr(j), Len(ruleArr(i))) = ruleArr(i) Then
                ' Swap current element with first element that matches the rule
                Dim temp As String
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' Combine sorted array into single string with "/"
    SortValues = Join(arr, "/")
End Function

Function SortValues2(ByRef cell As Range, ByVal rules As String) As String
    
    Dim values() As String
    values = Split(cell.Value, "/")
    
    Dim sortedValues() As String
    ReDim sortedValues(0 To UBound(values))
    
    Dim rule() As String
    rule = Split(rules, ",")
    
    Dim i As Long, j As Long
    Dim numSorted As Long
    numSorted = -1 ' index of the last sorted value
    
    For j = 0 To UBound(rule)
        For i = 0 To UBound(values)
            If Left(values(i), Len(rule(j))) = rule(j) Then
                numSorted = numSorted + 1
                sortedValues(numSorted) = values(i)
            End If
        Next i
    Next j
    
    ' Add unsorted values to the end
    For i = 0 To UBound(values)
        If InStr(1, Join(sortedValues, "/"), values(i)) = 0 Then
            numSorted = numSorted + 1
            sortedValues(numSorted) = values(i)
        End If
    Next i
    
    SortValues2 = Join(sortedValues, "/")
        
    
End Function





Sub 능력치스트링정렬() '능력치스트링를 인게임방식으로정렬
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    Dim rules As String
    
    rules = "물리공격,마법공격,공속,물리명중,마법명중,힘,민첩,지능,HP회복,MP회복,치명타공격"
    
    For i = 1 To lastRow
        If InStr(1, Cells(i, 1).Value, "/") > 0 Then
            '규칙입력'
            Cells(i, 1).Value = SortValues2(Cells(i, 1), rules)
        End If
    Next i
End Sub

Sub ColorOddRows()
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row ' get the last row with a value in column A
    
    For i = 1 To lastRow ' loop through each row from row 1 to the last row with a value in column A
        If i Mod 2 <> 0 Then ' if the row number is odd
            Range("A" & i & ":A" & i).Interior.Color = RGB(222, 222, 222) ' set the fill color of columns A and B for the current row to light gray
        Else
            Range("A" & i & ":A" & i).Interior.Color = RGB(255, 255, 255) ' set the fill color of columns A and B for the current row to light gray
        
        End If
    Next i
End Sub


