vvbb
Sub AllocateTasks()

    Dim wsData As Worksheet
    Dim wsFilter As Worksheet
    Dim lastRow As Long
    Dim filterRow As Long
    Dim issuerID As String
    Dim analystCount As Integer
    Dim taskCount As Long
    Dim i As Long, j As Long
    Dim analystIndex As Integer

    ' Set worksheets
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsFilter = ThisWorkbook.Sheets("Filter")

    ' Get the last row of data and filter list
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    filterRow = wsFilter.Cells(wsFilter.Rows.Count, "A").End(xlUp).Row

    ' Set number of analysts
    analystCount = 8 ' Change this to your actual number of analysts

    ' Make sure there's a column for the assignment (assumed Column C here)
    wsData.Cells(1, 3).Value = "Assigned To" ' Header for assignment

    ' Loop through each Issuer ID in the filter list
    For i = 2 To filterRow ' Assuming Filter tab has a header in row 1

        issuerID = wsFilter.Cells(i, 1).Value
        analystIndex = 1

        ' Loop through data and assign analysts round-robin
        For j = 2 To lastRow ' Assuming Data sheet has a header in row 1
            If wsData.Cells(j, 1).Value = issuerID Then
                wsData.Cells(j, 3).Value = "Analyst " & analystIndex
                analystIndex = analystIndex + 1
                If analystIndex > analystCount Then
                    analystIndex = 1
                End If
            End If
        Next j
    Next i

    MsgBox "Task allocation complete!", vbInformation

End Sub
