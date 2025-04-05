Sub AllocateTasks()

    Dim wsData As Worksheet
    Dim wsFilter As Worksheet
    Dim lastRow As Long
    Dim filterRow As Long
    Dim issuerID As String
    Dim analystCount As Integer
    Dim taskCount As Long
    Dim i As Long, j As Long

    ' Set worksheets
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsFilter = ThisWorkbook.Sheets("Filter")

    ' Get the last row of data
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    ' Get the last row of filter list
    filterRow = wsFilter.Cells(wsFilter.Rows.Count, "A").End(xlUp).Row

    ' Count number of analysts (assuming they are in a specific range, adjust as needed)
    analystCount = 8 ' Change this number based on actual analyst count

    ' Loop through each issuer ID in the filter list
    For i = 1 To filterRow

        issuerID = wsFilter.Cells(i, 1).Value

        ' Initialize task count for the current issuer
        taskCount = 0

        ' Find all rows with the current issuer ID and count tasks
        For j = 1 To lastRow
            If wsData.Cells(j, 1).Value = issuerID Then
