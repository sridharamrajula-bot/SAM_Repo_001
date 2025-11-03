Sub SaveSheetAsImage(sheetName As String, imagePath As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRowA As Long, lastRowH As Long, lastRow As Long
    'Dim co As ChartObject

    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Activate

    ' Get last used rows in column A and H
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowH = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    ' Determine the maximum row between A and H
    lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowH)

    ' Define range from A1 to H:lastRow
    Set rng = ws.Range("A1:H" & lastRow)
    
    ' Copy picture (clipboard only – no chart)
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    '--- ChartObject logic is commented out ---
    'Set co = ws.ChartObjects.Add(Left:=rng.Left, Top:=rng.Top, Width:=rng.Width + 10, Height:=rng.Height + 10)
    'co.Chart.Paste
    'co.Chart.Export Filename:=imagePath, FilterName:="PNG"
    'co.Delete
    '------------------------------------------
End Sub
