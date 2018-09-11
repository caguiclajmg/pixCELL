Attribute VB_Name = "modMacro"
Option Explicit

Public Sub ImageToSpreadsheet()
    Dim dialog As FileDialog: Set dialog = Application.FileDialog(msoFileDialogOpen)
    With dialog
        .AllowMultiSelect = False
        .InitialFileName = Application.path
        .Title = "Open Image"
        With .Filters
            .Clear
            .Add "Bitmap (*.bmp)", "*.bmp"
        End With
    End With
    If dialog.Show <> -1 Then Exit Sub
    
    Dim path As String: path = dialog.SelectedItems(1)

    Dim disableUpdate As Boolean: disableUpdate = (MsgBox("Disable screen updating?", vbInformation Or vbYesNo, vbNullString) = vbYes)
    If disableUpdate Then
        With Application
            .Calculation = xlCalculationManual
            .DisplayStatusBar = False
            .ScreenUpdating = False
            .EnableEvents = False
        End With
    End If
    
#If Win64 Then
    Dim bitmapHandle As LongPtr
#Else
    Dim bitmapHandle As Long
#End If
    bitmapHandle = LoadImage(Application.Hinstance, path, 0, 0, 0, &H10)
    Dim bitmapInfo As BITMAP: GetObject bitmapHandle, LenB(bitmapInfo), bitmapInfo
    ReDim bitmapbits(1 To (bitmapInfo.bmBitsPixel / 8), 1 To bitmapInfo.bmWidth, 1 To bitmapInfo.bmHeight) As Byte: GetBitmapBits bitmapHandle, bitmapInfo.bmWidthBytes * bitmapInfo.bmHeight, bitmapbits(1, 1, 1)
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    
    Dim rangeTarget As Range: Set rangeTarget = ws.Range(ws.Cells(1, 1), ws.Cells(bitmapInfo.bmHeight, bitmapInfo.bmWidth))
    With rangeTarget
        .EntireRow.RowHeight = 18
        .EntireColumn.ColumnWidth = 2.43
        .Select
        ActiveWindow.Zoom = True
    End With
    
    Dim y As Long
    For y = 1 To bitmapInfo.bmHeight
        Dim x As Long
        For x = 1 To bitmapInfo.bmWidth
            rangeTarget(y, x).Interior.Color = (CLng(bitmapbits(1, x, y)) * &H10000) Or (CLng(bitmapbits(2, x, y)) * &H100) Or CLng(bitmapbits(3, x, y))
        Next
    Next
    
    DeleteObject bitmapHandle
    
    If disableUpdate Then
        With Application
            .Calculation = xlCalculationAutomatic
            .DisplayStatusBar = True
            .ScreenUpdating = True
            .EnableEvents = True
        End With
    End If
End Sub
