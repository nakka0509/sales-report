import re

file_path = r'C:\Users\sawak\OneDrive\デスクトップ\売上メール\rebuild_vba.py'
with open(file_path, 'r', encoding='utf-8') as f:
    text = f.read()

# Replace the broken ClearDuplicateCol and ApplyDayStripe with working versions
vba_replacement = """
Sub ClearDuplicateCol(ws As Worksheet, col As Long, startRow As Long, lastRow As Long, Optional parentCol1 As Long = 0, Optional parentCol2 As Long = 0)
    Dim r As Long
    Application.DisplayAlerts = False
    
    Dim prevValCol As String, prevValP1 As String, prevValP2 As String
    prevValCol = ""
    prevValP1 = ""
    prevValP2 = ""
    
    For r = startRow To lastRow
        If ws.Cells(r, 3).Value = "日合計" Or (ws.Cells(r, 4).Value = "" And ws.Cells(r, 3).Value = "") Then
            prevValCol = ""
            prevValP1 = ""
            prevValP2 = ""
            GoTo NextRow
        End If
        
        Dim curValCol As String, curValP1 As String, curValP2 As String
        curValCol = CStr(ws.Cells(r, col).Value)
        If curValCol = "" Then GoTo NextRow
        
        curValP1 = ""
        If parentCol1 > 0 Then
            curValP1 = CStr(ws.Cells(r, parentCol1).Value)
            If curValP1 = "" Then curValP1 = prevValP1
        End If
        
        curValP2 = ""
        If parentCol2 > 0 Then
            curValP2 = CStr(ws.Cells(r, parentCol2).Value)
            If curValP2 = "" Then curValP2 = prevValP2
        End If
        
        If prevValCol = "" Then
            prevValCol = curValCol
            prevValP1 = curValP1
            prevValP2 = curValP2
            GoTo NextRow
        End If
        
        Dim doClear As Boolean
        doClear = True
        
        If curValCol <> prevValCol Then doClear = False
        If parentCol1 > 0 And curValP1 <> prevValP1 Then doClear = False
        If parentCol2 > 0 And curValP2 <> prevValP2 Then doClear = False
        
        If doClear Then
            ws.Cells(r, col).ClearContents
            On Error Resume Next
            ws.Cells(r, col).Borders(xlEdgeTop).LineStyle = xlNone
            On Error GoTo 0
        Else
            prevValCol = curValCol
            prevValP1 = curValP1
            prevValP2 = curValP2
        End If
        
NextRow:
    Next r
    Application.DisplayAlerts = True
End Sub

Sub ApplyDayStripe(ws As Worksheet, startRow As Long, lastRow As Long)
    Dim fillClr As Long: fillClr = RGB(235, 242, 250)
    Dim r As Long, storeIdx As Long, c As Long
    Dim prevStore As String, curStore As String
    storeIdx = 0: prevStore = ""
    
    For r = startRow To lastRow
        If ws.Cells(r, 3).Value = "日合計" Then GoTo NextDR
        If ws.Cells(r, 4).Value = "" And ws.Cells(r, 3).Value = "" Then GoTo NextDR
        
        curStore = CStr(ws.Cells(r, 2).Value)
        If curStore = "" Then curStore = prevStore
        
        If curStore <> prevStore And curStore <> "" Then
            storeIdx = storeIdx + 1
            prevStore = curStore
        End If
        
        If storeIdx Mod 2 = 1 Then
            For c = 1 To 7
                ws.Cells(r, c).Interior.Color = fillClr
            Next c
        End If
NextDR:
    Next r
End Sub
"""

# Extract everything before ClearDuplicateCol string
start_idx = text.find('Sub ClearDuplicateCol(')
end_idx = text.find('Sub ApplyTotalStripe(')

if start_idx != -1 and end_idx != -1:
    new_text = text[:start_idx] + vba_replacement.strip() + '\n\n' + text[end_idx:]
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_text)
    print('Successfully patched rebuild_vba.py')
else:
    print('Failed to find replacement bounds')
