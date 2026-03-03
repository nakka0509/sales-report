import os
import time
import tempfile
import shutil

BASE_DIR = r"C:\Users\sawak\OneDrive\デスクトップ\売上メール"
REPORT_FILE = os.path.join(BASE_DIR, "売上確認.xlsm")

# VBAコードのDBパスは ThisWorkbook.Path から直接取得（日本語文字化けを完全回避）
VBA_CODE = """Option Explicit

Sub FetchSalesData()
    Dim wsS As Worksheet
    Set wsS = ThisWorkbook.Sheets("検索")

    If Not IsDate(wsS.Range("D9").Value) Or Not IsDate(wsS.Range("D10").Value) Then
        MsgBox "開始日と終了日を正しく入力してください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    Dim sd As Date, ed As Date
    sd = CDate(wsS.Range("D9").Value)
    ed = CDate(wsS.Range("D10").Value)
    If sd > ed Then
        MsgBox "開始日は終了日より前にしてください。", vbExclamation, "入力エラー"
        Exit Sub
    End If

    Dim siteFilter As String
    siteFilter = Trim(wsS.Range("D8").Value)
    If siteFilter = "" Then siteFilter = "すべて"

    Dim dp As String, basePath As String
    basePath = ThisWorkbook.Path
    If InStr(1, basePath, "https", vbTextCompare) > 0 Or InStr(1, basePath, "http", vbTextCompare) > 0 Then
        basePath = Environ("USERPROFILE")
        If Dir(basePath & "\OneDrive", vbDirectory) <> "" Then
            basePath = basePath & "\OneDrive"
        End If
        Dim tryPath As String
        tryPath = basePath & "\Desktop"
        If Dir(tryPath, vbDirectory) <> "" Then basePath = tryPath
        tryPath = basePath & "\デスクトップ"
        If Dir(tryPath, vbDirectory) <> "" Then basePath = tryPath
        tryPath = basePath & "\売上メール"
        If Dir(tryPath, vbDirectory) <> "" Then basePath = tryPath
    End If
    dp = ""
    Dim f As String
    f = Dir(basePath & "\*", vbDirectory)
    Do While f <> ""
        If (GetAttr(basePath & "\" & f) And vbDirectory) <> 0 Then
            Dim testXls As String
            testXls = basePath & "\" & f & "\売上管理表.xlsx"
            If Dir(testXls) <> "" Then dp = testXls: Exit Do
        End If
        f = Dir()
    Loop
    If dp = "" Then dp = basePath & "\データベース\売上管理表.xlsx"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(dp) Then
        MsgBox "DBが見つかりません:" & vbCrLf & dp, vbCritical, "エラー"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "データ読込中..."

    Dim wbD As Workbook
    Set wbD = Workbooks.Open(dp, ReadOnly:=True)
    Dim wsD As Worksheet
    Set wsD = wbD.Sheets("Data")
    Dim lr As Long
    lr = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then wbD.Close False: GoTo Cleanup

    Dim d As Variant
    d = wsD.Range("A1:M" & lr).Value
    wbD.Close False

    Dim cD As Long, cSrc As Long, cSt As Long, cIt As Long
    Dim cPr As Long, cQ As Long, cSb As Long, cTmax As Long, cTmin As Long
    Dim j As Long
    For j = 1 To UBound(d, 2)
        Select Case Trim(CStr(d(1, j)))
            Case "日付": cD = j
            Case "取得元": cSrc = j
            Case "店舗名": cSt = j
            Case "品名": cIt = j
            Case "単価": cPr = j
            Case "数量": cQ = j
            Case "小計": cSb = j
            Case "最高気温": cTmax = j
            Case "最低気温": cTmin = j
        End Select
    Next j
    If cSrc = 0 Then MsgBox "取得元列が見つかりません。", vbCritical: GoTo Cleanup

    Dim fc As Long: fc = 0
    Dim i As Long, rd As Date, matchSite As Boolean
    For i = 2 To UBound(d, 1)
        If IsDate(d(i, cD)) Then
            rd = CDate(d(i, cD))
            If rd >= sd And rd <= ed Then
                matchSite = True
                If siteFilter <> "すべて" Then
                    If StrComp(Trim(CStr(d(i, cSrc))), siteFilter, vbTextCompare) <> 0 Then matchSite = False
                End If
                If matchSite Then fc = fc + 1
            End If
        End If
    Next i
    If fc = 0 Then MsgBox "データがありません。", vbInformation: GoTo Cleanup

    Dim fd() As Variant
    ReDim fd(1 To fc, 1 To 8)
    Dim ix As Long: ix = 0
    For i = 2 To UBound(d, 1)
        If IsDate(d(i, cD)) Then
            rd = CDate(d(i, cD))
            If rd >= sd And rd <= ed Then
                matchSite = True
                If siteFilter <> "すべて" Then
                    If StrComp(Trim(CStr(d(i, cSrc))), siteFilter, vbTextCompare) <> 0 Then matchSite = False
                End If
                If matchSite Then
                    ix = ix + 1
                    fd(ix, 1) = rd: fd(ix, 2) = d(i, cSt): fd(ix, 3) = d(i, cIt)
                    fd(ix, 4) = d(i, cPr): fd(ix, 5) = d(i, cQ): fd(ix, 6) = d(i, cSb)
                    If cTmax > 0 Then fd(ix, 7) = d(i, cTmax) Else fd(ix, 7) = ""
                    If cTmin > 0 Then fd(ix, 8) = d(i, cTmin) Else fd(ix, 8) = ""
                End If
            End If
        End If
    Next i

    Application.StatusBar = "シート作成中..."
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("日別成績").Delete
    ThisWorkbook.Sheets("トータル成績").Delete
    ThisWorkbook.Sheets("グラフ").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' ===== 日別成績シート =====
    Dim wsDay As Worksheet
    Set wsDay = ThisWorkbook.Sheets.Add(After:=wsS)
    wsDay.Name = "日別成績"
    wsDay.Cells(1, 1).Value = "日付": wsDay.Cells(1, 2).Value = "店舗名"
    wsDay.Cells(1, 3).Value = "品名": wsDay.Cells(1, 4).Value = "単価"
    wsDay.Cells(1, 5).Value = "数量": wsDay.Cells(1, 6).Value = "小計"
    wsDay.Cells(1, 7).Value = "日別売上合計"

    Dim dicDT As Object, dicDayAll As Object
    Set dicDT = CreateObject("Scripting.Dictionary")
    Set dicDayAll = CreateObject("Scripting.Dictionary")
    Dim k As String, dk As String
    For i = 1 To fc
        k = Format(fd(i, 1), "yyyymmdd") & "|" & CStr(fd(i, 2))
        If dicDT.Exists(k) Then dicDT(k) = dicDT(k) + CDbl(fd(i, 6)) Else dicDT.Add k, CDbl(fd(i, 6))
        dk = Format(fd(i, 1), "yyyymmdd")
        If dicDayAll.Exists(dk) Then dicDayAll(dk) = dicDayAll(dk) + CDbl(fd(i, 6)) Else dicDayAll.Add dk, CDbl(fd(i, 6))
    Next i

    Dim od() As Variant
    ReDim od(1 To fc, 1 To 7)
    For i = 1 To fc
        od(i, 1) = fd(i, 1): od(i, 2) = fd(i, 2): od(i, 3) = fd(i, 3)
        od(i, 4) = fd(i, 4): od(i, 5) = fd(i, 5): od(i, 6) = fd(i, 6)
        k = Format(fd(i, 1), "yyyymmdd") & "|" & CStr(fd(i, 2))
        od(i, 7) = dicDT(k)
    Next i
    wsDay.Range("A2").Resize(fc, 7).Value = od

    With wsDay.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsDay.Range("A2:A" & fc + 1), Order:=xlAscending
        .SortFields.Add Key:=wsDay.Range("B2:B" & fc + 1), Order:=xlAscending
        .SortFields.Add Key:=wsDay.Range("C2:C" & fc + 1), Order:=xlAscending, CustomOrder:="和花,切花"
        .SortFields.Add Key:=wsDay.Range("D2:D" & fc + 1), Order:=xlAscending
        .SetRange wsDay.Range("A1:G" & fc + 1)
        .Header = xlYes: .Apply
    End With

    ' 日合計行挿入（下から処理）
    Dim r As Long
    For r = fc + 1 To 3 Step -1
        If wsDay.Cells(r, 1).Value <> wsDay.Cells(r - 1, 1).Value And wsDay.Cells(r - 1, 1).Value <> "" Then
            wsDay.Rows(r).Insert Shift:=xlDown
            wsDay.Cells(r, 3).Value = "日合計": wsDay.Cells(r, 3).Font.Bold = True
            Dim dayKey As String
            dayKey = Format(wsDay.Cells(r - 1, 1).Value, "yyyymmdd")
            If dicDayAll.Exists(dayKey) Then
                wsDay.Cells(r, 6).Value = dicDayAll(dayKey): wsDay.Cells(r, 6).Font.Bold = True
            End If
            With wsDay.Range(wsDay.Cells(r, 1), wsDay.Cells(r, 7))
                .Interior.Color = RGB(218, 230, 242): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(180, 198, 220)
            End With
            wsDay.Rows(r + 1).Insert Shift:=xlDown
            wsDay.Rows(r + 1).Interior.ColorIndex = xlNone: wsDay.Rows(r + 1).Borders.LineStyle = xlNone
        End If
    Next r

    Dim dataLast As Long
    dataLast = wsDay.Cells(wsDay.Rows.Count, 4).End(xlUp).Row
    If dataLast >= 2 Then
        wsDay.Rows(dataLast + 1).Insert Shift:=xlDown
        wsDay.Cells(dataLast + 1, 3).Value = "日合計": wsDay.Cells(dataLast + 1, 3).Font.Bold = True
        dayKey = ""
        Dim sr As Long
        For sr = dataLast To 2 Step -1
            If wsDay.Cells(sr, 1).Value <> "" Then
                dayKey = Format(wsDay.Cells(sr, 1).Value, "yyyymmdd"): Exit For
            End If
        Next sr
        If dicDayAll.Exists(dayKey) Then
            wsDay.Cells(dataLast + 1, 6).Value = dicDayAll(dayKey): wsDay.Cells(dataLast + 1, 6).Font.Bold = True
        End If
        With wsDay.Range(wsDay.Cells(dataLast + 1, 1), wsDay.Cells(dataLast + 1, 7))
            .Interior.Color = RGB(218, 230, 242): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(180, 198, 220)
        End With
    End If

    Call StyleSheet(wsDay, wsDay.Cells(wsDay.Rows.Count, 4).End(xlUp).Row, 7)
    wsDay.Range("D:G").NumberFormat = "#,##0"

    Dim newLastR As Long
    newLastR = wsDay.Cells(wsDay.Rows.Count, 4).End(xlUp).Row
    Call MergeCol(wsDay, 1, 2, newLastR)
    Call MergeCol(wsDay, 2, 2, newLastR)
    Call MergeCol(wsDay, 3, 2, newLastR)
    Call MergeCol(wsDay, 7, 2, newLastR)
    Call ApplyDayStripe(wsDay, 2, newLastR)

    ' ===== トータル成績シート =====
    Application.StatusBar = "トータル成績作成中..."
    Dim wsTot As Worksheet
    Set wsTot = ThisWorkbook.Sheets.Add(After:=wsDay)
    wsTot.Name = "トータル成績"

    wsTot.Range("A1:E1").Merge
    wsTot.Range("A1").Value = "期間: " & Format(sd, "yyyy/mm/dd") & " ～ " & Format(ed, "yyyy/mm/dd") & "  [" & siteFilter & "]"
    With wsTot.Range("A1")
        .Font.Size = 12: .Font.Bold = True: .Font.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter: .RowHeight = 30: .Interior.Color = RGB(240, 245, 252)
    End With

    wsTot.Cells(2, 1).Value = "品名": wsTot.Cells(2, 2).Value = "単価"
    wsTot.Cells(2, 3).Value = "数量": wsTot.Cells(2, 4).Value = "小計"
    wsTot.Cells(2, 5).Value = "トータル売上"

    Dim dicSub As Object, dicQty As Object
    Set dicSub = CreateObject("Scripting.Dictionary")
    Set dicQty = CreateObject("Scripting.Dictionary")
    Dim oT As Double: oT = 0
    For i = 1 To fc
        k = CStr(fd(i, 3)) & "|" & CStr(fd(i, 4))
        If dicSub.Exists(k) Then
            dicSub(k) = dicSub(k) + CDbl(fd(i, 6)): dicQty(k) = dicQty(k) + CLng(fd(i, 5))
        Else
            dicSub.Add k, CDbl(fd(i, 6)): dicQty.Add k, CLng(fd(i, 5))
        End If
        oT = oT + CDbl(fd(i, 6))
    Next i

    Dim ni As Long: ni = dicSub.Count
    Dim ak As Variant, av As Variant
    ak = dicSub.Keys: av = dicSub.Items
    Dim ot2() As Variant
    ReDim ot2(1 To ni, 1 To 5)
    Dim pts() As String
    For i = 0 To ni - 1
        pts = Split(ak(i), "|")
        ot2(i + 1, 1) = pts(0): ot2(i + 1, 2) = CLng(pts(1))
        ot2(i + 1, 3) = dicQty(ak(i)): ot2(i + 1, 4) = av(i): ot2(i + 1, 5) = oT
    Next i
    wsTot.Range("A3").Resize(ni, 5).Value = ot2

    With wsTot.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsTot.Range("A3:A" & ni + 2), Order:=xlAscending, CustomOrder:="和花,切花"
        .SortFields.Add Key:=wsTot.Range("B3:B" & ni + 2), Order:=xlAscending
        .SetRange wsTot.Range("A2:E" & ni + 2): .Header = xlYes: .Apply
    End With

    Call StyleSheetAt(wsTot, 2, ni + 2, 5)
    wsTot.Range("B:E").NumberFormat = "#,##0"
    Call MergeCol(wsTot, 1, 3, ni + 2)
    Call MergeCol(wsTot, 5, 3, ni + 2)
    Call ApplyTotalStripe(wsTot, 3, ni + 2)

    ' ===== グラフシート ★ScreenUpdating=True後に作成 =====
    Application.ScreenUpdating = True
    Application.StatusBar = "グラフ作成中..."
    Call MakeChart(fd, fc, sd, ed, siteFilter)

    wsS.Range("D14").Value = "期間: " & Format(sd, "yyyy/mm/dd") & " - " & Format(ed, "yyyy/mm/dd")
    wsS.Range("D15").Value = "対象: " & fc & " 件 / 品目: " & ni & " 種類"
    wsS.Range("D16").Value = "総売上: " & Format(oT, "#,##0") & " 円"
    wsS.Range("D14:D16").Font.Size = 11
    wsS.Range("D16").Font.Bold = True: wsS.Range("D16").Font.Color = RGB(31, 78, 120)

    wsS.Activate
    MsgBox "完了! " & fc & "件取得 (" & siteFilter & ")", vbInformation, "レポート生成完了"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub

Sub StyleSheet(ws As Worksheet, lr As Long, lc As Long)
    Call StyleSheetAt(ws, 1, lr, lc)
End Sub

Sub StyleSheetAt(ws As Worksheet, hdrRow As Long, lr As Long, lc As Long)
    With ws.Range(ws.Cells(hdrRow, 1), ws.Cells(hdrRow, lc))
        .Interior.Color = RGB(31, 78, 120): .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True: .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 28
    End With
    Dim ri As Long
    For ri = hdrRow + 1 To lr
        With ws.Range(ws.Cells(ri, 1), ws.Cells(ri, lc))
            If .Cells(1).Interior.Color <> RGB(218, 230, 242) Then
                .Interior.Color = RGB(255, 255, 255)
            End If
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
    Next ri
    With ws.Range(ws.Cells(hdrRow, 1), ws.Cells(lr, lc)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(180, 198, 220)
    End With
    Dim c As Long
    For c = 1 To lc
        ws.Columns(c).AutoFit
        If ws.Columns(c).ColumnWidth < 13 Then ws.Columns(c).ColumnWidth = 13
    Next c
    ws.Cells(hdrRow + 1, 1).Select
    ActiveWindow.FreezePanes = True
End Sub

Sub MergeCol(ws As Worksheet, col As Long, startRow As Long, lastRow As Long)
    Dim r As Long, sm As Long
    sm = startRow
    Application.DisplayAlerts = False
    For r = startRow + 1 To lastRow + 1
        Dim isSep As Boolean: isSep = False
        If r <= lastRow Then
            isSep = (ws.Cells(r, 3).Value = "日合計") Or (ws.Cells(r, 4).Value = "" And ws.Cells(r, 3).Value = "")
        End If
        If ws.Cells(r, col).Value <> ws.Cells(r - 1, col).Value Or r > lastRow Or isSep Then
            If r - 1 > sm And ws.Cells(sm, col).Value <> "" Then
                ws.Range(ws.Cells(sm, col), ws.Cells(r - 1, col)).Merge
                ws.Cells(sm, col).HorizontalAlignment = xlCenter: ws.Cells(sm, col).VerticalAlignment = xlCenter
            End If
            sm = r
            If isSep Then sm = r + 1
        End If
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
        If ws.Cells(r, 2).MergeCells Then
            curStore = CStr(ws.Cells(r, 2).MergeArea.Cells(1, 1).Value)
        Else
            curStore = CStr(ws.Cells(r, 2).Value)
        End If
        If curStore <> prevStore And curStore <> "" Then storeIdx = storeIdx + 1: prevStore = curStore
        If storeIdx Mod 2 = 1 Then
            For c = 1 To 7
                If ws.Cells(r, c).MergeCells Then ws.Cells(r, c).MergeArea.Interior.Color = fillClr _
                Else ws.Cells(r, c).Interior.Color = fillClr
            Next c
        End If
NextDR:
    Next r
End Sub

Sub ApplyTotalStripe(ws As Worksheet, startRow As Long, lastRow As Long)
    Dim fillClr As Long: fillClr = RGB(235, 242, 250)
    Dim r As Long, dataIdx As Long, c As Long
    dataIdx = 0
    For r = startRow To lastRow
        If ws.Cells(r, 2).Value = "" Then GoTo NextTR
        dataIdx = dataIdx + 1
        If dataIdx Mod 2 = 1 Then
            For c = 2 To 4: ws.Cells(r, c).Interior.Color = fillClr: Next c
        End If
NextTR:
    Next r
End Sub

Sub MakeChart(fd As Variant, fc As Long, sd As Date, ed As Date, siteFilter As String)
    Dim dicDate As Object, dicTemp As Object, dicTempCnt As Object
    Set dicDate = CreateObject("Scripting.Dictionary")
    Set dicTemp = CreateObject("Scripting.Dictionary")
    Set dicTempCnt = CreateObject("Scripting.Dictionary")
    Dim i As Long, dk2 As String
    For i = 1 To fc
        dk2 = Format(fd(i, 1), "yyyy/mm/dd")
        If dicDate.Exists(dk2) Then
            dicDate(dk2) = dicDate(dk2) + CDbl(fd(i, 6))
        Else
            dicDate.Add dk2, CDbl(fd(i, 6))
            dicTemp.Add dk2, 0: dicTempCnt.Add dk2, 0
        End If
        If fd(i, 7) <> "" And IsNumeric(fd(i, 7)) Then
            dicTemp(dk2) = dicTemp(dk2) + CDbl(fd(i, 7))
            dicTempCnt(dk2) = dicTempCnt(dk2) + 1
        End If
    Next i
    If dicDate.Count = 0 Then Exit Sub

    Dim wsG As Worksheet
    Set wsG = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsG.Name = "グラフ"
    wsG.Tab.Color = RGB(31, 78, 120)
    wsG.Columns("A").ColumnWidth = 3

    wsG.Range("B2").Value = "売上・気温グラフ"
    With wsG.Range("B2")
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(31, 78, 120)
    End With
    wsG.Range("B4").Value = "集計軸:"
    wsG.Range("B4").Font.Bold = True: wsG.Range("B4").Font.Size = 11

    Dim dicStores As Object, dicItems As Object
    Set dicStores = CreateObject("Scripting.Dictionary")
    Set dicItems = CreateObject("Scripting.Dictionary")
    For i = 1 To fc
        If Not dicStores.Exists(CStr(fd(i, 2))) Then dicStores.Add CStr(fd(i, 2)), 1
        If Not dicItems.Exists(CStr(fd(i, 3))) Then dicItems.Add CStr(fd(i, 3)), 1
    Next i
    Dim fullList As String, kv As Variant
    fullList = "日付別(全体)"
    For Each kv In dicStores.Keys: fullList = fullList & "," & CStr(kv): Next
    For Each kv In dicItems.Keys: fullList = fullList & "," & CStr(kv): Next

    With wsG.Range("C4")
        .Value = "日付別(全体)": .Font.Size = 11: .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(31, 78, 120): .Borders.Weight = xlThin
    End With
    wsG.Columns("C").ColumnWidth = 18
    With wsG.Range("C4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=fullList
        .ShowError = True
    End With

    Dim btn As Object
    Set btn = wsG.Shapes.AddShape(5, wsG.Range("D4").Left + 5, wsG.Range("D4").Top + 3, 70, 28)
    btn.Name = "btnUpdateChart"
    btn.TextFrame.Characters().Text = "反映"
    btn.TextFrame.Characters().Font.Size = 12: btn.TextFrame.Characters().Font.Bold = True
    btn.TextFrame.Characters().Font.ColorIndex = 2
    btn.Fill.ForeColor.RGB = RGB(31, 78, 120): btn.Line.Visible = False
    btn.TextFrame.HorizontalAlignment = 2: btn.TextFrame.VerticalAlignment = 3
    btn.OnAction = "UpdateChart"

    ' データをP列(16)に書き込み（表示状態のまま）
    Dim col As Long: col = 16
    wsG.Cells(1, col).Value = "日付"
    wsG.Cells(1, col + 1).Value = "売上金額"
    wsG.Cells(1, col + 2).Value = "最高気温"

    Dim dkeys As Variant: dkeys = dicDate.Keys
    Dim ii As Long, jj As Long, tmpV As Variant
    For ii = 0 To UBound(dkeys) - 1
        For jj = 0 To UBound(dkeys) - 1 - ii
            If CDate(dkeys(jj)) > CDate(dkeys(jj + 1)) Then
                tmpV = dkeys(jj): dkeys(jj) = dkeys(jj + 1): dkeys(jj + 1) = tmpV
            End If
        Next jj
    Next ii

    Dim nC As Long: nC = 0
    For i = 0 To UBound(dkeys)
        nC = nC + 1
        wsG.Cells(nC + 1, col).Value = dkeys(i)
        wsG.Cells(nC + 1, col + 1).Value = dicDate(dkeys(i))
        If dicTempCnt(dkeys(i)) > 0 Then
            wsG.Cells(nC + 1, col + 2).Value = Round(dicTemp(dkeys(i)) / dicTempCnt(dkeys(i)), 1)
        End If
    Next i

    ' ★ChartType を先に設定してからSetSourceData
    Dim chtObj As ChartObject
    Set chtObj = wsG.ChartObjects.Add(wsG.Range("B6").Left, wsG.Range("B6").Top, 700, 380)
    chtObj.Name = "SalesChart"
    Dim cht As Chart
    Set cht = chtObj.Chart
    cht.ChartType = xlColumnClustered

    ' データ列が表示中なのでSetSourceData可能
    cht.SetSourceData Source:=wsG.Range(wsG.Cells(1, col), wsG.Cells(nC + 1, col + 2))

    ' シリーズ設定
    With cht.SeriesCollection(1)
        .Name = "売上金額"
        .Format.Fill.ForeColor.RGB = RGB(31, 78, 120)
    End With
    If cht.SeriesCollection.Count >= 2 Then
        With cht.SeriesCollection(2)
            .Name = "最高気温"
            .ChartType = xlLine
            .AxisGroup = xlSecondary
            .Format.Line.ForeColor.RGB = RGB(255, 102, 0)
            .Format.Line.Weight = 2.5
        End With
    End If

    cht.HasTitle = True
    cht.ChartTitle.Text = "売上金額・気温 (" & Format(sd, "mm/dd") & "～" & Format(ed, "mm/dd") & ") " & siteFilter
    cht.ChartTitle.Font.Size = 13: cht.ChartTitle.Font.Bold = True
    cht.Axes(xlValue, xlPrimary).HasTitle = True
    cht.Axes(xlValue, xlPrimary).AxisTitle.Text = "売上金額 (円)"
    cht.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0"
    If cht.SeriesCollection.Count >= 2 Then
        cht.Axes(xlValue, xlSecondary).HasTitle = True
        cht.Axes(xlValue, xlSecondary).AxisTitle.Text = "気温 (℃)"
    End If
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom
    cht.PlotArea.Interior.Color = RGB(245, 248, 252)
    cht.ChartArea.Border.LineStyle = xlContinuous
    cht.ChartArea.Border.Color = RGB(180, 198, 220)

    ' ★グラフ設定完了後に列を非表示
    wsG.Columns(col).Hidden = True
    wsG.Columns(col + 1).Hidden = True
    wsG.Columns(col + 2).Hidden = True
    wsG.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Sub UpdateChart()
    Dim wsG As Worksheet
    On Error GoTo ErrExit
    Set wsG = ThisWorkbook.Sheets("グラフ")
    On Error GoTo 0

    Dim axis As String
    axis = Trim(wsG.Range("C4").Value)
    If axis = "" Then axis = "日付別(全体)"

    Dim wsDay As Worksheet
    On Error Resume Next
    Set wsDay = ThisWorkbook.Sheets("日別成績")
    On Error GoTo 0
    If wsDay Is Nothing Then MsgBox "先にデータ取得を実行してください。": Exit Sub

    Dim lastR As Long
    lastR = wsDay.Cells(wsDay.Rows.Count, 4).End(xlUp).Row
    If lastR < 2 Then Exit Sub

    Dim chtObj As ChartObject
    On Error Resume Next
    Set chtObj = wsG.ChartObjects("SalesChart")
    On Error GoTo 0
    If chtObj Is Nothing Then Exit Sub

    Dim cht As Chart
    Set cht = chtObj.Chart
    Dim dicAgg As Object
    Set dicAgg = CreateObject("Scripting.Dictionary")
    Dim i As Long, grpKey As String
    Dim curDate As String, curStore As String, cellItem As String, cellSub As Double
    curDate = "": curStore = ""

    For i = 2 To lastR
        If wsDay.Cells(i, 3).Value = "日合計" Then GoTo SkipRow
        If wsDay.Cells(i, 1).Value <> "" Then curDate = CStr(wsDay.Cells(i, 1).Value)
        If wsDay.Cells(i, 2).MergeCells Then
            If wsDay.Cells(i, 2).MergeArea.Cells(1, 1).Value <> "" Then curStore = CStr(wsDay.Cells(i, 2).MergeArea.Cells(1, 1).Value)
        Else
            If wsDay.Cells(i, 2).Value <> "" Then curStore = CStr(wsDay.Cells(i, 2).Value)
        End If
        cellItem = CStr(wsDay.Cells(i, 3).Value)
        If wsDay.Cells(i, 6).Value = "" Then GoTo SkipRow
        cellSub = CDbl(wsDay.Cells(i, 6).Value)
        grpKey = ""
        If axis = "日付別(全体)" Then
            grpKey = curDate
        ElseIf StrComp(axis, curStore, vbTextCompare) = 0 Then
            grpKey = curDate
        ElseIf StrComp(axis, cellItem, vbTextCompare) = 0 Then
            grpKey = curDate
        End If
        If grpKey = "" Then GoTo SkipRow
        If dicAgg.Exists(grpKey) Then dicAgg(grpKey) = dicAgg(grpKey) + cellSub Else dicAgg.Add grpKey, cellSub
SkipRow:
    Next i

    If dicAgg.Count = 0 Then MsgBox "データがありません。": Exit Sub

    Dim col As Long: col = 16
    wsG.Columns(col).Hidden = False
    wsG.Columns(col + 1).Hidden = False
    wsG.Range(wsG.Cells(2, col), wsG.Cells(1000, col + 1)).ClearContents
    wsG.Cells(1, col).Value = "日付": wsG.Cells(1, col + 1).Value = "売上金額"

    Dim kv As Variant, rowIdx As Long: rowIdx = 2
    For Each kv In dicAgg.Keys
        wsG.Cells(rowIdx, col).Value = CStr(kv)
        wsG.Cells(rowIdx, col + 1).Value = dicAgg(kv)
        rowIdx = rowIdx + 1
    Next kv
    Dim nC As Long: nC = rowIdx - 2

    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    cht.ChartType = xlColumnClustered
    cht.SetSourceData Source:=wsG.Range(wsG.Cells(1, col), wsG.Cells(nC + 1, col + 1))
    cht.SeriesCollection(1).Name = "売上金額"
    cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(31, 78, 120)
    cht.ChartTitle.Text = "売上金額 (" & axis & ")"

    wsG.Columns(col).Hidden = True
    wsG.Columns(col + 1).Hidden = True
    MsgBox "グラフ更新完了 (" & axis & ")", vbInformation
    Exit Sub
ErrExit:
    MsgBox "エラー: グラフシートが見つかりません。先にデータ取得を実行してください。"
End Sub
"""


def setup():
    import win32com.client
    import subprocess

    subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True, text=True)
    time.sleep(2)

    if os.path.exists(REPORT_FILE):
        for attempt in range(5):
            try:
                os.remove(REPORT_FILE)
                break
            except PermissionError:
                time.sleep(1)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Add()
        while wb.Sheets.Count > 1:
            wb.Sheets(wb.Sheets.Count).Delete()

        ws = wb.Sheets(1)
        ws.Name = "検索"

        ws.Range("A1:H25").Interior.Color = 0xFFF8F0

        ws.Range("C2:F2").Merge()
        c = ws.Range("C2")
        c.Value = "売上レポートシステム"
        c.Font.Size = 20
        c.Font.Bold = True
        c.Font.Color = 0x784E1F
        c.HorizontalAlignment = -4108

        ws.Range("C3:F3").Merge()
        c = ws.Range("C3")
        c.Value = "期間を指定してデータを取得"
        c.Font.Size = 10
        c.Font.Color = 0x808080
        c.HorizontalAlignment = -4108

        ws.Range("C5:F5").Borders(9).LineStyle = 1
        ws.Range("C5:F5").Borders(9).Color = 0xC8A878
        ws.Range("C5:F5").Borders(9).Weight = 2

        ws.Range("C7").Value = "抽出条件"
        ws.Range("C7").Font.Size = 13
        ws.Range("C7").Font.Bold = True
        ws.Range("C7").Font.Color = 0x784E1F

        cl = ws.Range("C8")
        cl.Value = "取得元"
        cl.Interior.Color = 0x784E1F
        cl.Font.Color = 0xFFFFFF
        cl.Font.Bold = True
        cl.Font.Size = 11
        cl.HorizontalAlignment = -4108
        for b in [7, 8, 9, 10]:
            cl.Borders(b).LineStyle = 1
            cl.Borders(b).Color = 0x784E1F

        cv = ws.Range("D8")
        cv.Value = "すべて"
        cv.Font.Size = 12
        cv.HorizontalAlignment = -4108
        cv.Interior.Color = 0xFFFFFF
        for b in [7, 8, 9, 10]:
            cv.Borders(b).LineStyle = 1
            cv.Borders(b).Color = 0x784E1F

        val = cv.Validation
        val.Delete()
        val.Add(3, 1, 1, "すべて,きむら,JA")
        val.ShowError = True

        for label, cell_l, cell_v, default in [
            ("開始日", "C9", "D9", "2026/01/01"),
            ("終了日", "C10", "D10", "2026/12/31"),
        ]:
            cl = ws.Range(cell_l)
            cl.Value = label
            cl.Interior.Color = 0x784E1F
            cl.Font.Color = 0xFFFFFF
            cl.Font.Bold = True
            cl.Font.Size = 11
            cl.HorizontalAlignment = -4108
            for b in [7, 8, 9, 10]:
                cl.Borders(b).LineStyle = 1
                cl.Borders(b).Color = 0x784E1F

            cv = ws.Range(cell_v)
            cv.Value = default
            cv.NumberFormat = "yyyy/mm/dd"
            cv.Font.Size = 12
            cv.HorizontalAlignment = -4108
            cv.Interior.Color = 0xFFFFFF
            for b in [7, 8, 9, 10]:
                cv.Borders(b).LineStyle = 1
                cv.Borders(b).Color = 0x784E1F

        left = ws.Range("C12").Left
        top = ws.Range("C12").Top + 5
        width = ws.Range("D12").Left + ws.Range("D12").Width - ws.Range("C12").Left
        btn = ws.Shapes.AddShape(5, left, top, width, 35)
        btn.Name = "btnFetch"
        btn.TextFrame.Characters().Text = "データ取得"
        btn.TextFrame.Characters().Font.Size = 13
        btn.TextFrame.Characters().Font.Bold = True
        btn.TextFrame.Characters().Font.ColorIndex = 2
        btn.Fill.ForeColor.RGB = 0xB6752E
        btn.Line.Visible = False
        btn.TextFrame.HorizontalAlignment = 2
        btn.TextFrame.VerticalAlignment = 3
        btn.Shadow.Visible = True
        btn.Shadow.OffsetX = 2
        btn.Shadow.OffsetY = 2
        btn.Shadow.Transparency = 0.7

        ws.Range("C13").Value = "実行結果"
        ws.Range("C13").Font.Size = 13
        ws.Range("C13").Font.Bold = True
        ws.Range("C13").Font.Color = 0x784E1F
        ws.Range("D13:D15").Font.Color = 0x555555

        ws.Columns("C").ColumnWidth = 12
        ws.Columns("D").ColumnWidth = 25
        ws.Columns("F").ColumnWidth = 15

        excel.ActiveWindow.DisplayGridlines = False

        vba_ok = False
        try:
            mod = wb.VBProject.VBComponents.Add(1)
            mod.Name = "SalesReport"
            mod.CodeModule.AddFromString(VBA_CODE)
            btn.OnAction = "FetchSalesData"
            vba_ok = True
        except Exception as e:
            print(f"VBA書込エラー: {e}")

        temp_path = os.path.join(tempfile.gettempdir(), "売上確認_temp.xlsm")
        if os.path.exists(temp_path):
            os.remove(temp_path)
        wb.SaveAs(temp_path, 52)
        wb.Close(False)

        shutil.copy2(temp_path, REPORT_FILE)
        os.remove(temp_path)
        print("COMPLETE!" + (" VBA正常書込。" if vba_ok else " VBA手動インポート要。"))

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            excel.Quit()
        except:
            pass


if __name__ == "__main__":
    setup()