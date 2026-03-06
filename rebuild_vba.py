"""
rebuild_vba.py
SalesReport モジュールを Shift-JIS(cp932) で保存してExcelにインポートする。
Python文字列 UTF-8 → cp932 変換でVBAの文字化け問題を完全回避。
"""
import os, time, subprocess, tempfile
import win32com.client

BASE_DIR = r"C:\Users\sawak\OneDrive\デスクトップ\売上メール"
XLS_FILE = os.path.join(BASE_DIR, "売上確認.xlsm")

# =============================================================================
# VBAコード（Pythonの文字列としてUTF-8で保持、書き出し時にcp932変換）
# =============================================================================
VBA = r"""Attribute VB_Name = "SalesReport"
Option Explicit

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
    Dim sep As String: sep = Application.PathSeparator
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = ThisWorkbook.Path

    If InStr(1, basePath, "https", vbTextCompare) > 0 Then
        basePath = Environ("USERPROFILE")
        If fso.FolderExists(basePath & sep & "OneDrive") Then
            basePath = basePath & sep & "OneDrive"
        End If
        If fso.FolderExists(basePath & sep & "Desktop") Then
            basePath = basePath & sep & "Desktop"
        End If
        If fso.FolderExists(basePath & sep & "デスクトップ") Then
            basePath = basePath & sep & "デスクトップ"
        End If
        If fso.FolderExists(basePath & sep & "売上メール") Then
            basePath = basePath & sep & "売上メール"
        End If
    End If



    dp = basePath & sep & "データベース" & sep & "売上管理表.xlsx"
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
    Dim k As String, dk2 As String
    For i = 1 To fc
        k = Format(fd(i, 1), "yyyymmdd") & "|" & CStr(fd(i, 2))
        If dicDT.Exists(k) Then dicDT(k) = dicDT(k) + CDbl(fd(i, 6)) Else dicDT.Add k, CDbl(fd(i, 6))
        dk2 = Format(fd(i, 1), "yyyymmdd")
        If dicDayAll.Exists(dk2) Then dicDayAll(dk2) = dicDayAll(dk2) + CDbl(fd(i, 6)) Else dicDayAll.Add dk2, CDbl(fd(i, 6))
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
            If wsDay.Cells(sr, 1).Value <> "" Then dayKey = Format(wsDay.Cells(sr, 1).Value, "yyyymmdd"): Exit For
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
    Call ClearDuplicateCol(wsDay, 1, 2, newLastR) ' 日付
    Call ClearDuplicateCol(wsDay, 2, 2, newLastR, 1) ' 店舗名 (親: 日付)
    Call ClearDuplicateCol(wsDay, 3, 2, newLastR, 2, 1) ' 品名 (親: 店舗名, 日付)
    Call ClearDuplicateCol(wsDay, 7, 2, newLastR, 2, 1) ' 日別売上合計 (親: 店舗名, 日付)
    Call ApplyDayStripe(wsDay, 2, newLastR)

    ' ★ 印刷設定（1日=1ページ）
    Call SetupDayPrint(wsDay, newLastR)

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
    Call ClearDuplicateCol(wsTot, 1, 3, ni + 2) ' 品名
    Call ClearDuplicateCol(wsTot, 5, 3, ni + 2) ' トータル売上
    Call ApplyTotalStripe(wsTot, 3, ni + 2)

    Application.ScreenUpdating = True
    Application.StatusBar = "グラフ作成中..."
    Call MakeChart(fd, fc, sd, ed, siteFilter)

    ' ===== 店別データシート =====
    Application.StatusBar = "店別データ作成中..."
    Call BuildStoreSheet(fd, fc)

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

' ★ 1日=1ページの印刷設定
Sub SetupDayPrint(ws As Worksheet, lastRow As Long)
    ws.ResetAllPageBreaks

    ' 日合計行の2行後（次の日の最初の行）に改ページを設定
    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, 3).Value = "日合計" Then
            Dim nextStart As Long
            nextStart = r + 2  ' 日合計→空白→次の日
            If nextStart <= lastRow Then
                ws.HPageBreaks.Add Before:=ws.Rows(nextStart)
            End If
        End If
    Next r

    ' ページ設定
    With ws.PageSetup
        .Orientation = xlLandscape        ' 横向き
        .FitToPagesWide = 1               ' 横は必ず1ページに収める
        .FitToPagesTall = False           ' 縦は自動（改ページ優先）
        .Zoom = False
        .PrintTitleRows = "$1:$1"         ' 1行目(ヘッダー)を毎ページ印刷
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
        .PrintGridlines = False
    End With
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
            If .Cells(1).Interior.Color <> RGB(218, 230, 242) Then .Interior.Color = RGB(255, 255, 255)
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
    ' 日付ごとに売上・最高気温・最低気温を集計
    Dim dicDate As Object, dicTmax As Object, dicTmin As Object, dicCnt As Object
    Set dicDate = CreateObject("Scripting.Dictionary")
    Set dicTmax = CreateObject("Scripting.Dictionary")
    Set dicTmin = CreateObject("Scripting.Dictionary")
    Set dicCnt = CreateObject("Scripting.Dictionary")
    Dim i As Long, dKey As String
    For i = 1 To fc
        dKey = Format(fd(i, 1), "yyyy/mm/dd")
        If Not dicDate.Exists(dKey) Then
            dicDate.Add dKey, 0: dicTmax.Add dKey, 0
            dicTmin.Add dKey, 0: dicCnt.Add dKey, 0
        End If
        dicDate(dKey) = dicDate(dKey) + CDbl(fd(i, 6))
        If fd(i, 7) <> "" And IsNumeric(fd(i, 7)) Then
            dicTmax(dKey) = dicTmax(dKey) + CDbl(fd(i, 7))
            dicCnt(dKey) = dicCnt(dKey) + 1
        End If
        If fd(i, 8) <> "" And IsNumeric(fd(i, 8)) Then
            dicTmin(dKey) = dicTmin(dKey) + CDbl(fd(i, 8))
        End If
    Next i
    If dicDate.Count = 0 Then Exit Sub

    Dim wsG As Worksheet
    Set wsG = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsG.Name = "グラフ"
    wsG.Tab.Color = RGB(31, 78, 120)
    wsG.Columns("A").ColumnWidth = 3
    wsG.Range("B2").Value = "売上・気温グラフ"
    With wsG.Range("B2"): .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(31, 78, 120): End With

    ' --- 集計軸ドロップダウン（C4）---
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

    ' --- 折れ線ドロップダウン（C5）---
    wsG.Range("B5").Value = "折れ線:"
    wsG.Range("B5").Font.Bold = True: wsG.Range("B5").Font.Size = 11
    With wsG.Range("C5")
        .Value = "最高気温": .Font.Size = 11: .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(31, 78, 120): .Borders.Weight = xlThin
    End With
    With wsG.Range("C5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="最高気温,最低気温,なし"
        .ShowError = True
    End With

    ' --- 反映ボタン（C4:C5の高さに合わせて配置）---
    Dim btn As Object
    Dim btnH As Double: btnH = wsG.Range("C4:C5").Height
    Dim btnL As Double: btnL = wsG.Range("D4").Left + 4
    Dim btnT As Double: btnT = wsG.Range("C4").Top + 2
    Set btn = wsG.Shapes.AddShape(5, btnL, btnT, 75, btnH - 4)
    btn.Name = "btnUpdateChart"
    btn.TextFrame.Characters().Text = "反映"
    btn.TextFrame.Characters().Font.Size = 12: btn.TextFrame.Characters().Font.Bold = True
    btn.TextFrame.Characters().Font.ColorIndex = 2
    btn.Fill.ForeColor.RGB = RGB(31, 78, 120): btn.Line.Visible = False
    btn.TextFrame.HorizontalAlignment = 2: btn.TextFrame.VerticalAlignment = 3
    btn.OnAction = "UpdateChart"

    ' --- 日付ソート ---
    Dim dkeys As Variant: dkeys = dicDate.Keys
    Dim ii As Long, jj As Long, tmpV As Variant
    For ii = 0 To UBound(dkeys) - 1
        For jj = 0 To UBound(dkeys) - 1 - ii
            If CDate(dkeys(jj)) > CDate(dkeys(jj + 1)) Then
                tmpV = dkeys(jj): dkeys(jj) = dkeys(jj + 1): dkeys(jj + 1) = tmpV
            End If
        Next jj
    Next ii
    Dim nC As Long: nC = UBound(dkeys) + 1

    ' --- VBA配列にデータを格納（SetSourceData不使用） ---
    Dim salesArr() As Double, tmaxArr() As Double, tminArr() As Double, labArr() As String
    ReDim salesArr(0 To nC - 1): ReDim tmaxArr(0 To nC - 1)
    ReDim tminArr(0 To nC - 1): ReDim labArr(0 To nC - 1)
    For i = 0 To nC - 1
        labArr(i) = CStr(dkeys(i))
        salesArr(i) = dicDate(dkeys(i))
        If dicCnt(dkeys(i)) > 0 Then
            tmaxArr(i) = Round(dicTmax(dkeys(i)) / dicCnt(dkeys(i)), 1)
            tminArr(i) = Round(dicTmin(dkeys(i)) / dicCnt(dkeys(i)), 1)
        End If
    Next i

    ' --- グラフ作成（NewSeries + 配列方式） ---
    Dim chtObj As ChartObject
    Set chtObj = wsG.ChartObjects.Add(wsG.Range("B7").Left, wsG.Range("B7").Top, 720, 380)
    chtObj.Name = "SalesChart"
    Dim cht As Chart
    Set cht = chtObj.Chart
    cht.ChartType = xlColumnClustered
    Do While cht.SeriesCollection.Count > 0: cht.SeriesCollection(1).Delete: Loop

    ' 系列1: 売上（棒）
    Dim srs1 As Series
    Set srs1 = cht.SeriesCollection.NewSeries
    srs1.ChartType = xlColumnClustered
    srs1.Name = "売上金額"
    srs1.Values = salesArr
    srs1.XValues = labArr
    srs1.Format.Fill.ForeColor.RGB = RGB(31, 78, 120)
    srs1.AxisGroup = xlPrimary

    ' 系列2: 最高気温（折れ線、既定）
    Dim srs2 As Series
    Set srs2 = cht.SeriesCollection.NewSeries
    srs2.ChartType = xlLine
    srs2.Name = "最高気温"
    srs2.Values = tmaxArr
    srs2.XValues = labArr
    srs2.Format.Line.ForeColor.RGB = RGB(255, 102, 0)
    srs2.Format.Line.Weight = 2.5
    srs2.AxisGroup = xlSecondary

    ' --- タイトル・軸設定 ---
    cht.HasTitle = True
    cht.ChartTitle.Text = "売上金額・気温 (" & Format(sd, "mm/dd") & "～" & Format(ed, "mm/dd") & ") " & siteFilter
    cht.ChartTitle.Font.Size = 13: cht.ChartTitle.Font.Bold = True
    cht.Axes(xlValue, xlPrimary).HasTitle = True
    cht.Axes(xlValue, xlPrimary).AxisTitle.Text = "売上金額 (円)"
    cht.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = "#,##0"
    cht.Axes(xlValue, xlSecondary).HasTitle = True
    cht.Axes(xlValue, xlSecondary).AxisTitle.Text = "気温 (℃)"
    cht.HasLegend = True: cht.Legend.Position = xlLegendPositionBottom
    cht.PlotArea.Interior.Color = RGB(245, 248, 252)
    cht.ChartArea.Border.LineStyle = xlContinuous
    cht.ChartArea.Border.Color = RGB(180, 198, 220)

    ' データをP列（非表示）に保存（UpdateChart用）
    Dim col As Long: col = 16
    wsG.Cells(1, col) = "日付": wsG.Cells(1, col+1) = "売上": wsG.Cells(1, col+2) = "最高気温": wsG.Cells(1, col+3) = "最低気温"
    For i = 0 To nC - 1
        wsG.Cells(i+2, col) = labArr(i): wsG.Cells(i+2, col+1) = salesArr(i)
        wsG.Cells(i+2, col+2) = tmaxArr(i): wsG.Cells(i+2, col+3) = tminArr(i)
    Next i
    wsG.Columns(col).Hidden = True: wsG.Columns(col+1).Hidden = True
    wsG.Columns(col+2).Hidden = True: wsG.Columns(col+3).Hidden = True
    wsG.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Sub UpdateChart()
    Dim wsG As Worksheet
    On Error GoTo ErrExit
    Set wsG = ThisWorkbook.Sheets("グラフ")
    On Error GoTo 0

    Dim axis As String, lineType As String
    axis = Trim(wsG.Range("C4").Value)
    lineType = Trim(wsG.Range("C5").Value)
    If axis = "" Then axis = "日付別(全体)"
    If lineType = "" Then lineType = "最高気温"

    Dim chtObj As ChartObject
    On Error Resume Next
    Set chtObj = wsG.ChartObjects("SalesChart")
    On Error GoTo 0
    If chtObj Is Nothing Then MsgBox "グラフが見つかりません。先にデータ取得を実行してください。": Exit Sub

    ' P列から保存データを読み込む
    Dim col As Long: col = 16
    wsG.Columns(col).Hidden = False: wsG.Columns(col+1).Hidden = False
    wsG.Columns(col+2).Hidden = False: wsG.Columns(col+3).Hidden = False
    Dim lastDataR As Long
    lastDataR = wsG.Cells(wsG.Rows.Count, col).End(xlUp).Row
    If lastDataR < 2 Then MsgBox "グラフデータがありません。先にデータ取得を実行してください。": GoTo HideCols

    ' 集計軸に応じてデータを集計
    Dim wsDay As Worksheet
    On Error Resume Next
    Set wsDay = ThisWorkbook.Sheets("日別成績")
    On Error GoTo 0

    Dim nC As Long: nC = lastDataR - 1
    Dim labArr() As String, salesArr() As Double, subArr() As Double
    ReDim labArr(0 To nC-1): ReDim salesArr(0 To nC-1): ReDim subArr(0 To nC-1)

    If axis = "日付別(全体)" Or wsDay Is Nothing Then
        ' P列のデータをそのまま使用
        Dim r2 As Long
        For r2 = 0 To nC - 1
            labArr(r2) = CStr(wsG.Cells(r2+2, col).Value)
            salesArr(r2) = CDbl(wsG.Cells(r2+2, col+1).Value)
        Next r2
        Select Case lineType
            Case "最高気温"
                For r2 = 0 To nC-1: subArr(r2) = CDbl(wsG.Cells(r2+2, col+2).Value): Next
            Case "最低気温"
                For r2 = 0 To nC-1: subArr(r2) = CDbl(wsG.Cells(r2+2, col+3).Value): Next
        End Select
    Else
        ' 日別成績シートから集計
        Dim dicAgg As Object, dicIdx As Object
        Set dicAgg = CreateObject("Scripting.Dictionary")
        Set dicIdx = CreateObject("Scripting.Dictionary")
        Dim lastR As Long: lastR = wsDay.Cells(wsDay.Rows.Count, 4).End(xlUp).Row
        Dim curDate As String, curStore As String, cellItem As String, grpKey As String
        curDate = "": curStore = ""
        Dim i As Long
        For i = 2 To lastR
            If wsDay.Cells(i, 3).Value = "日合計" Then GoTo SkipR
            If wsDay.Cells(i, 4).Value = "" And wsDay.Cells(i, 3).Value = "" Then GoTo SkipR
            If wsDay.Cells(i, 1).Value <> "" Then curDate = CStr(wsDay.Cells(i, 1).Value)
            If wsDay.Cells(i, 2).Value <> "" Then
                curStore = CStr(wsDay.Cells(i, 2).Value)
            End If
            cellItem = CStr(wsDay.Cells(i, 3).Value)
            If wsDay.Cells(i, 6).Value = "" Then GoTo SkipR
            grpKey = ""
            If StrComp(axis, curStore, vbTextCompare) = 0 Or StrComp(axis, cellItem, vbTextCompare) = 0 Then
                grpKey = curDate
            End If
            If grpKey = "" Then GoTo SkipR
            If dicAgg.Exists(grpKey) Then
                dicAgg(grpKey) = dicAgg(grpKey) + CDbl(wsDay.Cells(i, 6).Value)
            Else
                dicAgg.Add grpKey, CDbl(wsDay.Cells(i, 6).Value)
            End If
SkipR:
        Next i
        If dicAgg.Count = 0 Then MsgBox "条件に合うデータがありません。": GoTo HideCols
        nC = dicAgg.Count
        ReDim labArr(0 To nC-1): ReDim salesArr(0 To nC-1): ReDim subArr(0 To nC-1)
        ' P列のデータとマッチして気温も取得
        Dim p2Idx As Long: p2Idx = 0
        Dim kvKey As Variant
        For Each kvKey In dicAgg.Keys
            labArr(p2Idx) = CStr(kvKey)
            salesArr(p2Idx) = dicAgg(kvKey)
            ' P列から対応する気温を検索
            Dim rr As Long
            For rr = 2 To lastDataR
                If CStr(wsG.Cells(rr, col).Value) = CStr(kvKey) Then
                    Select Case lineType
                        Case "最高気温": subArr(p2Idx) = CDbl(wsG.Cells(rr, col+2).Value)
                        Case "最低気温": subArr(p2Idx) = CDbl(wsG.Cells(rr, col+3).Value)
                    End Select
                    Exit For
                End If
            Next rr
            p2Idx = p2Idx + 1
        Next kvKey
    End If

    ' グラフ更新（NewSeries + 配列方式）
    Dim cht As Chart: Set cht = chtObj.Chart
    Do While cht.SeriesCollection.Count > 0: cht.SeriesCollection(1).Delete: Loop

    Dim srs1 As Series
    Set srs1 = cht.SeriesCollection.NewSeries
    srs1.ChartType = xlColumnClustered
    srs1.Name = "売上金額"
    srs1.Values = salesArr
    srs1.XValues = labArr
    srs1.Format.Fill.ForeColor.RGB = RGB(31, 78, 120)
    srs1.AxisGroup = xlPrimary

    If lineType <> "なし" Then
        Dim srs2 As Series
        Set srs2 = cht.SeriesCollection.NewSeries
        srs2.ChartType = xlLine
        srs2.Name = lineType
        srs2.Values = subArr
        srs2.XValues = labArr
        If lineType = "最高気温" Then
            srs2.Format.Line.ForeColor.RGB = RGB(255, 102, 0)
        Else
            srs2.Format.Line.ForeColor.RGB = RGB(0, 130, 200)
        End If
        srs2.Format.Line.Weight = 2.5
        srs2.AxisGroup = xlSecondary
        cht.Axes(xlValue, xlSecondary).HasTitle = True
        cht.Axes(xlValue, xlSecondary).AxisTitle.Text = "気温 (℃)"
    End If

    cht.HasTitle = True
    cht.ChartTitle.Text = "売上金額・" & lineType & " (" & axis & ")"
    MsgBox "グラフ更新完了 (" & axis & " / " & lineType & ")", vbInformation

HideCols:
    wsG.Columns(col).Hidden = True: wsG.Columns(col+1).Hidden = True
    wsG.Columns(col+2).Hidden = True: wsG.Columns(col+3).Hidden = True
    Exit Sub
ErrExit:
    MsgBox "エラー: 先にデータ取得を実行してください。"
End Sub

' ===== 店別データシート（全店舗・商品×日付クロス表） =====
Sub BuildStoreSheet(fd As Variant, fc As Long)
    If fc = 0 Then Exit Sub

    ' --- 既存シートを削除 ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("店別データ").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' --- シート作成 ---
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "店別データ"
    ws.Tab.Color = RGB(0, 112, 192)

    ' --- 日付リスト（昇順）を収集 ---
    Dim dicDates As Object
    Set dicDates = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To fc
        Dim dk As String
        dk = Format(fd(i, 1), "yyyy/mm/dd")
        If Not dicDates.Exists(dk) Then dicDates.Add dk, CDate(fd(i, 1))
    Next i

    ' 日付を昇順ソート
    Dim dateKeys() As String
    ReDim dateKeys(0 To dicDates.Count - 1)
    Dim di As Long: di = 0
    Dim kv As Variant
    For Each kv In dicDates.Keys
        dateKeys(di) = CStr(kv): di = di + 1
    Next kv
    ' バブルソート
    Dim si As Long, sj As Long, tmp As String
    For si = 0 To UBound(dateKeys) - 1
        For sj = si + 1 To UBound(dateKeys)
            If dicDates(dateKeys(si)) > dicDates(dateKeys(sj)) Then
                tmp = dateKeys(si): dateKeys(si) = dateKeys(sj): dateKeys(sj) = tmp
            End If
        Next sj
    Next si
    Dim nDates As Long: nDates = UBound(dateKeys) + 1

    ' --- 店舗リスト（出現順）を収集 ---
    Dim dicStores As Object
    Set dicStores = CreateObject("Scripting.Dictionary")
    For i = 1 To fc
        Dim st As String: st = CStr(fd(i, 2))
        If st <> "" And Not dicStores.Exists(st) Then dicStores.Add st, dicStores.Count
    Next i

    ' --- 店舗×商品×日付 の数量を集計 ---
    ' キー: 店舗名|品名|日付 → 数量合計
    Dim dicQty As Object
    Set dicQty = CreateObject("Scripting.Dictionary")
    ' 店舗別の品名リスト（順序保持）
    Dim dicStoreItems As Object
    Set dicStoreItems = CreateObject("Scripting.Dictionary")

    For i = 1 To fc
        Dim storeName As String: storeName = CStr(fd(i, 2))
        Dim baseItem As String:  baseItem  = CStr(fd(i, 3))
        Dim priceStr As String:  priceStr  = CStr(fd(i, 4))
        
        Dim itemName As String
        If priceStr <> "" Then
            itemName = baseItem & priceStr
        Else
            itemName = baseItem
        End If

        Dim dateStr As String:   dateStr   = Format(fd(i, 1), "yyyy/mm/dd")
        Dim qty As Long:         qty       = CLng(fd(i, 5))

        Dim qKey As String: qKey = storeName & "|" & itemName & "|" & dateStr
        If dicQty.Exists(qKey) Then
            dicQty(qKey) = dicQty(qKey) + qty
        Else
            dicQty.Add qKey, qty
        End If

        ' 品名リスト
        If Not dicStoreItems.Exists(storeName) Then
            Set dicStoreItems(storeName) = CreateObject("Scripting.Dictionary")
        End If
        If Not dicStoreItems(storeName).Exists(itemName) Then
            dicStoreItems(storeName).Add itemName, dicStoreItems(storeName).Count
        End If
    Next i

    ' --- シートへ書き出し ---
    Dim curRow As Long: curRow = 1
    Dim storeKey As Variant

    ' 色定義
    Dim hdrBg As Long:  hdrBg  = RGB(31, 78, 120)   ' ヘッダ背景（紺）
    Dim hdrFg As Long:  hdrFg  = RGB(255, 255, 255)  ' ヘッダ文字（白）
    Dim stBg As Long:   stBg   = RGB(0, 112, 192)    ' 店舗名行背景（青）
    Dim rowA As Long:   rowA   = RGB(255, 255, 255)   ' 行色A
    Dim rowB As Long:   rowB   = RGB(235, 242, 250)   ' 行色B

    For Each storeKey In dicStores.Keys
        Dim sName As String: sName = CStr(storeKey)

        ' --- 店舗ヘッダ行（店舗名 | 日付1 | 日付2 | ...） ---
        With ws.Cells(curRow, 1)
            .Value = sName
            .Font.Bold = True
            .Font.Color = hdrFg
            .Interior.Color = stBg
        End With
        Dim ci As Long
        For ci = 0 To nDates - 1
            Dim hdCell As Range
            Set hdCell = ws.Cells(curRow, ci + 2)
            hdCell.Value = CDate(dicDates(dateKeys(ci)))
            hdCell.NumberFormat = "m/d"
            hdCell.Font.Bold = True
            hdCell.Font.Color = hdrFg
            hdCell.Interior.Color = hdrBg
            hdCell.HorizontalAlignment = xlCenter
        Next ci

        ' ヘッダ行の高さ
        ws.Rows(curRow).RowHeight = 22
        curRow = curRow + 1

        ' --- 商品のソート（和花(1)→切花(2)→その他五十音順(3)、かつ単価順） ---
        Dim nItems As Long
        nItems = dicStoreItems(sName).Count
        Dim arrItem() As String
        Dim arrSortKey() As String
        ReDim arrItem(0 To nItems - 1)
        ReDim arrSortKey(0 To nItems - 1)
        
        Dim iIdx As Long: iIdx = 0
        Dim itemKey As Variant
        For Each itemKey In dicStoreItems(sName).Keys
            Dim cName As String: cName = CStr(itemKey)
            arrItem(iIdx) = cName
            
            ' 数字部分（単価）と文字部分（純粋な品名）に分解
            Dim numStr As String, pIdx As Long, ch As String
            Dim baseName As String
            numStr = ""
            For pIdx = Len(cName) To 1 Step -1
                ch = Mid(cName, pIdx, 1)
                If IsNumeric(ch) Then
                    numStr = ch & numStr
                Else
                    Exit For
                End If
            Next pIdx
            
            baseName = Left(cName, Len(cName) - Len(numStr))
            If numStr = "" Then numStr = "0"
            
            ' カテゴリ判定
            Dim prio As String
            If Left(cName, 2) = "和花" Then
                prio = "1"
            ElseIf Left(cName, 2) = "切花" Then
                prio = "2"
            Else
                prio = "3"
            End If
            
            ' ソートキー: [優先度]_[文字部分]_[単価ゼロ埋め]
            ' 優先度1,2は元々和花/切花で同じなので後ろの単価で決まる。
            ' 優先度3は文字部分でソートされ、文字が同じ（鉢物など）なら単価順になる。
            Dim keyStr As String
            keyStr = prio & "_" & baseName & "_" & Right("000000" & numStr, 6)
            arrSortKey(iIdx) = keyStr
            iIdx = iIdx + 1
        Next itemKey
        
        ' バブルソート
        Dim k1 As Long, k2 As Long, tKey As String, tItm As String
        If nItems > 1 Then
            For k1 = 0 To nItems - 2
                For k2 = k1 + 1 To nItems - 1
                    If arrSortKey(k1) > arrSortKey(k2) Then
                        tKey = arrSortKey(k1): arrSortKey(k1) = arrSortKey(k2): arrSortKey(k2) = tKey
                        tItm = arrItem(k1): arrItem(k1) = arrItem(k2): arrItem(k2) = tItm
                    End If
                Next k2
            Next k1
        End If

        ' --- 商品行書き出し ---
        Dim itemIdx As Long: itemIdx = 0
        For iIdx = 0 To nItems - 1
            Dim iName As String: iName = arrItem(iIdx)
            Dim rowColor As Long
            If itemIdx Mod 2 = 0 Then rowColor = rowA Else rowColor = rowB

            With ws.Cells(curRow, 1)
                .Value = iName
                .Interior.Color = rowColor
            End With

            For ci = 0 To nDates - 1
                Dim qk As String: qk = sName & "|" & iName & "|" & dateKeys(ci)
                Dim cellVal As Long
                If dicQty.Exists(qk) Then cellVal = dicQty(qk) Else cellVal = 0

                With ws.Cells(curRow, ci + 2)
                    If cellVal > 0 Then
                        .Value = cellVal
                    Else
                        .Value = ""
                    End If
                    .Interior.Color = rowColor
                    .HorizontalAlignment = xlCenter
                End With
            Next ci

            ws.Rows(curRow).RowHeight = 18
            curRow = curRow + 1
            itemIdx = itemIdx + 1
        Next iIdx

        ' 店舗間の空行
        ws.Rows(curRow).RowHeight = 8
        ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, nDates + 1)).Interior.ColorIndex = xlNone
        curRow = curRow + 1
    Next storeKey

    ' --- 列幅自動調整 ---
    ws.Columns(1).AutoFit
    If ws.Columns(1).ColumnWidth < 14 Then ws.Columns(1).ColumnWidth = 14
    For ci = 2 To nDates + 1
        ws.Columns(ci).ColumnWidth = 7
    Next ci

    ' --- 枠線 ---
    Dim dataRng As Range
    Set dataRng = ws.Range(ws.Cells(1, 1), ws.Cells(curRow - 1, nDates + 1))
    With dataRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 198, 220)
    End With
    
    ' --- オートフィルタ設定 ---
    ws.Rows(1).AutoFilter
End Sub
"""


# =============================================================================
# cp932 で一時ファイルに書き出してExcelにインポート
# =============================================================================
def rebuild():
    subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True)
    time.sleep(2)

    # 一時 basファイルを cp932 で作成
    tmp_bas = os.path.join(tempfile.gettempdir(), "SalesReport_cp932.bas")
    with open(tmp_bas, "w", encoding="cp932", errors="replace") as f:
        f.write(VBA)
    print(f"BAS written (cp932): {tmp_bas}")

    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(XLS_FILE)
        proj = wb.VBProject
        # 既存モジュール削除
        for ci in range(proj.VBComponents.Count, 0, -1):
            comp = proj.VBComponents.Item(ci)
            if comp.Name == "SalesReport":
                proj.VBComponents.Remove(comp)
                print("Removed old SalesReport module")
                break
        # インポート
        proj.VBComponents.Import(tmp_bas)
        print("Imported SalesReport.bas (cp932)")
        # ボタンのOnAction設定
        ws = wb.Sheets("検索")
        for si in range(1, ws.Shapes.Count + 1):
            shp = ws.Shapes(si)
            if shp.Name == "btnFetch":
                shp.OnAction = "FetchSalesData"
                print("Button OnAction: FetchSalesData")
        wb.Save()
        wb.Close(False)
        print("COMPLETE! 売上確認.xlsm を更新しました。")
    except Exception as e:
        print(f"Error: {e}")
        import traceback; traceback.print_exc()
    finally:
        try: xl.Quit()
        except: pass
        if os.path.exists(tmp_bas):
            os.remove(tmp_bas)

if __name__ == "__main__":
    rebuild()
