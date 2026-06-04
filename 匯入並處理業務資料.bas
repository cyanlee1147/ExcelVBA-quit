Attribute VB_Name = "Module2"
Option Explicit

'==============================================================================
' 模組：業務員自動清理工具
' 功能：匯入原始保單匯出檔，自動完成篩選、刪欄、標色、結構調整與去重複
' 作者：Cyan Lee
' 更新：2026-06
'------------------------------------------------------------------------------
' 設計重點：
'   - 以「欄位標題名稱」動態定位欄位，不寫死欄號，適應不同來源檔案
'   - 由下往上反向迴圈刪列，避免索引位移
'   - 集中錯誤處理，發生例外時仍會還原 Excel 設定並關閉停用狀態
'==============================================================================

Sub 匯入並處理新檔案_獨立開啟()

    Dim 匯入檔案 As Variant
    Dim wbNew As Workbook
    Dim ws As Worksheet

    ' --- 1. 選擇要處理的檔案（僅限 .xlsx）---
    匯入檔案 = Application.GetOpenFilename( _
        "Excel 活頁簿 (*.xlsx), *.xlsx", , "請選擇要處理的 Excel 檔案")
    If 匯入檔案 = False Then Exit Sub        ' 使用者按取消

    ' --- 2. 提升效能並開始錯誤監控 ---
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wbNew = Workbooks.Open(匯入檔案)
    Set ws = wbNew.Sheets(1)

    ' 防呆：空白工作表直接結束
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        MsgBox "選取的檔案沒有任何資料。", vbExclamation
        GoTo CleanExit
    End If

    處理工作表 ws

    MsgBox "完成匯入、清理、標色與去除重複手機！", vbInformation, "處理完成"

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    ' 任何例外都會走到這裡，確保設定被還原、不會讓 Excel 卡在停用狀態
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "處理過程發生錯誤：" & vbCrLf & _
           "錯誤代碼 " & Err.Number & "：" & Err.Description, _
           vbCritical, "錯誤"
End Sub


'------------------------------------------------------------------------------
' 主處理流程（拆成獨立程序，邏輯較清楚、也方便日後維護）
'------------------------------------------------------------------------------
Private Sub 處理工作表(ws As Worksheet)

    Dim lastRow As Long
    Dim colIndex As Long
    Dim i As Long
    Dim 保單欄 As Long, col經手人 As Long, col經手人ID As Long, col手機 As Long

    ' ===== 步驟 1：刪除「保單號碼」開頭為 XX 的列 =====
    保單欄 = 尋找欄位(ws, "保單號", True)        ' True = 模糊比對（含「保單號碼」等變體）
    If 保單欄 > 0 Then
        Dim cellValue As String
        lastRow = ws.Cells(ws.Rows.Count, 保單欄).End(xlUp).Row
        For i = lastRow To 2 Step -1
            cellValue = Trim(ws.Cells(i, 保單欄).Text)
            ' 清掉可能混入的隱藏字元與全/半形單引號
            cellValue = Replace(cellValue, "'", "")
            cellValue = Replace(cellValue, Chr(8217), "")   ' 全形右單引號
            cellValue = Replace(cellValue, Chr(160), "")    ' 不換行空白
            If Left$(cellValue, 2) = "XX" Then ws.Rows(i).Delete
        Next i
    End If

    ' ===== 步驟 2：刪除「修改日」、「修改人」欄 =====
    For colIndex = ws.UsedRange.Columns.Count To 1 Step -1
        Select Case Trim(ws.Cells(1, colIndex).Value)
            Case "修改日", "修改人"
                ws.Columns(colIndex).Delete
        End Select
    Next colIndex

    ' ===== 步驟 3：定位欄位並把業務相關欄標黃 =====
    For colIndex = 1 To ws.UsedRange.Columns.Count
        Select Case Trim(ws.Cells(1, colIndex).Value)
            Case "業務員", "經手人", "經手人ID"
                ws.Columns(colIndex).Interior.Color = RGB(255, 255, 0)  ' 標黃
        End Select
    Next colIndex

    ' 個別記錄關鍵欄位位置（修正原版 Select Case 永遠抓不到的 bug）
    col經手人 = 尋找欄位(ws, "經手人", False)
    col經手人ID = 尋找欄位(ws, "經手人ID", False)
    col手機 = 尋找欄位(ws, "要保人手機", False)

    ' ===== 步驟 4：在「經手人」與「經手人ID」之間插入空欄 =====
    If col經手人 > 0 And col經手人ID > 0 Then
        ws.Columns(Application.WorksheetFunction.Max(col經手人, col經手人ID)) _
            .Insert Shift:=xlToRight
        ' 插入後欄位位置會位移，重新定位手機欄以免後續抓錯
        col手機 = 尋找欄位(ws, "要保人手機", False)
    End If

    ' ===== 步驟 5：將「要保人手機」設為紅字 =====
    If col手機 > 0 Then ws.Columns(col手機).Font.Color = RGB(255, 0, 0)

    ' ===== 步驟 6：在最前方新增「承接者」欄並塗黃 =====
    ws.Columns(1).Insert Shift:=xlToRight
    ws.Cells(1, 1).Value = "承接者"
    ws.Columns(1).Interior.Color = RGB(255, 255, 0)
    If col手機 > 0 Then col手機 = col手機 + 1     ' 整體右移一欄

    ' ===== 步驟 7：依「要保人手機」去除重複列 =====
    If col手機 > 0 Then
        Dim 手機Dict As Object
        Set 手機Dict = CreateObject("Scripting.Dictionary")
        Dim phoneVal As String
        lastRow = ws.Cells(ws.Rows.Count, col手機).End(xlUp).Row
        For i = lastRow To 2 Step -1
            phoneVal = Trim(ws.Cells(i, col手機).Text)
            If Len(phoneVal) > 0 Then
                If 手機Dict.Exists(phoneVal) Then
                    ws.Rows(i).Delete
                Else
                    手機Dict.Add phoneVal, True
                End If
            End If
        Next i
    End If

End Sub


'------------------------------------------------------------------------------
' 輔助函式：依標題列文字找出欄號，找不到回傳 0
'   fuzzy = True  → 模糊比對（標題「包含」關鍵字）
'   fuzzy = False → 完全相符
'------------------------------------------------------------------------------
Private Function 尋找欄位(ws As Worksheet, 關鍵字 As String, _
                          fuzzy As Boolean) As Long
    Dim c As Long, 標題 As String
    尋找欄位 = 0
    For c = 1 To ws.UsedRange.Columns.Count
        標題 = Trim(ws.Cells(1, c).Value)
        If fuzzy Then
            If 標題 Like "*" & 關鍵字 & "*" Then 尋找欄位 = c: Exit Function
        Else
            If 標題 = 關鍵字 Then 尋找欄位 = c: Exit Function
        End If
    Next c
End Function
