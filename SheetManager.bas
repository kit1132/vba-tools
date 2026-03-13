Option Explicit

Private Const LIST_SHEET_NAME As String = "シート一覧"

' 全マクロを一括実行する
' 1. 全シートの先頭に空行を追加
' 2. シート一覧を生成・更新
Public Sub RunAll()
    Dim skipped As String
    skipped = AddRowsAndGoHome_Internal()
    UpdateSheetList

    If Len(skipped) > 0 Then
        MsgBox "以下のシートは保護されているためスキップしました:" _
            & skipped, vbInformation
    End If
End Sub

' シート一覧を再生成する
' 全シート名をハイパーリンク付きで一覧表示し、ブックの先頭に配置する
Public Sub UpdateSheetList()
    Dim ws As Worksheet
    Dim listSheet As Worksheet
    Dim i As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 既存の一覧シートを削除（存在しなくてもエラーにしない）
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(LIST_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo Cleanup

    ' 一覧シートを追加して一番左に配置
    Set listSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    listSheet.Name = LIST_SHEET_NAME

    ' ヘッダー
    With listSheet.Cells(1, 1)
        .Value = "シート名"
        .Font.Bold = True
    End With

    ' 各シート名をハイパーリンク付きで記入
    i = 2
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> LIST_SHEET_NAME Then
            listSheet.Hyperlinks.Add _
                Anchor:=listSheet.Cells(i, 1), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            i = i + 1
        End If
    Next ws

    listSheet.Columns("A:A").AutoFit

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "シート一覧の更新中にエラーが発生しました: " _
            & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub

' 全シートの先頭に2行追加し、一番左のシートへ移動する
' シート一覧・保護シートは処理対象から除外する
Public Sub AddRowsAndGoHome()
    Dim skipped As String
    skipped = AddRowsAndGoHome_Internal()
    ThisWorkbook.Worksheets(1).Activate

    If Len(skipped) > 0 Then
        MsgBox "以下のシートは保護されているためスキップしました:" _
            & skipped, vbInformation
    End If
End Sub

' 空行追加の実処理（MsgBoxを出さずにスキップ情報を返す）
Private Function AddRowsAndGoHome_Internal() As String
    Dim ws As Worksheet
    Dim skipped As String

    On Error GoTo ErrHandler

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> LIST_SHEET_NAME Then
            If ws.ProtectContents Then
                skipped = skipped & vbNewLine & "  " & ws.Name
            Else
                ws.Rows("1:2").Insert Shift:=xlDown
            End If
        End If
    Next ws

    ThisWorkbook.Worksheets(1).Activate
    AddRowsAndGoHome_Internal = skipped
    Exit Function

ErrHandler:
    MsgBox "空行追加中にエラーが発生しました: " _
        & Err.Description, vbExclamation
    Err.Clear
    AddRowsAndGoHome_Internal = skipped
End Function
