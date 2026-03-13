Option Explicit

Private Const LIST_SHEET_NAME As String = "シート一覧"

' シート一覧を生成・更新する
Public Sub UpdateSheetList()
    Dim ws As Worksheet
    Dim listSheet As Worksheet
    Dim i As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(LIST_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set listSheet = Worksheets.Add(Before:=Worksheets(1))
    listSheet.Name = LIST_SHEET_NAME

    With listSheet.Cells(1, 1)
        .Value = "シート名"
        .Font.Bold = True
    End With

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
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' 全シートの先頭に2行追加し、一番左のシートへ移動する
' シート一覧・保護シートは処理対象から除外する
' ※ 実行するたびに2行ずつ増えるため、必要なときだけ実行すること
Public Sub AddRows()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> LIST_SHEET_NAME Then
            If Not ws.ProtectContents Then
                ws.Rows("1:2").Insert Shift:=xlDown
            End If
        End If
    Next ws

    ThisWorkbook.Worksheets(1).Activate
End Sub
