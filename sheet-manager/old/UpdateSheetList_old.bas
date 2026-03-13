Public Sub UpdateSheetList()
    Dim ws As Worksheet
    Dim listSheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim i As Long

    ' アクティブなワークブックを対象とする
    Set targetWorkbook = ActiveWorkbook

    ' PERSONAL.XLSBやアドイン自体には実行しない
    If targetWorkbook Is Nothing Then
        MsgBox "ワークブックが開かれていません。", vbExclamation
        Exit Sub
    End If

    If InStr(1, targetWorkbook.Name, "PERSONAL.XLSB", vbTextCompare) > 0 Then
        MsgBox "PERSONAL.XLSBには実行できません。対象のワークブックを開いてください。", vbExclamation
        Exit Sub
    End If

    ' 高速化とエラー抑制
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error Resume Next
    Application.DisplayAlerts = False
    targetWorkbook.Worksheets("シート一覧").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 一覧シートを追加して一番左に配置
    Set listSheet = targetWorkbook.Worksheets.Add(Before:=targetWorkbook.Worksheets(1))
    listSheet.Name = "シート一覧"
    listSheet.Cells(1, 1).Value = "シート名"

    ' 各シート名をハイパーリンク付きで記入
    i = 2
    For Each ws In targetWorkbook.Worksheets
        If ws.Name <> "シート一覧" Then
            listSheet.Hyperlinks.Add _
                Anchor:=listSheet.Cells(i, 1), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            i = i + 1
        End If
    Next ws

    ' 各シート(シート一覧以外)の1行目に「シート一覧に戻る」リンクを追加
    For Each ws In targetWorkbook.Worksheets
        If ws.Name <> "シート一覧" Then
            Dim hasReturnLink As Boolean
            hasReturnLink = False

            ' 既に「シート一覧に戻る」リンクがあるかチェック
            On Error Resume Next
            If ws.Range("A1").Hyperlinks.Count > 0 Then
                If InStr(ws.Range("A1").Hyperlinks(1).TextToDisplay, "シート一覧") > 0 Then
                    hasReturnLink = True
                End If
            End If
            On Error GoTo 0

            ' リンクが存在しない場合のみ行を挿入
            If Not hasReturnLink Then
                ws.Rows(1).Insert Shift:=xlDown
            Else
                ' 既存リンクを削除して再作成
                On Error Resume Next
                ws.Range("A1").Hyperlinks(1).Delete
                On Error GoTo 0
            End If

            ' ハイパーリンクを追加
            ws.Hyperlinks.Add _
                Anchor:=ws.Cells(1, 1), _
                Address:="", _
                SubAddress:="'シート一覧'!A1", _
                TextToDisplay:="← シート一覧に戻る"
        End If
    Next ws

    ' 仕上げ
    listSheet.Columns("A:A").AutoFit
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
