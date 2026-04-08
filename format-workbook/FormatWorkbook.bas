Sub FormatWorkbook()
    '全シートの書式を統一するマクロ（MacOS対応版）
    'セル・グラフ・テーブル・図形すべてに適用

    Dim ws As Worksheet
    Dim targetFont As String
    Dim cht As ChartObject
    Dim tbl As ListObject
    Dim shp As Shape
    Dim srs As Series

    'フォント名を設定
    targetFont = "Meiryo UI"

    '画面更新を停止してパフォーマンス向上
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    '全シートをループ処理
    For Each ws In ThisWorkbook.Worksheets
        With ws
            'セルのフォントを統一
            On Error Resume Next  'フォントが存在しない場合のエラーを無視
            .Cells.Font.Name = targetFont
            On Error GoTo 0

            'テーブル（ListObject）のフォントを統一
            If .ListObjects.Count > 0 Then
                For Each tbl In .ListObjects
                    On Error Resume Next
                    tbl.Range.Font.Name = targetFont
                    If Not tbl.HeaderRowRange Is Nothing Then
                        tbl.HeaderRowRange.Font.Name = targetFont
                    End If
                    If Not tbl.TotalsRowRange Is Nothing Then
                        tbl.TotalsRowRange.Font.Name = targetFont
                    End If
                    On Error GoTo 0
                Next tbl
            End If

            'グラフのフォントを統一
            If .ChartObjects.Count > 0 Then
                For Each cht In .ChartObjects
                    With cht.Chart
                        'グラフタイトル
                        On Error Resume Next
                        If .HasTitle Then
                            .ChartTitle.Font.Name = targetFont
                        End If
                        On Error GoTo ErrHandler

                        '軸ラベル・軸タイトル（軸が存在しない場合があるためエラー無視）
                        On Error Resume Next
                        .Axes(xlCategory).TickLabels.Font.Name = targetFont
                        .Axes(xlValue).TickLabels.Font.Name = targetFont
                        If .Axes(xlCategory).HasTitle Then
                            .Axes(xlCategory).AxisTitle.Font.Name = targetFont
                        End If
                        If .Axes(xlValue).HasTitle Then
                            .Axes(xlValue).AxisTitle.Font.Name = targetFont
                        End If
                        On Error GoTo ErrHandler

                        '凡例
                        On Error Resume Next
                        If .HasLegend Then
                            .Legend.Font.Name = targetFont
                        End If
                        On Error GoTo ErrHandler

                        'データラベル
                        On Error Resume Next
                        For Each srs In .SeriesCollection
                            If srs.HasDataLabels Then
                                srs.DataLabels.Font.Name = targetFont
                            End If
                        Next srs
                        On Error GoTo ErrHandler
                    End With
                Next cht
            End If

            '図形（テキストボックス・オートシェイプ・フリーフォーム等）のフォントを統一
            If .Shapes.Count > 0 Then
                For Each shp In .Shapes
                    On Error Resume Next
                    shp.TextFrame.Characters.Font.Name = targetFont
                    On Error GoTo ErrHandler
                Next shp
            End If

            'シートをアクティブ化してからZoom設定（これが重要）
            .Activate
            ActiveWindow.Zoom = 80

            'A1セルを選択
            .Range("A1").Select
        End With
    Next ws

    '最初のシートをアクティブにして終了
    ThisWorkbook.Worksheets(1).Activate
    ThisWorkbook.Worksheets(1).Range("A1").Select

    MsgBox "書式の統一が完了しました。" & vbCrLf & _
           "フォント: " & targetFont & vbCrLf & _
           "表示倍率: 80%" & vbCrLf & _
           "適用対象: セル、テーブル、グラフ、図形", vbInformation, "完了"

    GoTo Cleanup

ErrHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbExclamation, "エラー"

Cleanup:
    Application.ScreenUpdating = True

End Sub