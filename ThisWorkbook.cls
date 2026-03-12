Option Explicit

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    UpdateSheetList
End Sub

Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object)
    ' 削除完了後に一覧を更新（OnTimeで遅延実行）
    Application.OnTime Now + TimeSerial(0, 0, 1), "UpdateSheetList"
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' シート名のタブ変更はSheetChangeでは検知不可
    ' 必要に応じてUpdateSheetListを手動実行する
End Sub
