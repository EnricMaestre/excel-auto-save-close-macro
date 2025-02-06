Attribute VB_Name = "AutoSaveCloseMacro"
Dim TiempoInactivo As Date

Sub ReiniciarTemporizador()
    On Error Resume Next
    Application.OnTime EarliestTime:=TiempoInactivo, Procedure:="CerrarExcel", Schedule:=False
    On Error GoTo 0
    
    TiempoInactivo = Now + TimeValue("00:05:00")
    Application.OnTime EarliestTime:=TiempoInactivo, Procedure:="CerrarExcel", Schedule:=True
End Sub

Sub CerrarExcel()
    If Now >= TiempoInactivo Then
        Application.DisplayAlerts = False
        ThisWorkbook.Save
        ThisWorkbook.Close
        Application.DisplayAlerts = True
    End If
End Sub

Private Sub Workbook_Open()
    ReiniciarTemporizador
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ReiniciarTemporizador
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ReiniciarTemporizador
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ReiniciarTemporizador
End Sub
