Attribute VB_Name = "mUtilerias"
Public Sub initPerformance()
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
End Sub

Public Sub endPerformance()
  Application.ScreenUpdating = True
  Application.DisplayStatusBar = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
End Sub

Public Function cpyShtByNam(aShtNam As String, aShtTmp As Worksheet) As Worksheet
  Dim lSht As Worksheet
  
  ' Verifica: Existe una hoja con el mismo nombre.
  On Error Resume Next
  Set lSht = Worksheets(aShtNam)
  On Error GoTo 0
  If Not lSht Is Nothing Then
    Application.DisplayAlerts = False
    lSht.Delete
    Application.DisplayAlerts = True
    Set lSht = Nothing
  End If
  
  ' Copia: Hoja.
  aShtTmp.Visible = xlSheetVisible
  aShtTmp.Copy After:=aShtTmp
  aShtTmp.Visible = xlSheetHidden
  Set lSht = ActiveSheet
  lSht.Name = aShtNam
  
  Set cpyShtByNam = lSht
End Function

Public Function openExcelFile(sDialogTitle As String) As String
  Dim intChoice As Integer
  Dim strPath As String

  'Solo se permite seleccionar un archivo a la vez
  Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
  'Cambio el título del diálogo
  Application.FileDialog(msoFileDialogOpen).Title = sDialogTitle
  'Elimino los filtors anteriores
  Call Application.FileDialog(msoFileDialogOpen).Filters.Clear
  'Agrego los nuevos filtros.
  Call Application.FileDialog(msoFileDialogOpen).Filters.Add("Archivos de MS Excel", "*.xls, *.xlsx")
  'Mostramos el diálogo para seleccionar archivos
  intChoice = Application.FileDialog(msoFileDialogOpen).Show
  'determine what choice the user made
  If intChoice <> 0 Then
      'Obtenemos la ruta al archivo
      strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
  End If

  openExcelFile = strPath
    
End Function
