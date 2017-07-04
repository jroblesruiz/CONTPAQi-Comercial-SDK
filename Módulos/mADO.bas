Attribute VB_Name = "mADO"

' Publica conexión SQL
Public Function getSQLConnection() As ADODB.Connection
  Dim lConnectionString As String
  Dim lSQLConnection As ADODB.Connection
  
  Set getSQLConnection = Nothing
  
  ' RN: El nombre del servidor es obligatorio.
  If Len(Trim$(shtParametros.Range("prmSQLServer"))) = 0 Then
    MsgBox "El nombre del servidor es obligatorio.", vbCritical, "SQL Server"
    Exit Function
  End If
  
  ' RN: El nombre del usuario es obligatorio.
  If Len(Trim$(shtParametros.Range("prmSQLUser"))) = 0 Then
    MsgBox "El nombre del usuario es obligatorio.", vbCritical, "SQL Usuario"
    Exit Function
  End If
  
  ' RN: La contraseña del usuario es obligatoria.
  If Len(Trim$(shtParametros.Range("prmSQLPassword"))) = 0 Then
    MsgBox "La contraseña del usuario es obligatoria..", vbCritical, "SQL Constreña"
    Exit Function
  End If
  
  ' Conexión.
  lConnectionString = "Provider=SQLOLEDB;Data Source={sql_srv};Initial Catalog=master;User Id={sql_usr};Password={sql_psw};"
  lConnectionString = Replace(lConnectionString, "{sql_srv}", shtParametros.Range("prmSQLServer").Value)
  lConnectionString = Replace(lConnectionString, "{sql_ali}", shtParametros.Range("prmSQLDbAlias").Value)
  lConnectionString = Replace(lConnectionString, "{sql_usr}", shtParametros.Range("prmSQLUser").Value)
  lConnectionString = Replace(lConnectionString, "{sql_psw}", shtParametros.Range("prmSQLPassword").Value)
  Set lSQLConnection = New ADODB.Connection
  lSQLConnection.CommandTimeout = 2
  On Error Resume Next
  lSQLConnection.Open lConnectionString
  On Error GoTo 0
  If lSQLConnection.State = 0 Then
    Set lSQLConnection = Nothing
  End If
  
  Set getSQLConnection = lSQLConnection
End Function

' Ejecuta: Consulta.
Public Function executeQuery(lCommand As ADODB.Command) As ADODB.Recordset
  Dim lConnection As ADODB.Connection
  Dim lRecordset As ADODB.Recordset
  Dim lRecordsAffected As Integer
  
  Set executeQuery = Nothing
  
  ' RN: Error en la conexión con el servidor SQL, por favor verifique paramentros.
  Set lConnection = getSQLConnection()
  If lConnection Is Nothing Then
    MsgBox "Error en la conexión con el servidor SQL, por favor verifique parámetros.", vbCritical, "Conexión SQL"
    Exit Function
  End If
  
  Set lCommand.ActiveConnection = lConnection
  Set lRecordset = lCommand.Execute(RecordsAffected:=lRecordsAffected)
  If lRecordsAffected = 0 Then
    Set lRecordset = Nothing
    Exit Function
  End If
  
  Set executeQuery = lRecordset
End Function

