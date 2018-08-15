Sub readdb()
  Dim FBcon As ADODB.Connection
  Dim FBcmd As ADODB.Command

  Set FBcon = CreateObject("ADODB.Connection")
  FBcon.Provider = "MSDASQL.1"
  FBcon.ConnectionString = "ODBC;DSN=ABC;DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA;PWD=masterkey;DBNAME=c:\my.fdb"
  FBcon.Open

  Set FBcmd = CreateObject("ADODB.Command")
  FBcmd.ActiveConnection = FBcon
  FBcmd.CommandType = adCmdText
  FBcmd.CommandText = "select * from BTD"
  Set r = FBcmd.Execute

  Dim i As Integer
  i = 1
  Do While Not r.EOF
    Dim j As Integer
    For j = 0 To r.Fields.Count - 1 Step 1
      Cells(i, j + 1) = r.Fields(j).Value
    Next j
    i = i + 1
    r.MoveNext
  Loop
  FBcon.Close
  Set FBcon = Nothing
End Sub
