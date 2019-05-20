Option Explicit
Const adOpenStatic = "3"
Const adLockReadOnly = "1"

Call Main()
Sub Main()
  On Error Resume Next

  Dim connect
  Dim rs
  Set connect = CreateObject("ADODB.Connection")
  Set rs = CreateObject("ADODB.Recordset")
  WScript.Echo "OracleDB ê⁄ë±äJén"
  Dim strConnStr
  strConnStr = "Provider=MSDASQL.1;Password=systemsss;Persist Security Info=True;User ID=jdbc_user;Extended Properties=""DRIVER={Oracle in XE};SERVER=XE;UID=system;PWD=systemsss;DBQ=XE;DBA=W;APA=T;EXC=F;XSM=Default;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=Me;CSR=F;FWC=F;FBS=60000;TLO=O;MLD=0;ODA=F;"""
'--  connect.Open "Driver={Microsoft ODBC for XE};" & _
'--              "CONNECTSTRING=XE; UID=system; PWD=systemsss;"
  connect.Open = strConnStr
  If Err.Number <> 0 Then
    WScript.Echo Err.Number
    WScript.Echo Err.Source
    WScript.Echo Err.Description
    WScript.Quit(-1)
  End If
  WScript.Echo "SQL é¿çs"
  rs.Open "select * from employee;",connect, adOpenStatic, adLockReadOnly
  If Err.Number <> 0 Then
    WScript.Echo Err.Number
    WScript.Echo Err.Source
    WScript.Echo Err.Description
    WScript.Quit(-1)
  End If
  Dim x,s
  s = ""
  For Each x in Rs.Fields
  	s = s & x.Name & ","
  Next
  s = s & vbCrLf
  Do until(rs.EOF)
	  For Each x in Rs.Fields
	  	s = s & x.Value
	  Next
	  s = s & vbCrLf
	  rs.MoveNext()
  Loop
  WScript.Echo s & vbCrLf

  WScript.Echo "OracleDB êÿíf"
  connect.Close
  Set connect = Nothing
End Sub
