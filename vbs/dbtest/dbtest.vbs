Option Explicit

'//*********************************************************
'//* @procedure cmdExecSqlCommand_Click
'//*********************************************************
Private Sub cmdExecSqlCommand_Click()
	'-- var Object
	Dim objCn
	Dim objRs
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath
	
	'-- var Integer
	Dim intRecNum
	

	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	strConnStr = createConnectionStringCsv("C:\Users\winridge\Documents\mydata\csv")

	strConnStr = "File name=C:/Users/winridge/Documents/tools/vbs/dbtest/csv.udl"

	objCn.Open strConnStr
	
	strSql = f_txa_sql_command.value
	
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	createTable_Basic objRs, "id_view_result", "id"

	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	Set objCn = Nothing
End Sub

'//*********************************************************
'//* @procedure cmdExecSqlCommand_Click
'//*********************************************************
Private Sub Main()
	cmdExecSqlCommand_Click
End Sub
