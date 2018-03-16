Option Explicit

'//*********************************************************
'//* Global var definitiions
'//*********************************************************
'-- var String
Dim g_strDataHome
Dim g_strTableName
Dim g_strDataDir
Dim g_strConnStr

'//*********************************************************
'//* @procedure window_onload
'//*********************************************************
Private Sub window_onload
	'-- var Object
	Dim objArgs

	objArgs = window.dialogArguments
	
	g_strDataHome = objArgs(0)
	g_strTableName = objArgs(1)
	g_strDataDir = objArgs(2)
	g_strConnStr = objArgs(3)

'--	g_strDataHome = "C:\Users\winridge\Documents\mydata\csv"
'--	g_strTableName = "review_logs_to_rfces#csv"
'--	g_strDataDir = "C:\Users\winridge\Documents\tools\hta\dbtest\form\..\data\"
	body_onload
End Sub

'//*********************************************************
'//* @procedure body_onload
'//*********************************************************
Private Sub body_onload
	Dim strHtml
	
	strHtml = createForm_Basic( _
			g_strDataHome, _
			g_strTableName, _
			"Register", _
			g_strConnStr _
			)

	document.getElementById("id_data_form").innerHtml = strHtml		
End Sub

'//*********************************************************
'//* @procedure cmdRegisterData_Click
'//*********************************************************
Private Sub cmdRegisterData_Click
	'-- var Object
	Dim objCn
	Dim objRs
	Dim objField
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strText
	
	'-- var Integer
	Dim i
	
'--	strConnStr = createConnectionStringCsv(g_strDataHome)
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	objCn.Open g_strConnStr
	
	strSql = "select max(id) as MAX_ID from %1;"
	strSql = Replace(strSql, "%1", g_strTableName)
	
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	If f_item(0).value = "" Then f_item(0).value = objRs("MAX_ID") + 1
	
	objRs.Close
'--	Set objRs = Nothing
	
	strSql = "select top 1 * from %1;"
	strSql = Replace(strSql, "%1", g_strTableName)
	
	objRs.Open strSql, objCn, adOpenDynamic, adLockOptimistic
	on error resume next
	objRs.AddNew
	If Err Then msgbox "addnew result:" & err.description
	i = 0
	strText = ""
	For Each objField In objRs.Fields
		objField.value = f_item(i).value
		strText = strText & f_item(i).value & vbCrLf
		i = i + 1
	Next
'--	msgbox strText
	objRs.Update

	If Err Then msgbox "Update result:" & err.description:Exit Sub
	
	msgbox "ìoò^äÆóπÇµÇ‹ÇµÇΩÅB",vbInformation,"ìoò^"

	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	Set objCn = Nothing
End Sub
