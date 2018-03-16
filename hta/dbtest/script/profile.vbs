Option Explicit

'//********************************************************
'//* Gloabl Constant definitions
'//********************************************************
Const SQL_ID_LIST_FILE_PATH = 3
Const SQL_ID_FILE_PATH_EXCEL = 3
Const SQL_ID_SQL_FILTER = 4
Const SQL_ID_GET_DATA_MS_DB = 45
Const SQL_ID_GET_DATA_NOT_MS_DB = 46
Const SQL_ID_FILE_PATH_ALL_EXCEL = 48
Const SQL_ID_SCHEMA_FUNC = 73

'//********************************************************
'//* Gloabl var definitions
'//********************************************************
'-- var Object
Dim g_objFso
Dim g_objCsvFile

'-- var String
Dim g_strAppHome
Dim g_strDataDir
Dim g_strDataHome

'//*********************************************************
'//* @procedure window_onload
'//*********************************************************
Private Sub window_onload
	'-- var String
	Dim strCmdLine
	window.resizeTo 1800, 768
	
	Set g_objFso = CreateObject("Scripting.FileSystemObject")
	
	strCmdLine = Replace(oHtaApp.CommandLine,"""","")
	
	g_strAppHome = g_objFso.GetParentFolderName(strCmdLine) & "\..\"
	
	g_strDataDir = g_strAppHome & "data\"
	
'--	f_ddl_file_paths.value = "csv"
'--	viewSchema_Tables
'--	createDDL_ExcelFilePaths "viewSchema_Tables","id_excel_file_path"
	cmdFilePathKey_Click f_txf_file_path_key
	Call createDDL_ExcelFuncs( _
			"cmdActionEventExcel_Click", _
			"id_ddl_excel_funcs", _
			"" _
			)
End Sub

'//*********************************************************
'//* @procedure cmdFilePathKey_Click
'//*********************************************************
Private Sub cmdFilePathKey_Click( _
	byref p_objForm _
	)
	'-- var String
	Dim strFilter
	
	strFilter = "%%"
	'//If f_cbx_file_path_filter.checked Then
		strFilter = "%" & p_objForm.value & "%"
	'//End If
	document.getElementById("id_ddl_file_path_excel").innerHtml = _
		createDDL_FilePath_Excel(strFilter)
End Sub
'//*********************************************************
'//* @procedure createDDL_ExcelFilePaths
'//*********************************************************
Private Sub createDDL_ExcelFilePaths( _
	byval p_strAction, _
	byval p_strFormId _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strHtml

	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	strConnStr = createConnectionStringCsv(g_strDataDir&"csv")
	
	objCn.Open strConnStr
	
	strSql = getSqlCommand(SQL_ID_LIST_FILE_PATH)
	
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	strHtml = "<select name=""f_ddl_excel_file_paths"" onchange=""" & p_strAction & """>" & vbCrLf
	
	strHtml = strHtml & _
		"<option value=""-"">-- excel-file-paths --</option>" & vbCrLf
	Do Until objRs.EOF
		strHtml = strHtml & _
			"<option value=""" & objRs("key_") & """>" & objRs("value_") & "</option>" & vbCrLf
		objRs.MoveNext
	Loop

	strHtml = strHtml & _
		"</select>" & vbCrLf
		
	document.getElementById(p_strFormId).innerHTML = strHtml
	
	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	set objCn = Nothing
End Sub

'//*********************************************************
'//* @procedure createDDL_ExcelFuncs
'//*********************************************************
Private Sub createDDL_ExcelFuncs( _
	byval p_strAction, _
	byval p_strFormId, _
	byval p_strSchemaId _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strHtml

	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	strConnStr = createConnectionStringCsv(g_strDataDir&"csv")
	
	objCn.Open strConnStr
	
	strSql = getSqlCommand(SQL_ID_SCHEMA_FUNC)
	strSql = Replace(strSql,"%1", p_strSchemaId)
	
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	strHtml = "<select name=""f_ddl_func_lists"" onchange=""" & p_strAction & """>" & vbCrLf
	
	strHtml = strHtml & _
		"<option value=""-"">-- excel-funcs --</option>" & vbCrLf
	Do Until objRs.EOF
		strHtml = strHtml & _
			"<option value=""" & objRs("key_") & """>" & objRs("value_") & "</option>" & vbCrLf
		objRs.MoveNext
	Loop

	strHtml = strHtml & _
		"</select>" & vbCrLf
		
	document.getElementById(p_strFormId).innerHTML = strHtml
	
	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	set objCn = Nothing
End Sub

