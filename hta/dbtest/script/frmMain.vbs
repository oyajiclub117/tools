Option Explicit

'-- gloabl var Object

'//*********************************************************
'//* @procedure cmdExecSqlCommand_Click
'//*********************************************************
Private Sub cmdExecSqlCommand_Click( _
	byval p_strSqlType _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	Dim objValues
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath
	
	'-- var Integer
	Dim intRecNum
	
	'-- Is db type selected
	If f_ddl_db_types.value = "-" Then Exit Sub
	
	objValues = Split(f_ddl_db_types.value, ",")

	strConnStr = getActiveConnection(objValues(0))

	If strConnStr = "" Then Exit Sub

	If f_txa_sql_command.value = "" Then
		Msgbox "sql command nothing!!",vbCritical,"未入力エラー"
		Exit Sub
	End If
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	objCn.Open strConnStr

	strSql = f_txa_sql_command.value
	
	If p_strSqlType = "query" Then
		objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	Else
		objCn.Execute strSql,intRecNum
		If intRecNum Then Msgbox "command successfully!!", vbInformation:Exit Sub
	End If
	
	createTable_Basic objRs, "id_view_result", "id"
	
	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	Set objCn = Nothing
End Sub

'//*********************************************************
'//* @procedure viewSchema_Tables
'//*********************************************************
Private Sub viewSchema_Tables( _
	byval p_strKey _
	)
	'-- var Object
	Dim objRs
	Dim objValues
	
	'-- var String
	Dim strPath
	Dim strHtml
	Dim strTableName
	Dim strConnStr

	If f_ddl_db_types.value = "-" Then Exit Sub
	
	objValues = Split(f_ddl_db_types.value, ",")

	strConnStr = getActiveConnection(objValues(0))
	
	If strConnStr = "" Then Exit Sub
	
'##	If objValues(0) = "excel" _
'##	Then
		'--createHtmlDDL_ExcelTables
'##	Else
		Set objRs = getSchemaInfo(strConnStr, adSchemaTables)
	
		strHtml = "<select name=""" & FORM_NAME_DDL_TABLES & """ onchange=""cmdViewList_Click"">" & vbCrLf & _
			"<option value=""-"">-- tables --</option>" & vbCrLf
		
		Do Until objRs.EOF
			If p_strKey = "" _
			Or InStr(objRs("table_name"),p_strKey) Then
				strTableName = objRs("table_name")
				strHtml = strHtml & _
					"<option value=""" & strTableName & """>" & strTableName & "</option>" & vbCrLf
			End If

			objRs.MoveNext
		Loop
		
		strHtml = strHtml & _
			"</select>" & vbCrLf
'##	End If

	strHtml = strHtml & _
		"<span><input type=""text"" name=""f_txf_table_key"" size=""10"" max_size=""40""></input>" & vbCrLf & _
		"<button onclick=""viewSchema_Tables f_txf_table_key.value"">search</button></span>" & vbCrLf
	
		
	document.getElementById(FORM_ID_DDL_TABLES).InnerHtml = strHtml
	'-- clear file path
End Sub

'//*********************************************************
'//* @procedure cmdAddNew_Click
'//*********************************************************
Private Sub cmdAddNew_Click
	'-- var Object
	Dim objArgs
	Dim objValues
	
	'-- var String
	Dim strPath
	Dim strConnStr

	'-- Is db type selected
	If f_ddl_db_types.value = "-" Then Exit Sub
	
	objValues = Split(f_ddl_db_types.value, ",")

	strConnStr = getActiveConnection(objValues(0))

	If strConnStr = "" Then Exit Sub

	If f_ddl_tables.value = "-" Then Msgbox "table not selected!!",vbCritical,"選択エラー":Exit Sub

	objArgs = Array(g_strDataHome, "[" & f_ddl_tables.value & "]", g_strDataDir, strConnStr)

	showWindow_ModalDialog _
		"frmRegister.html", _
		objArgs, _
		800, _
		600
	'-- redrow
	cmdViewList_Click
End Sub

'//*********************************************************
'//* @procedure cmdActionEventExcel_Click
'//*********************************************************
Private Sub cmdActionEventExcel_Click
	'-- var Object
	Dim objVals
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath
	
	If f_ddl_func_lists.value = "-" Then Exit Sub

	objVals = Split(f_ddl_func_lists.value, ",")
	strSql = getSqlCommand(objVals(0))
	strSql = Replace(strSql, "%1", objVals(1))
	strSql = Replace(strSql, "%2", getSqlFilter(objVals(2)))
	strSql = Replace(strSql, "%3", f_txt_keyword.value)
	f_txa_sql_command.value = strSql
	
	cmdExecSqlCommand_Click "query"
End Sub

'//*********************************************************
'//* @procedure cmdSelectDbType_Click
'//*********************************************************
Private Sub cmdSelectDbType_Click( _
	byref p_objForm _
	)
	'-- var Object
	Dim objOpt
	Dim objValues
	Dim objForm
	
	If p_objForm.value = "-" Then Exit Sub
	
	For Each objOpt In f_ddl_db_types.Options
		If objOpt.value <> "-" Then
			objValues = Split(objOpt.value, ",")
			Set objForm = document.getElementById(objValues(1))
			objForm.style.display = "none"
		End If
	Next

	objValues = Split(f_ddl_db_types.value, ",")
	
	document.getElementById(objValues(1)).style.display = "inline"
End Sub

'//*********************************************************
'//* @procedure getActiveConnection
'//*********************************************************
Private Function getActiveConnection( _
	byval p_strDbType _
	)
	'-- var String
	Dim strConnStr
	Dim strPath
	Dim strUser
	Dim strPassword
	'-- var Object
	Dim objFields

	'-- clear connectio string
	strConnStr = ""

	Select Case p_strDbType
		Case "udl"
			If f_ddl_udl_files.value <> "-" Then
				strConnStr = createConnectionStringUdl(f_ddl_udl_files.value)
			Else
				'--Msgbox "udl file not selected!!",vbCritical,"選択エラー"
			End If
		Case "csv"
			If f_ddl_file_paths.value <> "-" Then
				If InStr(f_ddl_file_paths.value, "\") Then
					strPath = f_ddl_file_paths.value
				Else
					strPath = g_strDataDir & f_ddl_file_paths.value
				End If
				g_strDataHome = strPath
				strConnStr = createConnectionStringCsv(strPath)
			Else
				'--Msgbox "file path not selected!!",vbCritical,"選択エラー"
			End If
		Case "excel"
			If f_file_path_excel.value <> "" Then
				strPath = f_file_path_excel.value
				If f_cbx_row_header.checked = false Then
				strConnStr = createConnectionStringExcel_NOHD(strPath)
				Else
				strConnStr = createConnectionStringExcel(strPath)
				End If
				'--openExcelBook strPath
			ElseIf f_excel_paths.value <> "-" Then
				strPath = getFilePath_Excel(f_excel_paths.value)
				If f_cbx_row_header.checked = false Then
				strConnStr = createConnectionStringExcel_NOHD(strPath)
				Else
				strConnStr = createConnectionStringExcel(strPath)
				End If
				'--openExcelBook strPath
			Else
				'--Msgbox "file-path not selected!!",vbCritical,"選択エラー"
			End If
		Case "firebird"
			If f_file_path_firebird.value <> "" Then
				strPath = f_file_path_excel.value
				strUser=f_txt_user.value
				strPassword=f_txt_password.value
				strConnStr = createConnectionStringFirebird(strPath, strUser, strPassword)
			ElseIf f_firebird_paths.value <> "-" Then
				objFields = Split(f_firebird_paths.value, ",")
				If UBound(objFields) <> 2 Then Msgbox "schema-id,user,passwd not exist!!":Exit Function
				
				strPath = getFilePath_Excel(objFields(0))
				strUser=objFields(1)
				strPassword=objFields(2)
				strConnStr = createConnectionStringFirebird(strPath, strUser, strPassword)
			Else
				'--Msgbox "file-path not selected!!",vbCritical,"選択エラー"
			End If
		Case Else
				Msgbox "db type unknown!! [type=" & f_ddl_db_types.value & "]",vbCritical,"選択エラー"
	End Select

	getActiveConnection = strConnStr
End Function

'//*********************************************************
'//* @procedure cmdCreateCsvFile_Click
'//*********************************************************
Private Sub cmdCreateCsvFile_Click( _
	byval p_strDestFile _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	Dim objValues
	Dim objFields
	Dim objWshShell
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath
	Dim strDesktop
	Dim strCsvHead
	Dim strCsvData
	
	'-- Is db type selected
	If f_ddl_db_types.value = "-" Then Exit Sub
	
	objValues = Split(f_ddl_db_types.value, ",")

	strConnStr = getActiveConnection(objValues(0))

	If strConnStr = "" Then Exit Sub

	If f_txa_sql_command.value = "" Then
		Msgbox "sql command nothing!!",vbCritical,"未入力エラー"
		Exit Sub
	End If
	
	Set g_objFso = CreateObject("Scripting.FileSystemObject")
	set objWshShell = WScript.CreateObject("WScript.Shell")
    strDesktop = objWshShell.SpecialFolders("Desktop")
    strPath = strDesktop & "\" & p_strDestFile

	If g_objFso.FileExists(strPath) Then
		Set g_objCsvFile = g_objFso.OpenTextFile(strPath, ForAppending)
	Else
		Set g_objCsvFile = g_objFso.CreateTextFile(strPath)
		createCsvFile_Head
	End If
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	objCn.Open strConnStr
	
	strSql = f_txa_sql_command.value
	
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	
	strCsvHead = ""
	For Each objField In objRs.Fields
		strCsvHead = strCsvHead & objField.Name & ","
	Next
	
	g_objCsvFile.WriteLine Left(strCsvHead, Len(strCsvHead) -1)
	
'--	createTable_Basic objRs, "id_view_result", "id"
	strCsvData = ""
	Do Until objRs.EOF
		For Each objField In objRs.Fields
			strCsvHead = strCsvHead & objField.Name & ","
		Next
		
		g_objCsvFile.WriteLine Left(strCsvHead, Len(strCsvHead) -1)

		objRs.MoveNext
	Loop
	
	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	Set objCn = Nothing
End Sub

'//*********************************************************
'//* @procedure cmdGetSchemaInfo_Click
'//*********************************************************
Private Sub cmdGetSchemaInfo_Click
	'-- var Object
	Dim objRs
	Dim objValues
	
	'-- var String
	Dim strConnStr

	If f_ddl_db_types.value = "-" Then Exit Sub
	
	objValues = Split(f_ddl_db_types.value, ",")

	strConnStr = getActiveConnection(objValues(0))
	
	If strConnStr = "" Then Exit Sub
	
'	Set objRs = getSchemaInfo(strConnStr, adSchemaColumnsDomainUsage)
	Set objRs = getSchemaInfo(strConnStr, adSchemaColumns)
	
	createTable_Basic objRs, "id_view_result", "id"
	
	If IsObject(objRs) Then objRs.Close
	
	Set objRs = Nothing
End Sub

'//*********************************************************
'//* @procedure cmdActionToDo_Click
'//*********************************************************
Private Sub cmdActionToDo_Click
Const SQL_ID_ACTION_TODO = "19"
Const DDL_DB_TYPE_CSV = 2
Const DDL_FILE_PATH_SHARE = 2
	'-- var Object
	'-- var String
	Dim strSql
	
	f_ddl_db_types.options(DDL_DB_TYPE_CSV).selected = true
	f_ddl_file_paths.options(DDL_FILE_PATH_SHARE).selected = true
	
	strSql = getSqlCommand(SQL_ID_ACTION_TODO)
	f_txa_sql_command.value = strSql
	
	cmdExecSqlCommand_Click "query"
End Sub

'//*********************************************************
'//* @procedure openExcelBook
'//*********************************************************
Private Sub openExcelBook( _
	byval p_strPath _
	)
	'-- var Object
	Dim objExcel
	Dim objWkb

	If InStr(p_strPath,".xlsx") = False Then Exit Sub

	Set objExcel = CreateObject("Excel.Application")

	objExcel.Visible = True

	'-- Is Already Open Workbook
	Set objWkb = objExcel.Workbooks.Open(p_strPath,,False)
End Sub
