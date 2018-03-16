Option Explicit

'-- form id 
Const FORM_ID_DDL_TABLES = "id_ddl_tables"
Const FORM_NAME_DDL_TABLES = "f_ddl_tables"
Const FORM_SIZE_TEXTAREA_COLUMN_MAX = "80"
Const FILE_PATH_CONFIG = "csv"
Const DRIVE_NAME_HOME = "C:"

'//*********************************************************
'//* @procedure createTable_Basic
'//* @arg1 [form-id]
'//*********************************************************
Private Sub createTable_Basic( _
	byref p_objRs, _
	byval p_strFormId, _
	byval p_strId _
	)
	'-- var Object
	Dim objField
	
	'-- var String
	Dim strHtml
	Dim strDriveName
	
	'-- var Integer
	Dim intPage
	Dim intPageSize
	Dim intPageMax
	
	strHtml = "<table border=""3"" cellpadding=""3"" cellspacing=""3"">" & vbCrLf
	
	'-- set header
	strHtml = strHtml & _
		"<tr>" & vbCrLf
	Dim x
	For Each objField In p_objRs.Fields
		strHtml = strHtml & _
			"<th>" & objField.Name & "</th>" & vbCrLf
	Next
	
	strHtml = strHtml & _
		"</tr>" & vbCrLf

	'-- set data
	intPage = 0
	intPageSize = 0
	intPageMax = 0
	
'--	If p_objRs.PageCount > 0 Then
'--		intPageMax = p_objRs.PageCount
'--		intPageSize = p_objRs.PageSize
'--	End If
'--	msgbox p_objRs.PageCount & ":" & p_objRs.PageSize
	Do Until p_objRs.EOF
		strHtml = strHtml & _
			"<tr valign=""top"">" & vbCrLf

		For Each objField In p_objRs.Fields
			If objField.name = "content" _
			Or objField.Name = "comment" _
			Or objField.Name = "ì¸óÕ" _
			Or objField.Name = "äTóv_" _
			Or objField.Name = "èoóÕ" _
			Or InStr(objField.Name, "_éwìE") _
			Or InStr(objField.Name, "_ì‡óe") _
			Then
			strHtml = strHtml & _
				"<td valign=""top""><pre>" & objField.Value & "</pre></td>" & vbCrLf
			ElseIf objField.name = p_strId _
			Then
				strHtml = strHtml & _
					"<td valign=""top""><input type=""checkbox"" name=""f_cbx"">" & objField.value & "</input></td>" & vbCrLf
			ElseIf objField.name = "description" _
			Or objfield.Name = "äiî[êÊ" Then
				strHtml = strHtml & _
					"<td valign=""top""><textarea rows=""" & lenb(objField.value)/40 & """ cols=""40"">" & objField.Value & "</textarea></td>" & vbCrLf
			ElseIf objField.name = "path" Then
				strDriveName = ""
				If InStr(objField.Value, DRIVE_NAME_HOME) = 0 Then
					strDriveName = DRIVE_NAME_HOME
				End If
				strHtml = strHtml & _
					"<td><a href=""" & strDriveName & objField.Value & """ title=""" & strDriveName & objField.Value & """>" & "[link]</a><pre>" & createTree(strDriveName & objField.Value) & "</pre>" & "</td>" & vbCrLf
			ElseIf InStr(objField.name, "tree_path") Then
				strDriveName = ""
				If InStr(objField.name, DRIVE_NAME_HOME) = 0 Then
					strDriveName = DRIVE_NAME_HOME
				End If
				strHtml = strHtml & _
					"<td><a href=""" & strDriveName & objField.Value & """ title=""" & strDriveName & objField.Value & """>" & "[link]</a><pre>" & createTree(strDriveName & objField.Value) & "</pre>" & "</td>" & vbCrLf
			ElseIf objField.name = "file_path" Then
				strDriveName = ""
				If InStr(objField.Value, DRIVE_NAME_HOME) = 0 Then
					strDriveName = DRIVE_NAME_HOME
				End If
				strHtml = strHtml & _
					"<td><a href=""" & strDriveName & objField.Value & """>" & strDriveName & objField.Value & "</a></td>" & vbCrLf
			ElseIf Len(objField.value) > 60 Then
				Dim objArr
				Dim intMaxRow
				objArr = Split(objField.value, vbLf)
				intMaxRow = Round(LenB(objField.value) / 60, 0) + Round(UBound(objArr) / 1.5, 0)
				strHtml = strHtml & _
					"<td valign=""top""><textarea style=""overflow:auto;"" rows=""" & intMaxRow & """ cols=""60"">" & objField.Value & "</textarea></td>" & vbCrLf
'--					"<td valign=""top""><textarea style=""overflow:auto;"" rows=""" & Round((LenB(objField.value))/60,0)+3 & """ cols=""60"">" & objField.Value & "</textarea></td>" & vbCrLf
'--					"<td valign=""top""><span style=""width:480px;"">" & objField.Value & "</span></td>" & vbCrLf
			Else
				If objField.Name = "id" Then
					strHtml = strHtml & _
						"<td nowrap=""true"">" & objField.Value & "<br/></td>" & vbCrLf
				Else
					strHtml = strHtml & _
						"<td>" & objField.Value & "<br/></td>" & vbCrLf
				End If
			End If
		Next
		
		strHtml = strHtml & _
			"</tr>" & vbCrLf

		'-- add page count
		intPage = intPage + 1
		
		'-- next data
'--		If intPageMax Then
'--			If intPage = intPageSize Then
'--				Exit Do
'--			End If
'--		End If
		p_objRs.MoveNext
	Loop
	
	strHtml = strHtml & _
		"</table>" & vbCrLf

	document.getElementById(p_strFormId).innerHTML = strHtml
	
	set objField = Nothing
End Sub

'//*********************************************************
'//* @procedure DEBUG_LOG
'//* @arg1 [form-id]
'//*********************************************************
Private Sub DEBUG_LOG( _
	p_strLog _
	)
	f_txa_view_debug.value = _
		f_txa_view_debug.value & p_strLog & vbCrLf
End Sub

'//*********************************************************
'//* @procedure createDDL_ExcelFilePaths
'//*********************************************************
Private Sub createDDL_ExcelFilePaths
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
	
	strHtml = "<select name=""f_ddl_excel_file_paths"" onchange=""cmdExecSqlCmdExcel_Click"">" & vbCrLf
	
	strHtml = strHtml & _
		"<option value=""-"">-- excel-file-paths --</option>" & vbCrLf
	Do Until objRs.EOF
		strHtml = strHtml & _
			"<option value=""" & objRs("key_") & """>" & objRs("value_") & "</option>" & vbCrLf
		objRs.MoveNext
	Loop

	strHtml = strHtml & _
		"</select>" & vbCrLf
		
	document.getElementById("id_excel_file_path").innerHTML = strHtml
	
	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	set objCn = Nothing
End Sub

'//*********************************************************
'//* @function createForm_Basic
'//*********************************************************
Private Function createForm_Basic( _
	byval p_strFilePath, _
	byval p_strTableName, _
	byval p_strFormType, _
	byval p_strConnStr _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	Dim objField
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath
	Dim strHtml
	'-- var Integer
	Dim i
	
'--	strConnStr = createConnectionStringCsv(p_strFilePath)
	strConnstr = p_strConnStr
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	objCn.Open strConnStr
	strSql = getSqlCommand(SQL_ID_TABLE_ALL)
	
	strSql = Replace(strSql, "%1", p_strTableName)
	
	strSql = Replace("select * from %1 where id = 1;", "%1",p_strTableName)
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	strHtml = "<table border=""3"" cellpadding=""3"" cellspacing=""3"">" & vbCrLf
	
	i = 0
	
	For Each objField In objRs.Fields
		strHtml = strHtml & _
			"<tr>" & vbCrLf
		
		If LenB(objField) > 80  Then
			strHtml = strHtml & _
				"<th>" & objField.Name & "</th>" & vbCrLf & _
				"<td valign=""top""><textarea name=""f_item"" rows=""" & _
				(Round((LenB(objField.Value)+FORM_SIZE_TEXTAREA_COLUMN_MAX)/FORM_SIZE_TEXTAREA_COLUMN_MAX,0)) & _
				""" cols=""" & FORM_SIZE_TEXTAREA_COLUMN_MAX & """></textarea>" & vbCrLf
		ElseIf InStr(objField.Name, "created") _
		Or InStr(objField.Name, "completed") _
		Or InStr(objField.Name, "expired") _
		Then
			strHtml = strHtml & _
				"<th>" & objField.Name & "</th>" & vbCrLf & _
				"<td><input name=""f_item"" size=""20""></input><button onclick=""vbscript:f_item(" & i & ").value = formatdatetime(now,0)"">now-date</button>" & vbCrLf
		ElseIf InStr(objField.Name, "content") _
		Or InStr(objField.Name, "description") _
		Or InStr(objField.Name, "command") _
		Or InStr(objField.Name, "sql") _
		Then
			strHtml = strHtml & _
				"<th>" & objField.Name & "</th>" & vbCrLf & _
				"<td valign=""top""><textarea name=""f_item"" rows=""5"" " & _
				"cols=""" & FORM_SIZE_TEXTAREA_COLUMN_MAX & """></textarea>" & vbCrLf
		Else
			strHtml = strHtml & _
				"<th>" & objField.Name & "</th>" & vbCrLf & _
				"<td><input name=""f_item"" size=""20""></input>" & vbCrLf
		End If
		strHtml = strHtml & _
			"</tr>" & vbCrLf
		
		i = i + 1
	Next
	 
	If p_strFormType = "view" Then
		strHtml = strHtml & _
			"<tr align=""center""><td colspan=""2""><button onclick=""vbscript:window.close"">close</button></tr>" & vbCrLf
	ElseIf p_strFormType = "modify" Then
		strHtml = strHtml & _
			"<tr align=""center"">" & vbCrLf & "<td colspan=""2"">" & vbCrLf & _
			"<button onclick=""cmdUpdateData_Click"">update</button>" & vbCrLf & _
			"<button onclick=""cmdUpdateData_Click"">delete</button>" & vbCrLf & _
			"<button onclick=""vbscript:window.close"">close</button>" & vbCrLf & _
			"</tr>" & vbCrLf
	Else
		strHtml = strHtml & _
			"<tr align=""center"">" & vbCrLf & "<td colspan=""2"">" & vbCrLf & _
			"<button onclick=""cmdRegisterData_Click"">Register</button>" & vbCrLf & _
			"<button onclick=""vbscript:window.close"">close</button>" & vbCrLf & _
			"</tr>" & vbCrLf
	End If
	strHtml = strHtml & _
		"</table>" & vbCrLf
	
	createForm_Basic = strHtml
End Function

'//*********************************************************
'//* @function createTree
'//* @arg1 [path]
'//*********************************************************
Private Function createTree( _
	p_strPath _
	)
	'-- var Object
	Dim objFields

	'-- var String
	Dim strTree
	Dim strIndent
	
	'-- var Long
	Dim i
	Dim lngNumOfFlds
	
	strTree = ""
	strIndent = "  "

	objFields = Split(p_strPath, "\")
	
	lngNumOfFlds = UBound(objFields)
	
	If lngNumOfFlds = 0 Then
		Msgbox "spliter failer path delimiter string not exist",vbCritical

		createTree = strTree
		
		Exit Function
	End If
	
	i = 0
	
	For i = 0 To lngNumOfFlds - 1
		If i = 0 Then
			strTree = objFields(i) & vbCrLf
		Else
			strTree = strTree & _
				"Ñ§Ñü" & _
				objFields(i) & vbCrLf & _
				strIndent
			strIndent = strIndent & "  "
		End If
	Next
	
	createTree = strTree
End Function

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
	
	strHtml = strHtml & _
		"<span><input type=""text"" name=""f_txf_table_key"" size=""10"" max_size=""40""></input>" & vbCrLf & _
		"<button onclick=""viewSchema_Tables f_txf_table_key.value"">search</button></span>" & vbCrLf
	
		
	document.getElementById(FORM_ID_DDL_TABLES).InnerHtml = strHtml
	'-- clear file path
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
'//* @procedure cmdViewList_Click
'//*********************************************************
Private Sub cmdViewList_Click
	'-- var String
	Dim strSql

	'-- var Integer
	Dim intSqlId
	
	'-- var Object
	Dim objFields
	
	If f_ddl_tables.value = "-" Then Exit Sub
	
	objFields = Split(f_ddl_db_types.value, ",")
	
	Select Case objFields(0)
		Case "udl","csv","excel"
			intSqlId = SQL_ID_GET_DATA_MS_DB
		Case Else
			intSqlId = SQL_ID_GET_DATA_NOT_MS_DB
	End Select
	
	strSql = Replace(getSqlCommand(intSqlId), "%1", f_ddl_tables.value)

	f_txa_sql_command.value = strSql
	
	cmdExecSqlCommand_Click "query"
End Sub

