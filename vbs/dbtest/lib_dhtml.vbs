Option Explicit

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
			"<tr>" & vbCrLf

		For Each objField In p_objRs.Fields
			If objField.name = "content" _
			Or objField.Name = "comment" _
			Or InStr(objField.Name, "_éwìE") _
			Or InStr(objField.Name, "_ì‡óe") Then
			strHtml = strHtml & _
				"<td valign=""top""><pre>" & objField.Value & "</pre></td>" & vbCrLf
			ElseIf objField.name = "description" _
			Or objfield.Name = "äiî[êÊ" Then
			strHtml = strHtml & _
				"<td valign=""top""><textarea rows=""" & lenb(objField.value)/80 & """ cols=""80"">" & objField.Value & "</textarea></td>" & vbCrLf
			ElseIf objField.name = "path" Then
				strDriveName = ""
				If InStr(objField.Value, "C:\") = 0 Then
					strDriveName = p_objRs("drive_name")
				End If
				strHtml = strHtml & _
					"<td><a href=""" & strDriveName & objField.Value & """>" & objField.Value & "</a></td>" & vbCrLf
			ElseIf objField.name = "file_name" Then
				strDriveName = ""
				If InStr(p_objRs("path"), "C:\") = 0 Then
					strDriveName = p_objRs("drive_name")
				End If
				strHtml = strHtml & _
					"<td><a href=""" & strDriveName & p_objRs("path") & objField.Value & """>" & objField.Value & "</a></td>" & vbCrLf
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
				strHtml = strHtml & _
					"<td>" & objField.Value & "<br/></td>" & vbCrLf
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

