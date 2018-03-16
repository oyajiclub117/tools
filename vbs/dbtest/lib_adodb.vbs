Option Explicit

'-- CursorTypeEnum
Const adOpenStatic = 3
Const adOpenDynamic = 2

'-- LockTypeEnum
Const adLockReadOnly = 1
Const adLockOptimistic = 3

'-- SchemaEnum
Const adSchemaAsserts = "0"
Const adSchemaCatalogs = "1"
Const adSchemaCharacterSets = "2"
Const adSchemaCheckConstraints = "5"
Const adSchemaCollations = "3"
Const adSchemaColumnPrivileges = "13"
Const adSchemaColumns = "4"
Const adSchemaColumnsDomainUsage = "11"
Const adSchemaConstraintColumnUsage = "6"
Const adSchemaConstraintTableUsage = "7"
Const adSchemaCubes = "32"
Const adSchemaDBInfoKeywords = "30"
Const adSchemaDBInfoLiterals = "31"
Const adSchemaDimensions = "33"
Const adSchemaForeignKeys = "27"
Const adSchemaHierarchies = "34"
Const adSchemaIndexes = "12"
Const adSchemaKeyColumnUsage = "8"
Const adSchemaLevels = "35"
Const adSchemaMeasures = "36"
Const adSchemaMembers = "38"
Const adSchemaPrimaryKeys = "28"
Const adSchemaProcedureColumns = "29"
Const adSchemaProcedureParameters = "26"
Const adSchemaProcedures = "16"
Const adSchemaProperties = "37"
Const adSchemaProviderSpecific = "-1"
Const adSchemaProviderTypes = "22"
Const AdSchemaReferentialConstraints = "9"
Const adSchemaSchemata = "17"
Const adSchemaSQLLanguages = "18"
Const adSchemaStatistics = "19"
Const adSchemaTableConstraints = "10"
Const adSchemaTablePrivileges = "14"
Const adSchemaTables = "20"
Const adSchemaTranslations = "21"
Const adSchemaTrustees = "39"
Const adSchemaUsagePrivileges = "15"
Const adSchemaViewColumnUsage = "24"
Const adSchemaViews = "23"
Const adSchemaViewTableUsage = "25"

Const SQL_ID_TABLE_ALL = 1	

Const FORM_ID_DDL_TABLES = "id_ddl_tables"
Const FORM_NAME_DDL_TABLES = "f_ddl_tables"
Const FORM_SIZE_TEXTAREA_COLUMN_MAX = "80"
Const FILE_PATH_CONFIG = "conf"

'//*********************************************************
'//* @function createConnectionStringCsv
'//*********************************************************
Private Function createConnectionStringCsv( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"Provider=Microsoft.Jet.OLEDB.4.0;" & _
		"Data Source=""" & p_strPath & """;" & _
		"Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
	createConnectionStringCsv = strConnStr
End Function

'//*********************************************************
'//* @function createConnectionStringExcel
'//*********************************************************
Private Function createConnectionStringExcel( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & p_strPath & ";" & _
		"Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""
	createConnectionStringExcel = strConnStr
End Function

'//*********************************************************
'//* @function createConnectionStringAccess
'//*********************************************************
Private Function createConnectionStringAccess( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & p_strPath & ";" & _
		"Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""
	createConnectionStringExcel = strConnStr
End Function

'//*********************************************************
'//* @function createConnectionStringSqlServer
'//*********************************************************
Private Function createConnectionStringSqlServer( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & p_strPath & ";" & _
		"Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""
	createConnectionStringExcel = strConnStr
End Function

'//*********************************************************
'//* @function createConnectionStringMySql
'//*********************************************************
Private Function createConnectionStringMySql( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & p_strPath & ";" & _
		"Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""
	createConnectionStringExcel = strConnStr
End Function

'//*********************************************************
'//* @function createConnectionStringPostgreSql
'//*********************************************************
Private Function createConnectionStringPostgreSql( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & p_strPath & ";" & _
		"Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""
	createConnectionStringExcel = strConnStr
End Function

'//*********************************************************
'//* @function createConnectionStringSqlite
'//*********************************************************
Private Function createConnectionStringSqlite( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & p_strPath & ";" & _
		"Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"""
	createConnectionStringExcel = strConnStr
End Function

'//*********************************************************
'//* @function createConnectionStringUdl
'//*********************************************************
Private Function createConnectionStringUdl( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"File name=../conf/" & p_strPath
	createConnectionStringUdl = strConnStr
End Function

'//*********************************************************
'//* @function createConnectionStringFirebird
'//*********************************************************
Private Function createConnectionStringFirebird( _
	byval p_strPath, _
	byval p_strUserId, _
	byval p_strPasswd _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"DRIVER=Firebird/InterBase(r) driver;" & _
		"UID=" & p_strUserId & ";" & _
		"PWD=" & p_strPasswd & ";" & _
		"DBNAME=" & p_strPath & ";"
	createConnectionStringFirebird = strConnStr
End Function

'//*********************************************************
'//* @function getSchemaInfo_Tables
'//*********************************************************
Private Function getSchemaInfo_Tables( _
	byval p_strConnStr _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	objCn.Open p_strConnStr
	
	Set objRs = objCn.OpenSchema(adSchemaTables)
	
	Set getSchemaInfo_Tables = objRs
End Function

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
	
'--	strConnStr = createConnectionStringCsv(p_strFilePath)
	strConnstr = p_strConnStr
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	objCn.Open strConnStr
	strSql = getSqlCommand(SQL_ID_TABLE_ALL)
	
	strSql = Replace(strSql, "%1", p_strTableName)
	
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	strHtml = "<table border=""3"" cellpadding=""3"" cellspacing=""3"">" & vbCrLf
	
	For Each objField In objRs.Fields
		strHtml = strHtml & _
			"<tr>" & vbCrLf
		
		If LenB(objField) > 80 Then
			strHtml = strHtml & _
				"<th>" & objField.Name & "</th>" & vbCrLf & _
				"<td valign=""top""><textarea name=""f_item"" rows=""" & _
				(Round((LenB(objField.Value)+FORM_SIZE_TEXTAREA_COLUMN_MAX)/FORM_SIZE_TEXTAREA_COLUMN_MAX,0)) & _
				""" cols=""" & FORM_SIZE_TEXTAREA_COLUMN_MAX & """></textarea>" & vbCrLf
		Else
			strHtml = strHtml & _
				"<th>" & objField.Name & "</th>" & vbCrLf & _
				"<td><input name=""f_item"" size=""20""></input>" & vbCrLf
		End If
		strHtml = strHtml & _
			"</tr>" & vbCrLf
		
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
'//* @function getSqlCommand
'//* @arg1 sql-id
'//*********************************************************
Private Function getSqlCommand( _
	byval p_intSqlId _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath
	strConnStr = createConnectionStringCsv(g_strDataDir & FILE_PATH_CONFIG)
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	objCn.Open strConnStr
	
	strSql = "select sql_command from sql_commands.csv where id = " & p_intSqlId & ";"
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	getSqlCommand = objRs("sql_command")
	
	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	Set objCn = Nothing
End Function

'//*********************************************************
'//* @function getFilePath_Excel
'//*********************************************************
Private Function getFilePath_Excel( _
	byval p_intId _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath

	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	strConnStr = createConnectionStringCsv(g_strDataDir&"csv")
	
	objCn.Open strConnStr
	
	strSql = Replace(getSqlCommand(SQL_ID_FILE_PATH_EXCEL),"%1", p_intId)

	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	strPath = objRs("path")
	

	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	set objCn = Nothing

	getFilePath_Excel = strPath
End Function

'//*********************************************************
'//* @function getSqlFilter
'//*********************************************************
Private Function getSqlFilter( _
	byval p_intId _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strFilter

	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")
	
	strConnStr = createConnectionStringCsv(g_strDataDir&"conf")
	
	objCn.Open strConnStr
	
	strSql = Replace(getSqlCommand(SQL_ID_SQL_FILTER),"%1", p_intId)
	
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	strFilter = objRs("expression")
	

	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	set objCn = Nothing

	getSqlFilter = strFilter
End Function

'//*********************************************************
'//* @function getSchemaInfo
'//*********************************************************
Private Function getSchemaInfo( _
	byval p_strConnStr, _
	byval p_IntSchemaType _
	)
	'-- var Object
	Dim objCn
	Dim objRs
	
	'-- var String
	Dim strConnStr
	Dim strSql
	Dim strPath
	
	Set objCn = CreateObject("ADODB.Connection")
	Set objRs = CreateObject("ADODB.Recordset")

	objCn.Open p_strConnStr
	
	Set objRs = objCn.OpenSchema(p_IntSchemaType)
	
	Set getSchemaInfo = objRs
End Function

