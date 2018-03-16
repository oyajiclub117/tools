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
'//* @function createConnectionStringExcel_NOHD
'//*********************************************************
Private Function createConnectionStringExcel_NOHD( _
	byval p_strPath _
	)
	'-- var String
	Dim strConnStr
	
	strConnStr = _
		"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & p_strPath & ";" & _
		"Extended Properties=""Excel 12.0 Xml;HDR=NO;IMEX=1"""
	createConnectionStringExcel_NOHD = strConnStr
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
	createConnectionStringAccess = strConnStr
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
	createConnectionStringSqlServer = strConnStr
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
	createConnectionStringMySql = strConnStr
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
	createConnectionStringPostgreSql = strConnStr
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
'//* @function getSchemaInfo
'//*********************************************************
Private Function getSchemaInfo( _
	byval p_strSchemaName, _
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
	
	Set objRs = objCn.OpenSchema(p_strSchemaName)
	
	Set getSchemaInfo = objRs
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
	
	strPath = objRs("path") & objRs("file")
	
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
	
	strConnStr = createConnectionStringCsv(g_strDataDir & FILE_PATH_CONFIG)
	
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

'//*********************************************************
'//* @function createDDL_FilePath_Excel
'//*********************************************************
Private Function createDDL_FilePath_Excel( _
	byval p_strKey _
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
	
	strSql = Replace(getSqlCommand(SQL_ID_FILE_PATH_ALL_EXCEL), "%1", p_strKey)
	
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	strHtml = _
		"<select name=""f_excel_paths"" onclick=""viewSchema_Tables ''"" style=""display:inline;"">" & vbCrLf & _
		"<option value=""-"">-- file-path --</option>" & vbCrLf
	
	Do While Not objRs.EOF
		strHtml = strHtml & _
			"<option value=""" & objRs("id") & """>" & objRs("title") & "</option>" & vbCrLf
		objRs.MoveNext
	Loop
	
	strHtml = strHtml & _
		"</select>" & vbCrLf

	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	set objCn = Nothing

	createDDL_FilePath_Excel = strHtml
End Function

'//*********************************************************
'//* @function createDDL_SchemaFunc_Excel
'//*********************************************************
Private Function createDDL_SchemaFunc_Excel( _
	byval p_intSchemaId _
	)
Const SQL_ID_SCHEMA_FUNC_EXCEL = 49
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
	
	strSql = Replace(getSqlCommand(SQL_ID_SCHEMA_FUNC_EXCEL), "%1", p_intSchemaId)
	
	objRs.Open strSql, objCn, adOpenStatic, adLockReadOnly
	
	strHtml = _
		"<select name=""f_ddl_func_lists"" onchange=""cmdActionEventExcel_Click"" style=""display:inline;"">" & vbCrLf & _
		"<option value=""-"">-- func-list --</option>" & vbCrLf
	
	Do While Not objRs.EOF
		strHtml = strHtml & _
			"<option value=""" & objRs("sql_id") & "," & objRs("table") & "," & objRs("filter") & """>" & objRs("title") & "</option>" & vbCrLf
		objRs.MoveNext
	Loop
	
	strHtml = strHtml & _
		"</select>" & vbCrLf

	If IsObject(objRs) Then objRs.Close
	If IsObject(objCn) Then objCn.Close
	
	Set objRs = Nothing
	set objCn = Nothing

	createDDL_SchemaFunc_Excel = strHtml
End Function

