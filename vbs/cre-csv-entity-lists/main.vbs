Option Explicit

'//**********************************************************
'//* Global var definitions
'//**********************************************************
'-- var Object
Dim g_objFso
Dim g_objCsvFile
Dim g_objExcel

'-- varLong
Dim g_lngIdData

'//**********************************************************
'//* @procedure proc_workbook
'//**********************************************************
Private Sub proc_workbook( _
	byref p_objWkb _
	)
	'-- var Object
	Dim objWks
	
	For Each objWks In p_objWkb.Worksheets
		If (InStr(objWks.Name, "変更履歴") _
		Or InStr(objWks.Name, "bkup")) Then
			'Nothing
		ElseIf (InStr(objWks.Name, "エンティティ一覧")) Then
			g_objExcel.Application.StatusBar = "start sheet name=" & objWks.Name 
			proc_worksheet objWks
			g_objExcel.Application.StatusBar = "ended sheet name=" & objWks.Name 
		End If
	Next
End Sub

'//**********************************************************
'//* @procedure proc_worksheet
'//**********************************************************
Private Sub proc_worksheet( _
	byref p_objWks _
	)
	'-- var Object
	Dim objFieldColCells
	Dim objArrs
	
	'-- var Long
	Dim lngLastRow
	Dim lngLastCol
	Dim i
	Dim j
	Dim lngArrNum
	
	'-- var String
	Dim strCsvData
	Dim strText
	Dim strDelimString
	Dim strColCells
	Dim strValues()
	
	With p_objWks
	lngLastRow = .Range(SHEET_START_COL & SHEET_START_ROW).End(xlDown).Row
	lngLastCol = .Range(CELL_COL_MAX & SHEET_START_ROW).End(xlToLeft).Column
	
	strColCells = ""
	
	For i = 1 To lngLastCol
		If .Cells(SHEET_START_ROW, i) <> "" Then
			objArrs = Split(.Cells(SHEET_START_ROW, i).AddressLocal, "$")
			strColCells = strColCells & objArrs(1) & ","
		End If
		
	Next

	strColCells  = Left(strColCells, Len(strColCells) - 1)
	End With
	objFieldColCells = Split(strColCells, ",")
	
	lngArrNum = UBound(objFieldColCells)
	
	ReDim strValues(lngArrNum)
	
	For i = SHEET_START_ROW + 1 to lngLastRow
		g_lngIdData = g_lngIdData + 1

		With p_objWks
	
		strCsvData = _
			g_lngIdData & "," & _
			"""" & .Range(RANGE_FIX_COL1) & """," & _
			"""" & .Range(RANGE_FIX_COL2) & ""","
		If .Cells(i, "AZ") = "テーブル削除" Then
			'nothing
		Else
		For j = 0 to lngArrNum
			if .Cells(i, objFieldColCells(j)) = "" Then .Cells(i, objFieldColCells(j)) = strValues(j)
			If objFieldColCells(j) = CELL_ID_COL  Then
				strText = .Cells(i, objFieldColCells(j)).value
			Else
				strText = """" & .Cells(i, objFieldColCells(j)).value & """"
			End If
			
			strDelimString = ""
			
			If j < lngArrNum Then strDelimString = ","
			
			strCsvData = strCsvData & _
				strText & strDelimString
			
			if .Cells(i, objFieldColCells(j)) <> "" Then strValues(j) = .Cells(i, objFieldColCells(j))
		Next

		createCsvFile_Data strCsvData
		End If
		End With
	Next
End Sub

'//**********************************************************
'//* @procedure createCsvFile_Head
'//**********************************************************
Private Sub createCsvFile_Head
	'-- var String
	Dim strCsvHead
	
	strCsvHead = _
		"id," & _
		"func_group_id," & _
		"func_group_name," & _
		"list_id," & _
		"l_entity_name," & _
		"p_entity_name," & _
		"data_type," & _
		"description," & _
		"retention_period," & _
		"required"
	g_objCsvFile.WriteLine strCsvHead
End Sub

'//**********************************************************
'//* @procedure createCsvFile_Data
'//**********************************************************
Private Sub createCsvFile_Data( _
	byval p_strData _
	)
	g_objCsvFile.WriteLine p_strData
End Sub

'//**********************************************************
'//* @procedure main
'//**********************************************************
Private Sub main( _
	byval p_strDestFile _
	)
	Dim objWkb(3)
	Dim objArg
	Dim objWshShell
	
	'-- var String
	Dim strDesktop
	Dim strPath

	Set g_objFso = CreateObject("Scripting.FileSystemObject")
	Set g_objExcel = CreateObject("Excel.Application")
	set objWshShell = WScript.CreateObject("WScript.Shell")
    strDesktop = objWshShell.SpecialFolders("Desktop")
	g_objExcel.Visible = True
	strPath = strDesktop & "\" & p_strDestFile
	If g_objFso.FileExists(strPath) Then
		Set g_objCsvFile = g_objFso.OpenTextFile(strPath, ForAppending)
	Else
		Set g_objCsvFile = g_objFso.CreateTextFile(strPath)
		createCsvFile_Head
	End If
	

	g_lngIdData = 0

	For Each objArg In WScript.Arguments
		Set objWkb(0) = g_objExcel.Workbooks.Open(objArg,,True)
		
		g_objExcel.Application.StatusBar = "bookname=" & objWkb(0).Name
		
		proc_workbook objWkb(0)
		objWkb(0).Close False
	Next

	g_objExcel.Application.StatusBar = ""

	Msgbox "create tables successfully", vbInformation, "処理完了"

	g_objCsvFile.Close

	g_objExcel.Quit
End Sub
