Option Explicit

'-- enum Excel
Const xlDown = -4121
Const xlUp = -4162

'-- enum FileSystemObject
Const ForWriting = 2
Const ForAppending = 8

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
		If ( _
			InStr(objWks.Name, "変更履歴") _
		Or InStr(objWks.Name, "エンティティ一覧") _
		Or InStr(objWks.Name, "原紙") _
		Or InStr(objWks.Name, "作業完了") _
		) Then
			'Nothing
		Else
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
	Dim strCategory
	
Const SHEET_START_ROW = "B65535"
Const CELL_COL_MAX = "IV"
Const xlToLeft = -4159
	With p_objWks
	lngLastRow = .Range(SHEET_START_ROW).End(xlUp).Row
	lngLastCol = .Range(CELL_COL_MAX & "2").End(xlToLeft).Column
	
	strColCells = ""
	
	For i = 3 To lngLastCol
		If .Cells(2, i) <> "" Then
			objArrs = Split(.Cells(2, i).AddressLocal, "$")
			strColCells = strColCells & objArrs(1) & ","
		End If
		
	Next

	strColCells  = Left(strColCells, Len(strColCells) - 1)
	End With
	objFieldColCells = Split(strColCells, ",")
	
	lngArrNum = UBound(objFieldColCells)
	
	ReDim strValues(lngArrNum)

	createCsvFile_Head p_objWks, 2, lngLastCol

	strCategory = ""
		
	For i = 3 to lngLastRow
		g_lngIdData = g_lngIdData + 1

		With p_objWks
	
		strCsvData = _
			g_lngIdData & ","
		
		If InStr(.Cells(i, "B"), "インタフェース") _
		Or InStr(.Cells(i, "B"), "帳票") _
		Or InStr(.Cells(i, "B"), "バッチ") Then
			strCategory = Mid(.Cells(i, "B"),2,99)
		Else
		
		For j = 0 to lngArrNum
			if .Cells(i, objFieldColCells(j)) = "" Then .Cells(i, objFieldColCells(j)) = strValues(j)
			If objFieldColCells(j) = "A"  Then
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

		strCsvData = strCsvData & ",""" & strCategory & """"

		createCsvFile_Data strCsvData
		
		End If 


		End With
	Next
End Sub

'//**********************************************************
'//* @procedure createCsvFile_Head
'//**********************************************************
Private Sub createCsvFile_Head( _
	byref p_objWks, _
	byval p_lngStartRow, _
	byval p_lngLastCol _
	)
	'-- var String
	Dim strCsvHead
	'-- var Long
	Dim i
	
	strCsvHead = "id"

	For i = 3 To p_lngLastCol
		strCsvHead = strCsvHead & _
			"," & p_objWks.Cells(p_lngStartRow, i)
	Next
	
	g_objCsvFile.WriteLine strCsvHead & ",category"
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
Private Sub main
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
	g_objExcel.Application.ScreenUpdating = False
	strPath = strDesktop & "\" & "doc_rsrcs.csv"
'	If g_objFso.FileExists(strPath) Then
'		Set g_objCsvFile = g_objFso.OpenTextFile(strPath, ForAppending)
'	Else
		Set g_objCsvFile = g_objFso.CreateTextFile(strPath)
'		createCsvFile_Head
'	End If
	

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
