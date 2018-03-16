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
	
	'-- var Long
	Dim lngLastRow
	Dim i
	Dim j
	Dim lngArrNum
	
	'-- var String
	Dim strCsvData
	Dim strText
	Dim strDelimString
	
	lngLastRow = p_objWks.Range("B8").End(xlDown).Row
	
	objFieldColCells = Array("B","D", "L", "T", "X", "Z", "AB", "AD", "AF", "AH", "AK")
	
	lngArrNum = UBound(objFieldColCells)
	
	For i = 9 to lngLastRow
		g_lngIdData = g_lngIdData + 1

		With p_objWks
	
		strCsvData = _
			g_lngIdData & "," & _
			.Range("L3") & "," & _
			.Range("L6") & "," & _
			.Range("L5") & ","
		
		For j = 0 to lngArrNum
			If objFieldColCells(j) = "B" Then
				strText = .Cells(i, objFieldColCells(j))
			Else
				strText = """" & .Cells(i, objFieldColCells(j)) & """"
			End If
			
			strDelimString = ""
			
			If j < lngArrNum Then strDelimString = ","
			
			strCsvData = strCsvData & _
				strText & strDelimString
			
		Next

		createCsvFile_Data strCsvData

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
		"subsystem_name," & _
		"table_name," & _
		"table_name_j," & _
		"column_id," & _
		"column_name_j," & _
		"column_name," & _
		"data_type," & _
		"data_length," & _
		"data_precision," & _
		"is_null," & _
		"is_uniquekey," & _
		"is_primarykey," & _
		"default_value," & _
		"comment"
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
Private Sub main
	Dim objWkb(3)
	Dim objArg
	Dim objWshShell
	
	'-- var String
	Dim strDesktop
	Dim strPrefix

	Set g_objFso = CreateObject("Scripting.FileSystemObject")
	Set g_objExcel = CreateObject("Excel.Application")
	set objWshShell = WScript.CreateObject("WScript.Shell")
    strDesktop = objWshShell.SpecialFolders("Desktop")
	g_objExcel.Visible = True
	Set g_objCsvFile = g_objFso.CreateTextFile(strDesktop & "\bus_tab_columns.csv")
	
	createCsvFile_Head

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
