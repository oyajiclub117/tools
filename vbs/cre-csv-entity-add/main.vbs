Option Explicit

'//**********************************************************
'//* Global var definitions
'//**********************************************************
'-- var Object
Dim g_objFso
Dim g_objCsvFile
Dim g_objExcel
Dim g_objWkbTmpl
Dim g_objRang

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
	Dim objWkb
	
	For Each objWks In p_objWkb.Worksheets
		If (InStr(objWks.Name, "変更履歴") _
		Or InStr(objWks.Name, "bkup") _
		Or objWks.Name = "LIST") Then
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
	
	With p_objWks
		lngLastRow = .Range(SHEET_START_COL & SHEET_START_ROW).End(xlDown).Row
		lngLastCol = .Range(CELL_COL_MAX & SHEET_START_ROW).End(xlToLeft).Column
		
		strColCells = ""
		
'--		For i = 1 To lngLastCol
'--			If .Cells(SHEET_START_ROW, i) <> "" Then
'--				objArrs = Split(.Cells(SHEET_START_ROW, i).AddressLocal, "$")
'--				strColCells = strColCells & objArrs(1) & ","
'--			End If
'--			
'--		Next
'--
'--		strColCells  = Left(strColCells, Len(strColCells) - 1)
		j = 0
		For i = SHEET_START_ROW To lngLastRow
			If .Cells(i, "D") = "作成日時" _
			Or .Cells(i, "D") = "作成ユーザＩＤ" _
			Or .Cells(i, "D") = "更新日時" _
			Or .Cells(i, "D") = "更新ユーザＩＤ" _
			Or .Cells(i, "D") = "削除フラグ" Then
				j = j + 1
			End If
		Next

		If j = 0 Then
			g_objRang.Copy
			
			.Rows(lngLastRow + 1 & ":" & lngLastRow + 1).Insert xlDown
			
			g_objExcel.Application.CutCopyMode = False
		End If
	End With

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
	byval p_strDestFile, _
	byval p_strTmplFile _
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
'--	If g_objFso.FileExists(strPath) Then
'--		Set g_objCsvFile = g_objFso.OpenTextFile(strPath, ForAppending)
'--	Else
'--		Set g_objCsvFile = g_objFso.CreateTextFile(strPath)
'--		createCsvFile_Head
'--	End If
	

	g_lngIdData = 0

	Set g_objWkbTmpl = g_objExcel.Workbooks.Open(p_strTmplFile,,False)

	
	Set g_objRang = g_objWkbTmpl.Worksheets("資料請求テーブル＿ヘッダ").Rows("47:51")

	For Each objArg In WScript.Arguments
		Set objWkb(0) = g_objExcel.Workbooks.Open(objArg,,False)
		
		g_objExcel.Application.StatusBar = "bookname=" & objWkb(0).Name
		
		proc_workbook objWkb(0)
'--		objWkb(0).Close True
	Next

	g_objExcel.Application.StatusBar = ""

	Msgbox "create tables successfully", vbInformation, "処理完了"

'--	g_objCsvFile.Close
	g_objWkbTmpl.Close

'--	g_objExcel.Quit
End Sub
