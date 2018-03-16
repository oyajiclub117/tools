Option Explicit

'-- enum Excel
Const xlDown = -4121
Const xlUp = -4162

'-- enum FileSystemObject
Const ForReading = 1
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
		If InStr(objWks.Name, "�ύX����") _
		Or InStr(objWks.Name, "index") _
		Or InStr(objWks.Name, "LIST") Then
		'-- Nothing
		Else
		'	nothing
'--		If InStr(objWks.Name, "�����蒠") _
'--		Or InStr(objWks.Name,"���ޔ��ʒm�͂����Q�����Ҍ��C�p�Q�`���f�[�^") _
'--		Or InStr(objWks.Name,"�Q���؁Q�����Ҍ��C�p�Q�`���f�[�^") _
'--		Or InStr(objWks.Name,"�Q���Җ���Q�����Ҍ��C�p�Q�`���f�[�^") _
'--		Or InStr(objWks.Name,"�Q���؁Q�����Ҍ��C�đ��p�Q�`���f�[�^") Then
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
	
	'-- var Long
	Dim i
	Dim lngLastRow
	'-- var String
	Dim strCsvData
	
	lngLastRow = p_objWks.Range("B8").End(xlDown).Row

	For i = 9 To lngLastRow
		If p_objWks.Cells(i, "D") <> "" Then
			g_lngIdData = g_lngIdData + 1
			
			strCsvData = _
				g_lngIdData & "," & _
				"""" & p_objWks.Name & """,""" & p_objWks.Cells(i, "D") & """"
			createCsvFile_Data strCsvData
		End If
	Next

End Sub

'//**********************************************************
'//* @procedure createCsvFile_Head
'//**********************************************************
Private Sub createCsvFile_Head
	'-- var String
	Dim strCsvHead
	
	strCsvHead = _
		"id" & "," & _
		"table" & "," & _
		"column"
	
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
	byval p_strFileName _
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
	strPath = strDesktop & "\" & p_strFileName

	g_lngIdData = 0

'--	If g_objFso.FileExists(strPath) Then
'--		Set g_objCsvFile = g_objFso.OpenTextFile(strPath, ForReading )
'--		Do Until  g_objCsvFile.AtEndOfLine
'--			g_objCsvFile.ReadLine
'--			g_lngIdData = g_lngIdData + 1
'--		Loop
'--
'--		g_lngIdData = g_lngIdData - 1
'--
'--		g_objCsvFile.Close
'--		
'--		Set g_objCsvFile = g_objFso.OpenTextFile(strPath, ForAppending)
'--	Else
'--		Set g_objCsvFile = g_objFso.CreateTextFile(strPath)
		Set g_objCsvFile = g_objFso.CreateTextFile(strPath)
		createCsvFile_Head
'--	End If
	

	For Each objArg In WScript.Arguments
'--		Set objWkb(0) = g_objExcel.Workbooks.Open(objArg,,False)
		Set objWkb(0) = g_objExcel.Workbooks.Open(objArg,,true)
		
		g_objExcel.Application.StatusBar = "bookname=" & objWkb(0).Name
		
		proc_workbook objWkb(0)
'--		objWkb(0).Close true
		objWkb(0).Close false
	Next

	g_objExcel.Application.StatusBar = ""

	Msgbox "script execute successfully", vbInformation, "��������"

	g_objCsvFile.Close

	g_objExcel.Quit
End Sub
