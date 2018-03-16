Option Explicit

'//**********************************************************
'//* Global Constant Definitions
'//**********************************************************
Const msoPicture = 13

'//**********************************************************
'//* Global var definitions
'//**********************************************************
'-- var Object
Dim g_objFso
Dim g_objExcel
Dim g_strDesktop

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
	
    Dim objShape
    Dim objChart
    Dim i
    Dim strPath
    Dim strNumber
    
    i = 0
    For Each objShape In p_objWks.Shapes
        If objShape.Type = msoPicture Then
            objShape.CopyPicture
            Set objChart = p_objWks.ChartObjects.Add(0, 0, objShape.Width * 0.8, objShape.Height * 0.8).chart
            objChart.Paste
            g_lngIdData = g_lngIdData + 1
            strNumber = g_lngIdData
            strPath = g_strDesktop & "\" & "imgs\" &p_objWks.Name & "_" & String(3 - Len(strNumber), "0") & strNumber & ".png"
            objChart.Export strPath, "PNG", False
            objChart.Parent.Delete
        End If
    Next
End Sub

'//**********************************************************
'//* @procedure main
'//**********************************************************
Private Sub main
	Dim objWkb(3)
	Dim objArg
	Dim objWshShell
	
	'-- var String
	Dim strPrefix

	Set g_objFso = CreateObject("Scripting.FileSystemObject")
	Set g_objExcel = CreateObject("Excel.Application")
	set objWshShell = WScript.CreateObject("WScript.Shell")
    g_strDesktop = objWshShell.SpecialFolders("Desktop")
	g_objExcel.Visible = True
	
	g_lngIdData = 0

	For Each objArg In WScript.Arguments
		Set objWkb(0) = g_objExcel.Workbooks.Open(objArg,,True)
		
		g_objExcel.Application.StatusBar = "bookname=" & objWkb(0).Name
		
		proc_workbook objWkb(0)
		objWkb(0).Close False
	Next

	g_objExcel.Application.StatusBar = ""

	Msgbox "create tables successfully", vbInformation, "処理完了"

	g_objExcel.Quit
End Sub
