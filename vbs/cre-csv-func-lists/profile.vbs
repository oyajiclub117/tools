Option Explicit

'-- enum Excel
Const xlDown = -4121
Const xlUp = -4162
Const xlToLeft = -4159

'-- enum FileSystemObject
Const ForWriting = 2
Const ForAppending = 8

Const CELL_COL_MAX = "XFD"
Const CELL_ROW_MAX = 1048576

'//**********************************************************
'//* @procedure setConfigCell
'//**********************************************************
Private Sub setConfigCell( _
	byref p_objWkb _
	)
	COL_START_TARTGET = 8
	If InStr(p_objWkb.Name, "���X�g�E�I�[�_�[����") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "B"
		START_ROW_DATA = 8
		CELL_ID_COL = "B"
		CSV_ADD_COLS = "R2,R3"
		COL_START_TARTGET = 2
	ElseIf InStr(p_objWkb.Name, "���UDM") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "H"
		START_ROW_DATA = 8
		CELL_ID_COL = "H"
		CSV_ADD_COLS = "R2,R3"
	ElseIf InStr(p_objWkb.Name, "���i�E���ޔ���") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "A"
		START_ROW_DATA = 8
		CELL_ID_COL = "A"
		CSV_ADD_COLS = "L2,L3"
		COL_START_TARTGET = 1
	ElseIf InStr(p_objWkb.Name, "��������") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "H"
		START_ROW_DATA = 8
		CELL_ID_COL = "H"
		CSV_ADD_COLS = "R2,R3"
	ElseIf InStr(p_objWkb.Name, "�ύX�f�[�^") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "H"
		START_ROW_DATA = 8
		CELL_ID_COL = "H"
		CSV_ADD_COLS = "R2,R3"
	ElseIf InStr(p_objWkb.Name, "�T�u�V�X�e��") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "A"
		START_ROW_DATA = 8
		CELL_ID_COL = "A"
		CSV_ADD_COLS = "Q2,Q3"
		COL_START_TARTGET = 1
	ElseIf InStr(p_objWkb.Name, "��V�X�e������") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "B"
		START_ROW_DATA = 8
		CELL_ID_COL = "B"
		CSV_ADD_COLS = "R2,R3"
		COL_START_TARTGET = 2
	End If
End Sub
