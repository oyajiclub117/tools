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
	If InStr(p_objWkb.Name, "リスト・オーダー共通") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "B"
		START_ROW_DATA = 8
		CELL_ID_COL = "B"
		CSV_ADD_COLS = "R2,R3"
		COL_START_TARTGET = 2
	ElseIf InStr(p_objWkb.Name, "勧誘DM") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "H"
		START_ROW_DATA = 8
		CELL_ID_COL = "H"
		CSV_ADD_COLS = "R2,R3"
	ElseIf InStr(p_objWkb.Name, "商品・教材発送") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "A"
		START_ROW_DATA = 8
		CELL_ID_COL = "A"
		CSV_ADD_COLS = "L2,L3"
		COL_START_TARTGET = 1
	ElseIf InStr(p_objWkb.Name, "請求入金") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "H"
		START_ROW_DATA = 8
		CELL_ID_COL = "H"
		CSV_ADD_COLS = "R2,R3"
	ElseIf InStr(p_objWkb.Name, "変更データ") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "H"
		START_ROW_DATA = 8
		CELL_ID_COL = "H"
		CSV_ADD_COLS = "R2,R3"
	ElseIf InStr(p_objWkb.Name, "サブシステム") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "A"
		START_ROW_DATA = 8
		CELL_ID_COL = "A"
		CSV_ADD_COLS = "Q2,Q3"
		COL_START_TARTGET = 1
	ElseIf InStr(p_objWkb.Name, "基幹システム共通") Then
		SHEET_START_ROW = "5"
		SHEET_START_COL = "B"
		START_ROW_DATA = 8
		CELL_ID_COL = "B"
		CSV_ADD_COLS = "R2,R3"
		COL_START_TARTGET = 2
	End If
End Sub
