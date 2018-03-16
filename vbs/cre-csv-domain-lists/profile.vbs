Option Explicit

'-- enum Excel
Const xlDown = -4121
Const xlUp = -4162
Const xlToLeft = -4159

'-- enum FileSystemObject
Const ForWriting = 2
Const ForAppending = 8

Const CELL_COL_MAX = "XFD"

'//**********************************************************
'//* @procedure setConfigCell
'//**********************************************************
Private Sub setConfigCell( _
	byref p_objWkb _
	)
	COL_START_TARTGET = 1
	SHEET_START_ROW = "4"
	SHEET_START_COL = "A"
	START_ROW_DATA = 2
	CELL_ID_COL = "A"
	CSV_ADD_COLS = ""
	COL_START_TARTGET = 1
End Sub
