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

Const HEAD_START_ROW = 1
Const HEAD_START_COL = 1
Const CELL_DBSET_START = "B1"

'//**********************************************************
'//* @procedure setConfigCell
'//**********************************************************
Private Sub setConfigCell( _
	byref p_objWkb _
	)
	COL_START_TARTGET = 2
	SHEET_START_ROW = "5"
	SHEET_START_COL = "Y"
	START_ROW_DATA = 1
	CELL_ID_COL = "B"
	CSV_ADD_COLS = "L2,L3,S2,S3,AD2,AD3,AK2,Ak3"
End Sub
