Option Explicit

'//*********************************************************
'//* @procedure showWindow_ModalDialog
'//*********************************************************
Private Sub showWindow_ModalDialog( _
	byval p_strUrl, _
	byref p_objArgs, _
	byval p_intWidth, _
	byval p_intHeight _
	)
	'-- var Object
	Dim strFeatures
	
	strFeatures = "dialogWidth:" & p_intWidth & "px;" & _
		"dialogHeight:" & p_intHeight & "px"
	
	window.showModalDialog _
		p_strUrl, _
		p_objArgs, _
		strFeatures
	
End Sub
