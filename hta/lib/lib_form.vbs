Option Explicit

'//*********************************************************
'//* @procedure setDisplay
'//* @arg1 [form-id]
'//*********************************************************
Private Sub setDisplay( _
	byref p_objForm _
	)
	If p_objForm.style.display = "none" Then
		p_objForm.style.display = "inline"
	Else
		p_objForm.style.display = "none"
	End If
End Sub
