Attribute VB_Name = "ControlTrueFalse"
'-----------------------------------------------------------------------
' activa/inactiva los controles del form
'-----------------------------------------------------------------------
Sub fSetEnableFields(xForm As Form, lEnable As Boolean)

  Dim oform As Form, oCtrl As Control
  Set oform = xForm
  
  With oform
        For Each oCtrl In .Controls
            Select Case UCase(TypeName(oCtrl))
                   
                   Case Is = "TEXTBOX"
                        oCtrl.Enabled = lEnable
                   
                   Case Is = "CHECKBOX"
                        oCtrl.Enabled = lEnable
                   
                   Case Is = "COMMANDBUTTON"
                   
                   Case Is = "DATACOMBO"
                        oCtrl.Enabled = lEnable
                   
                   Case Is = "DTPICKER"
                        oCtrl.Enabled = lEnable
                    
                    Case Is = "RICHTEXTBOX"
                        oCtrl.Enabled = lEnable
                    
                    Case Is = "OPTIONBUTTON"
                        oCtrl.Enabled = lEnable
                    
                    Case Is = "COMBOBOX"
                        oCtrl.Enabled = lEnable
                    
                    Case Is = "MASKEDBOX"
                        oCtrl.Enabled = lEnable
'----> control para toolbar ----------
            '           If oCtrl.Name = "cmdGrabar" Or _
            '              oCtrl.Name = "cmdCancelar" Then
            '              oCtrl.Visible = lEnable
            '           Else
            '              oCtrl.Visible = Not lEnable
            '           End If

            End Select
        Next
  End With
  
  Set oform = Nothing
  Set oCtrl = Nothing

End Sub


