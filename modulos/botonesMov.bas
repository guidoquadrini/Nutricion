Attribute VB_Name = "botonesMov"
Sub enabledDesplaz() 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados

    If MDIForm1.ActiveForm.Data1.Recordset.RecordCount = 0 Or MDIForm1.ActiveForm.Data1.Recordset.RecordCount = 1 Then
            
            MDIForm1.ActiveForm.cmdSiguiente.Enabled = False
            MDIForm1.ActiveForm.cmdUltimo.Enabled = False
            
            MDIForm1.ActiveForm.Pic_Siguiente_Gris.ZOrder 0
            MDIForm1.ActiveForm.Pic_Ultimo_Gris.ZOrder 0
            
            MDIForm1.ActiveForm.cmdAnterior.Enabled = False
            MDIForm1.ActiveForm.cmdPrimero.Enabled = False
            
            MDIForm1.ActiveForm.Pic_Anterior_Gris.ZOrder 0
            MDIForm1.ActiveForm.Pic_Primero_Gris.ZOrder 0
            
            If MDIForm1.ActiveForm.Data1.Recordset.RecordCount = 1 Then
                MDIForm1.ActiveForm.cmdModificar.Enabled = True
                MDIForm1.ActiveForm.cmdBorrar.Enabled = True
                
                MDIForm1.ActiveForm.Pic_Modificar.ZOrder 0
                MDIForm1.ActiveForm.Pic_Borrar.ZOrder 0
            
            Else
                MDIForm1.ActiveForm.cmdModificar.Enabled = False
                MDIForm1.ActiveForm.cmdBorrar.Enabled = False
                
                MDIForm1.ActiveForm.Pic_Modificar_Gris.ZOrder 0
                MDIForm1.ActiveForm.Pic_Borrar_Gris.ZOrder 0
            End If
    Else
        
        MDIForm1.ActiveForm.cmdModificar.Enabled = True
        MDIForm1.ActiveForm.cmdBorrar.Enabled = True
        
        MDIForm1.ActiveForm.Pic_Modificar.ZOrder 0
        MDIForm1.ActiveForm.Pic_Borrar.ZOrder 0
                
        If MDIForm1.ActiveForm.Data1.Recordset.AbsolutePosition = 0 Then
            
            MDIForm1.ActiveForm.cmdSiguiente.Enabled = True
            MDIForm1.ActiveForm.cmdUltimo.Enabled = True
            
            MDIForm1.ActiveForm.Pic_Siguiente.ZOrder 0
            MDIForm1.ActiveForm.Pic_Ultimo.ZOrder 0
                
            MDIForm1.ActiveForm.cmdAnterior.Enabled = False
            MDIForm1.ActiveForm.cmdPrimero.Enabled = False
        
            MDIForm1.ActiveForm.Pic_Anterior_Gris.ZOrder 0
            MDIForm1.ActiveForm.Pic_Primero_Gris.ZOrder 0
        
        Else
            If MDIForm1.ActiveForm.Data1.Recordset.AbsolutePosition = MDIForm1.ActiveForm.Data1.Recordset.RecordCount - 1 Then
                
                MDIForm1.ActiveForm.cmdSiguiente.Enabled = False
                MDIForm1.ActiveForm.cmdUltimo.Enabled = False
                
                MDIForm1.ActiveForm.Pic_Siguiente_Gris.ZOrder 0
                MDIForm1.ActiveForm.Pic_Ultimo_Gris.ZOrder 0
                
                MDIForm1.ActiveForm.cmdAnterior.Enabled = True
                MDIForm1.ActiveForm.cmdPrimero.Enabled = True
            
                MDIForm1.ActiveForm.Pic_Anterior.ZOrder 0
                MDIForm1.ActiveForm.Pic_Primero.ZOrder 0
                
            Else
                    
                    MDIForm1.ActiveForm.cmdSiguiente.Enabled = True
                    MDIForm1.ActiveForm.cmdUltimo.Enabled = True
                    
                    MDIForm1.ActiveForm.Pic_Siguiente.ZOrder 0
                    MDIForm1.ActiveForm.Pic_Ultimo.ZOrder 0
                
                    MDIForm1.ActiveForm.cmdAnterior.Enabled = True
                    MDIForm1.ActiveForm.cmdPrimero.Enabled = True
                    
                    MDIForm1.ActiveForm.Pic_Anterior.ZOrder 0
                    MDIForm1.ActiveForm.Pic_Primero.ZOrder 0
                    
            End If
            
        End If
        
    End If
            
End Sub

