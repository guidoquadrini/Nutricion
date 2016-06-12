VERSION 5.00
Begin VB.Form o_botenes_abm_back 
   Caption         =   "Form1"
   ClientHeight    =   735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   735
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.PictureBox Pic_Cancelar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5160
         MouseIcon       =   "clase_btns_abm_back.frx":0000
         Picture         =   "clase_btns_abm_back.frx":0152
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   21
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Aceptar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4680
         MouseIcon       =   "clase_btns_abm_back.frx":0453
         Picture         =   "clase_btns_abm_back.frx":05A5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   20
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Modificar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         MouseIcon       =   "clase_btns_abm_back.frx":0861
         Picture         =   "clase_btns_abm_back.frx":09B3
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   19
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Borrar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3600
         MouseIcon       =   "clase_btns_abm_back.frx":0C26
         Picture         =   "clase_btns_abm_back.frx":0D78
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   18
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Agregar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         MouseIcon       =   "clase_btns_abm_back.frx":0F07
         Picture         =   "clase_btns_abm_back.frx":1059
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   17
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Ultimo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         MouseIcon       =   "clase_btns_abm_back.frx":1493
         Picture         =   "clase_btns_abm_back.frx":15E5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   16
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Siguiente 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2040
         MouseIcon       =   "clase_btns_abm_back.frx":1AB0
         Picture         =   "clase_btns_abm_back.frx":1C02
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   15
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Buscar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         MouseIcon       =   "clase_btns_abm_back.frx":2077
         Picture         =   "clase_btns_abm_back.frx":21C9
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   14
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Anterior 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         MouseIcon       =   "clase_btns_abm_back.frx":24A6
         Picture         =   "clase_btns_abm_back.frx":25F8
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   13
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Primero 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         MouseIcon       =   "clase_btns_abm_back.frx":2A66
         Picture         =   "clase_btns_abm_back.frx":2BB8
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Primero_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         MouseIcon       =   "clase_btns_abm_back.frx":3073
         Picture         =   "clase_btns_abm_back.frx":31C5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   31
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Anterior_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1080
         MouseIcon       =   "clase_btns_abm_back.frx":336A
         Picture         =   "clase_btns_abm_back.frx":34BC
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   30
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Buscar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1560
         MouseIcon       =   "clase_btns_abm_back.frx":3626
         Picture         =   "clase_btns_abm_back.frx":3778
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   29
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Siguiente_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2040
         MouseIcon       =   "clase_btns_abm_back.frx":3900
         Picture         =   "clase_btns_abm_back.frx":3A52
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   28
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Ultimo_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2520
         MouseIcon       =   "clase_btns_abm_back.frx":3BBF
         Picture         =   "clase_btns_abm_back.frx":3D11
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   27
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Agregar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3120
         MouseIcon       =   "clase_btns_abm_back.frx":3EB9
         Picture         =   "clase_btns_abm_back.frx":400B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   26
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Borrar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3600
         MouseIcon       =   "clase_btns_abm_back.frx":4158
         Picture         =   "clase_btns_abm_back.frx":42AA
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   25
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Modificar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         MouseIcon       =   "clase_btns_abm_back.frx":43C0
         Picture         =   "clase_btns_abm_back.frx":4512
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   24
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Aceptar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4680
         MouseIcon       =   "clase_btns_abm_back.frx":4633
         Picture         =   "clase_btns_abm_back.frx":4785
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   23
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Cancelar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5160
         MouseIcon       =   "clase_btns_abm_back.frx":48DE
         Picture         =   "clase_btns_abm_back.frx":4A30
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   22
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":4BC4
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "clase_btns_abm_back.frx":52D4
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":59D8
         Height          =   375
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":5B6C
         Picture         =   "clase_btns_abm_back.frx":5CBE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancelar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":6171
         Height          =   375
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":62CA
         Picture         =   "clase_btns_abm_back.frx":641C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Aceptar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":66D8
         Height          =   375
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":67F9
         Picture         =   "clase_btns_abm_back.frx":694B
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Modificar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrar 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":6BBE
         Height          =   375
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":6CD4
         Picture         =   "clase_btns_abm_back.frx":6E26
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Eliminar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":6FB5
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":7102
         Picture         =   "clase_btns_abm_back.frx":7254
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Agregar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUltimo 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":768E
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":7836
         Picture         =   "clase_btns_abm_back.frx":7988
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ultimo"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSiguiente 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":7E53
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":7FC0
         Picture         =   "clase_btns_abm_back.frx":8112
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Siguiente"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":8587
         Height          =   375
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":870F
         Picture         =   "clase_btns_abm_back.frx":8861
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Buscar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":8B3E
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":8CA8
         Picture         =   "clase_btns_abm_back.frx":8DFA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Anterior"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimero 
         Appearance      =   0  'Flat
         DisabledPicture =   "clase_btns_abm_back.frx":9268
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "clase_btns_abm_back.frx":940D
         Picture         =   "clase_btns_abm_back.frx":955F
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Primero"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
   End
End
Attribute VB_Name = "o_botenes_abm_back"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public estadoAbm As Integer ' define el estado de un formulario de abm
'                             1 = sin cambios; 2 = agregar; 3 = modificar
'el modulo "fSetEnableFields(MDIForm1.ActiveForm, vbFalse)" se debe agregar al proyecto

Private Sub cmdAceptar_Click()

If estadoAbm = 2 Or estadoAbm = 3 Then 'si el estado es agregar o modificar

    MDIForm1.ActiveForm.Data1.UpdateRecord
    MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
    
    'condiciones extras
        'If estadoAbm = 2 Then
        '    dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(MDIForm1.ActiveForm.Label1.Caption) & ", codalimento from alimentos where estado = true"
        'End If
        
    cmdBuscar.Enabled = True
    cmdAgregar.Enabled = True
    cmdBorrar.Enabled = True
    'cmdClose.Enabled = True
    cmdModificar.Enabled = True
    
    cmdAgregar.SetFocus
    cmdAgregar.Default = True
    cmdCancelar.Cancel = True
    
    cmdPrimero.Enabled = True
    cmdAnterior.Enabled = True
    cmdSiguiente.Enabled = True
    cmdUltimo.Enabled = True
   
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
    
Else

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        MDIForm1.ActiveForm.Hide
    
    End If
    
End If

End Sub

Private Sub cmdAgregar_Click()

Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

estadoAbm = 2 ' el estado es agregar

MDIForm1.ActiveForm.Data1.Recordset.AddNew

cmdAgregar.Enabled = False
cmdBorrar.Enabled = False
'cmdclose.Enabled = False
cmdModificar.Enabled = False
cmdBuscar.Enabled = False
cmdAceptar.Visible = True
cmdCancelar.Visible = True
cmdPrimero.Enabled = False
cmdAnterior.Enabled = False
cmdSiguiente.Enabled = False
cmdUltimo.Enabled = False

MDIForm1.ActiveForm.txtFields(1).SetFocus

Unload PrincipalFrm
Unload frm_formulaDesarrollada
Unload Form1

cmdAceptar.Default = True
cmdCancelar.Cancel = True

End Sub

Private Sub cmdAnterior_Click()
'If MDIForm1.ActiveForm.Data1.Recordset.BOF = False Then
    MDIForm1.ActiveForm.Data1.Recordset.MovePrevious
'Else
'    MDIForm1.ActiveForm.Data1.Recordset.MoveLast
'End If

If MDIForm1.ActiveForm.Data1.Recordset.AbsolutePosition = 0 Then

    cmdAnterior.Enabled = False
    cmdPrimero.Enabled = False
    
Else
    
    cmdSiguiente.Enabled = True
    cmdUltimo.Enabled = True

End If

End Sub

Private Sub cmdBorrar_Click()
  'esto puede producir un error si elimina el último
  'registro o el único registro del recordset

If MDIForm1.ActiveForm.Data1.Recordset.RecordCount > 0 And MDIForm1.ActiveForm.Data1.Recordset.EOF = False And MDIForm1.ActiveForm.Data1.Recordset.BOF = False Then
    msg = MsgBox("¿Desea Eliminar el registro actual?", vbYesNo, "Eliminar")
    
    If msg = vbYes Then
        'verifica que se pueda eliminar sin problemas y no perder integridad
        
            'strquery = "select * from alimenxpaciente where legajo = " & Val(Label1.Caption) & " and cantidad <> 0"
                    
            'Set MDIForm1.ActiveForm.tb = dbdiet.OpenRecordset(strquery)
            'strquery = "select * from menu where legajo = " & Val(Label1.Caption)
            
            'Set tb1 = dbdiet.OpenRecordset(strquery)
            'If tb.RecordCount = 0 And tb1.RecordCount = 0 Then
                Data1.Recordset.Delete
                Data1.Recordset.MovePrevious
            '    dbdiet.Execute "delete from alimenxpaciente where legajo = " & Val(Label1.Caption)
            '    dbdiet.Execute "delete from menu where legajo = " & Val(Label1.Caption)
            '    dbdiet.Execute "delete from platosmenu where legajo = " & Val(Label1.Caption)
            'Else
            '    MsgBox "No se puede eliminar '" & txtFields(1).Text & "' porque puede afectar la integridad del Sistema", , "Información"
            'End If
            'tb.Close
            'tb1.Close
        
    Else
        cmdAgregar.SetFocus
    End If
End If

End Sub

Private Sub cmdBuscar_Click()
Dim strquery As String
'aclare campo por el cual buscar
    'msg = InputBox("Ingrese apellido del paciente:", "Buscar por Apellido")
    
    'strquery = " select * from pacientes where apell like '" & msg & "*' order by apell, nombre"

With MDIForm1.ActiveForm.Data1
    .RecordSource = strquery
    .Refresh
End With

End Sub

Private Sub cmdCancelar_Click()
If estadoAbm = 2 Or estadoAbm = 3 Then ' el estado del form es agregar o modificar

    MDIForm1.ActiveForm.Data1.Recordset.CancelUpdate
    
    
    cmdBuscar.Enabled = True
    cmdAgregar.Enabled = True
    cmdBorrar.Enabled = True
    'cmdClose.Enabled = True
    cmdModificar.Enabled = True
    
    cmdAgregar.SetFocus
    cmdAgregar.Default = True
    'cmdClose.Cancel = True
    cmdPrimero.Enabled = True
    cmdAnterior.Enabled = True
    cmdSiguiente.Enabled = True
    cmdUltimo.Enabled = True
           
    
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
    
Else

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        MDIForm1.ActiveForm.Hide
    
    End If

End If
End Sub



Private Sub cmdImprimir_Click()
'aclare el filtro para imprimir
    'CrystalReport1.SelectionFormula = " {pacientes.legajo} = " & Val(Label1.Caption) '& " and {platosmenu.fechaMenu} in Date(" & Year(DTdesde.Value) & ", " & Month(DTdesde.Value) & ", " & Day(DTdesde.Value) & ") to Date(" & Year(DThasta.Value) & ", " & Month(DThasta.Value) & ", " & Day(DThasta.Value) & ") "

CrystalReport1.Destination = crptToWindow
CrystalReport1.PrintReport

End Sub

Private Sub cmdModificar_Click()

Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

If MDIForm1.ActiveForm.Data1.Recordset.BOF = True Or MDIForm1.ActiveForm.Data1.Recordset.EOF = True Then
    MDIForm1.ActiveForm.Data1.Recordset.MoveFirst
End If

cmdAgregar.Enabled = False
cmdBorrar.Enabled = False
'cmdclose.Enabled = False
cmdModificar.Enabled = False
cmdBuscar.Enabled = False
cmdAceptar.Visible = True
cmdCancelar.Visible = True
cmdPrimero.Enabled = False
cmdAnterior.Enabled = False
cmdSiguiente.Enabled = False
cmdUltimo.Enabled = False


MDIForm1.ActiveForm.Data1.Recordset.Edit
MDIForm1.ActiveForm.txtFields(1).SetFocus

cmdAceptar.Default = True
cmdCancelar.Cancel = True

estadoAbm = 3 ' el estado es modificar

End Sub

Private Sub cmdPrimero_Click()

MDIForm1.ActiveForm.Data1.Recordset.MoveFirst

cmdSiguiente.Enabled = True
cmdUltimo.Enabled = True

cmdAnterior.Enabled = False
cmdPrimero.Enabled = False

End Sub

Private Sub cmdSiguiente_Click()
'If MDIForm1.ActiveForm.Data1.Recordset.EOF = False Then
    MDIForm1.ActiveForm.Data1.Recordset.MoveNext
'Else
'    MDIForm1.ActiveForm.Data1.Recordset.MoveFirst
'End If

If Data1.Recordset.AbsolutePosition = Data1.Recordset.RecordCount - 1 Then

    cmdSiguiente.Enabled = False
    cmdUltimo.Enabled = False
    
Else

    cmdAnterior.Enabled = True
    cmdPrimero.Enabled = True
     
End If

End Sub

Private Sub cmdUltimo_Click()

MDIForm1.ActiveForm.Data1.Recordset.MoveLast

cmdSiguiente.Enabled = False
cmdUltimo.Enabled = False

cmdAnterior.Enabled = True
cmdPrimero.Enabled = True

End Sub

Private Sub Form_Paint()

'If MDIForm1.ActiveForm.Data1.Recordset.AbsolutePosition = 0 Then
'
'    cmdSiguiente.Enabled = True
'    cmdUltimo.Enabled = True
'
'    cmdAnterior.Enabled = False
'    cmdPrimero.Enabled = False
'
'Else
'    If MDIForm1.ActiveForm.Data1.Recordset.AbsolutePosition = MDIForm1.ActiveForm.Data1.Recordset.RecordCount - 1 Then
'
'        cmdSiguiente.Enabled = False
'        cmdUltimo.Enabled = False
'
'        cmdAnterior.Enabled = True
'        cmdPrimero.Enabled = True
'
'    End If
'End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call f_Boton_Zorder

End Sub

Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdAceptar.ZOrder 0

Me.cmdPrimero.ZOrder 1
Me.cmdAnterior.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Agregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdAgregar.ZOrder 0

Me.cmdPrimero.ZOrder 1
Me.cmdAnterior.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Anterior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdAnterior.ZOrder 0

Me.cmdPrimero.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Borrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdBorrar.ZOrder 0

Me.cmdPrimero.ZOrder 1
Me.cmdAnterior.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Buscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdBuscar.ZOrder 0

Me.cmdPrimero.ZOrder 1
Me.cmdAnterior.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub





Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdCancelar.ZOrder 0

Me.cmdPrimero.ZOrder 1
Me.cmdAnterior.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1

End Sub

Private Sub Pic_Modificar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdModificar.ZOrder 0

Me.cmdPrimero.ZOrder 1
Me.cmdAnterior.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Primero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdPrimero.ZOrder 0

Me.cmdAnterior.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Siguiente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdSiguiente.ZOrder 0

Me.cmdPrimero.ZOrder 1
Me.cmdAnterior.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Ultimo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdUltimo.ZOrder 0

Me.cmdPrimero.ZOrder 1
Me.cmdAnterior.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub


Sub f_Boton_Zorder()

Me.cmdPrimero.ZOrder 1
Me.cmdAnterior.ZOrder 1
Me.cmdBuscar.ZOrder 1
Me.cmdSiguiente.ZOrder 1
Me.cmdUltimo.ZOrder 1
Me.cmdAgregar.ZOrder 1
Me.cmdBorrar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub
