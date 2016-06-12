VERSION 5.00
Begin VB.UserControl abmControl 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5745
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   615
   ScaleWidth      =   5745
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":0000
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":0710
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":0E14
      Height          =   375
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":1524
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":1C28
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":2338
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdModificar 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":2A3C
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":313E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBorrar 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":3842
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":3F52
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAgregar 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":4656
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":4D66
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdUltimo 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":546A
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":5B7A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdSiguiente 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":627E
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":698E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBuscar 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":7092
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":77A2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAnterior 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":7EA6
      Height          =   375
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":85B6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdPrimero 
      Appearance      =   0  'Flat
      DisabledPicture =   "clsBotonesAbm.ctx":8CBA
      Height          =   375
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "clsBotonesAbm.ctx":93CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Label ContenedorBotones 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5745
   End
End
Attribute VB_Name = "abmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim aux

Private Sub cmdAceptar_Click()

If estadoAbm = 2 Or estadoAbm = 3 Then 'si el estado es agregar o modificar

    MDIForm1.ActiveForm.Data1.UpdateRecord
    MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
    
    If aux = 0 Then
        dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(MDIForm1.ActiveForm.Label1.Caption) & ", codalimento from alimentos where estado = true"
    End If
        
    cmdbuscar.Enabled = True
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
    
    aux = 1
    
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
    
Else

    MDIForm1.ActiveForm.Hide
    
End If

End Sub

Private Sub cmdAgregar_Click()

Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

aux = 0
estadoAbm = 2 ' el estado es agregar

MDIForm1.ActiveForm.Data1.Recordset.AddNew

cmdAgregar.Enabled = False
cmdBorrar.Enabled = False
'cmdclose.Enabled = False
cmdModificar.Enabled = False
cmdbuscar.Enabled = False
cmdAceptar.Visible = True
cmdCancelar.Visible = True
cmdPrimero.Enabled = False
cmdAnterior.Enabled = False
cmdSiguiente.Enabled = False
cmdUltimo.Enabled = False

MDIForm1.ActiveForm.txtFields(1).SetFocus

Unload PrincipalFrm
Unload tabla1frm
Unload Form1

cmdAceptar.Default = True
cmdCancelar.Cancel = True

End Sub

Private Sub cmdAnterior_Click()
If MDIForm1.ActiveForm.Data1.Recordset.BOF = False Then
    MDIForm1.ActiveForm.Data1.Recordset.MovePrevious
Else
    MDIForm1.ActiveForm.Data1.Recordset.MoveLast
End If

End Sub

Private Sub cmdBorrar_Click()
  'esto puede producir un error si elimina el último
  'registro o el único registro del recordset

If MDIForm1.ActiveForm.Data1.Recordset.RecordCount > 0 And MDIForm1.ActiveForm.Data1.Recordset.EOF = False And MDIForm1.ActiveForm.Data1.Recordset.BOF = False Then
    msg = MsgBox("¿Desea Eliminar el registro actual?", vbYesNo, "Eliminar")
    
    If msg = vbYes Then
        'verifica que se pueda eliminar sin problemas y no perder integridad
        strquery = "select * from alimenxpaciente where legajo = " & Val(Label1.Caption) & " and cantidad <> 0"
                
        Set MDIForm1.ActiveForm.tb = dbdiet.OpenRecordset(strquery)
        strquery = "select * from menu where legajo = " & Val(Label1.Caption)
        
        Set tb1 = dbdiet.OpenRecordset(strquery)
        If tb.RecordCount = 0 And tb1.RecordCount = 0 Then
            Data1.Recordset.Delete
            Data1.Recordset.MovePrevious
            dbdiet.Execute "delete from alimenxpaciente where legajo = " & Val(Label1.Caption)
            dbdiet.Execute "delete from menu where legajo = " & Val(Label1.Caption)
            dbdiet.Execute "delete from platosmenu where legajo = " & Val(Label1.Caption)
        Else
            MsgBox "No se puede eliminar '" & txtFields(1).Text & "' porque puede afectar la integridad del Sistema", , "Información"
        End If
        tb.Close
        tb1.Close
        
    Else
        CmdAdd.SetFocus
    End If
End If

End Sub

Private Sub cmdCancelar_Click()
If estadoAbm = 2 Or estadoAbm = 3 Then ' el estado del form es agregar o modificar

    MDIForm1.ActiveForm.Data1.Recordset.CancelUpdate
    
    
    cmdbuscar.Enabled = True
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
    
    aux = 1
    
    
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
    
Else

    MDIForm1.ActiveForm.Hide

End If
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
cmdbuscar.Enabled = False
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

aux = 1
estadoAbm = 3 ' el estado es modificar

End Sub

Private Sub cmdPrimero_Click()

MDIForm1.ActiveForm.Data1.Recordset.MoveFirst

End Sub

Private Sub cmdSiguiente_Click()
If MDIForm1.ActiveForm.Data1.Recordset.EOF = False Then
    MDIForm1.ActiveForm.Data1.Recordset.MoveNext
Else
    MDIForm1.ActiveForm.Data1.Recordset.MoveFirst
End If

End Sub

Private Sub cmdUltimo_Click()
MDIForm1.ActiveForm.Data1.Recordset.MoveLast
End Sub

