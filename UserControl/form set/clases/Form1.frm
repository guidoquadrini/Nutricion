VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":0000
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":0710
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":0E14
      Height          =   375
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":1524
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancelar"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":1C28
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":2338
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Aceptar"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdModificar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":2A3C
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":313E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Modificar"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBorrar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":3842
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":3F52
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Eliminar"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAgregar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":4656
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":4D66
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Agregar"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdUltimo 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":546A
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":5B7A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ultimo"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdSiguiente 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":627E
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":698E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Siguiente"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBuscar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":7092
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":77A2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAnterior 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":7EA6
      Height          =   375
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":85B6
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Anterior"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdPrimero 
      Appearance      =   0  'Flat
      DisabledPicture =   "Form1.frx":8CBA
      Height          =   375
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":93CA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Primero"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Label ContenedorBotones 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5745
   End
End
Attribute VB_Name = "Form1"
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

    MDIForm1.ActiveForm.Hide
    
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
Unload tabla1frm
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

    MDIForm1.ActiveForm.Hide

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

If MDIForm1.ActiveForm.Data1.Recordset.AbsolutePosition = 0 Then
    
    cmdSiguiente.Enabled = True
    cmdUltimo.Enabled = True
    
    cmdAnterior.Enabled = False
    cmdPrimero.Enabled = False

Else
    If MDIForm1.ActiveForm.Data1.Recordset.AbsolutePosition = MDIForm1.ActiveForm.Data1.Recordset.RecordCount - 1 Then
        
        cmdSiguiente.Enabled = False
        cmdUltimo.Enabled = False
        
        cmdAnterior.Enabled = True
        cmdPrimero.Enabled = True

    End If
End If
End Sub

