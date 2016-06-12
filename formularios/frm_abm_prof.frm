VERSION 5.00
Begin VB.Form frm_abm_prof 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profesionales"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frm_abm_prof.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6030
   Begin VB.Frame Frame2 
      Caption         =   "profesional"
      Height          =   495
      Left            =   1800
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label Label1 
         Caption         =   "Label1"
         DataField       =   "prf_codigo"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":0ECA
      Height          =   375
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":15DA
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Imprimir"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":1CDE
      Height          =   375
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":23EE
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Cancelar"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":2AF2
      Height          =   375
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":3202
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Aceptar"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdModificar 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":3906
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":4008
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Modificar"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBorrar 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":470C
      Height          =   375
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":4E1C
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Eliminar"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAgregar 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":5520
      Height          =   375
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":5C30
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Agregar"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdUltimo 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":6334
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":6A44
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Ultimo"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdSiguiente 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":7148
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":7858
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Siguiente"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBuscar 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":7F5C
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":866C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Buscar"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAnterior 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":8D70
      Height          =   375
      Left            =   1320
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":9480
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Anterior"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdPrimero 
      Appearance      =   0  'Flat
      DisabledPicture =   "frm_abm_prof.frx":9B84
      Height          =   375
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frm_abm_prof.frx":A294
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Primero"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "db1nueva prueba anterior sin replica.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Profesionales"
      Top             =   3525
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "Ac&tualizar"
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         ToolTipText     =   "Mostrar Todos"
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_durTurN"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox Text1 
         DataField       =   "prf_memo"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   1215
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_nombre"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_durTurD"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2280
         TabIndex        =   3
         Top             =   1350
         Width           =   615
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_apell"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   1
         Top             =   690
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Duracion Turno Normal:"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "minutos"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   11
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   390
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Duracion Turno Demanda:"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Caption         =   "minutos"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   8
         Top             =   1380
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Caption         =   "Observaciones:"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Apellido:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frm_abm_prof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Titulo As String 'titulo del form

Private Sub Command1_Click()
Dim strquery As String
strquery = " select * from profesionales order by prf_apell, prf_nombre"

With Data1
    .RecordSource = strquery
    .Refresh
End With

Call enabledDesplaz
End Sub




Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 4350
Me.Width = 6150
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Call enabledDesplaz 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados

End Sub

Private Sub Form_Load()
Data1.DatabaseName = Lugar

estadoAbm = 1 ' el estado es sim cambios

Titulo = Me.Caption

'-------------------------
'se refresca el data1 para que el metodo enabledDesplaz funcione correctamente con el recordset cargado
strquery = " select * from profesionales order by prf_apell, prf_nombre"

With Data1
    .RecordSource = strquery
    .Refresh
End With
'--------------------------------

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdCancelar_Click

End Sub

Private Sub Label1_Change()

Me.Caption = Titulo & " - Nro. " & Val(Label1.Caption)

End Sub

Private Sub Text1_GotFocus()
cmdAceptar.Default = False
cmdCancelar.Cancel = False

End Sub

Private Sub Text1_LostFocus()

cmdAceptar.Default = True
cmdCancelar.Cancel = True

End Sub

Private Sub cmdAceptar_Click()

If estadoAbm = 2 Or estadoAbm = 3 Then 'si el estado es agregar o modificar

    MDIForm1.ActiveForm.Data1.UpdateRecord
    MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
    
'    'condiciones extras
    If estadoAbm = 2 Then
        
        For i = 0 To 6
            dbdiet.Execute "insert into horarios (hrs_idprof, hrs_dia) select " & Val(MDIForm1.ActiveForm.Label1.Caption) & ", " & i
        Next
        
'        dbdiet.Execute "insert into histclinicas (legajo) select " & Val(MDIForm1.ActiveForm.Label1.Caption) '& ", codalimento from alimentos where estado = true"
    End If
        
    cmdBuscar.Enabled = True
    cmdAgregar.Enabled = True
'    cmdBorrar.Enabled = True
    'cmdClose.Enabled = True
'    cmdModificar.Enabled = True
    
    cmdAgregar.SetFocus
    cmdAgregar.Default = True
    cmdCancelar.Cancel = True
    
'    cmdPrimero.Enabled = True
'    cmdAnterior.Enabled = True
'    cmdSiguiente.Enabled = True
'    cmdUltimo.Enabled = True
      
    Call enabledDesplaz 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
        
    Call enabledDesplaz
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

cmdAceptar.Default = True
cmdCancelar.Cancel = True

End Sub

Private Sub cmdAnterior_Click()
'If MDIForm1.ActiveForm.Data1.Recordset.BOF = False Then
    MDIForm1.ActiveForm.Data1.Recordset.MovePrevious
'Else
'    MDIForm1.ActiveForm.Data1.Recordset.MoveLast
'End If
Call enabledDesplaz

End Sub

Private Sub cmdBorrar_Click()
  'esto puede producir un error si elimina el último
  'registro o el único registro del recordset

If MDIForm1.ActiveForm.Data1.Recordset.RecordCount > 0 And MDIForm1.ActiveForm.Data1.Recordset.EOF = False And MDIForm1.ActiveForm.Data1.Recordset.BOF = False Then
    msg = MsgBox("¿Desea Eliminar el registro actual?", vbYesNo, "Eliminar")
    
    If msg = vbYes Then
        'verifica que se pueda eliminar sin problemas y no perder integridad
        
'            strquery = "select * from alimenxpaciente where legajo = " & Val(Label1.Caption) & " and cantidad <> 0"
'
'            Set tb = dbdiet.OpenRecordset(strquery)
'            strquery = "select * from menu where legajo = " & Val(Label1.Caption)
'
'            Set tb1 = dbdiet.OpenRecordset(strquery)
'            If tb.RecordCount = 0 And tb1.RecordCount = 0 Then
                Data1.Recordset.Delete
                Data1.Recordset.MovePrevious
'                dbdiet.Execute "delete from alimenxpaciente where legajo = " & Val(Label1.Caption)
'                dbdiet.Execute "delete from menu where legajo = " & Val(Label1.Caption)
'                dbdiet.Execute "delete from platosmenu where legajo = " & Val(Label1.Caption)
'            Else
'                MsgBox "No se puede eliminar '" & txtFields(1).Text & "' porque puede afectar la integridad del Sistema", , "Información"
'            End If
'            tb.Close
'            tb1.Close
        
    Else
        cmdAgregar.SetFocus
    End If
End If

End Sub

Private Sub cmdBuscar_Click()
Dim strquery As String

strquery = " select * from profesionales order by prf_apell, prf_nombre"

With Data1
    .RecordSource = strquery
    .Refresh
End With

'aclare campo por el cual buscar
msg = InputBox("Ingrese apellido del profesional:", "Buscar por Apellido")
   
If msg <> "" Then

    strquery = " select * from profesionales where prf_apell like '" & msg & "*' order by prf_apell, prf_nombre"
    
    With MDIForm1.ActiveForm.Data1
        .RecordSource = strquery
        .Refresh
    End With
        
End If

Call enabledDesplaz
End Sub

Private Sub cmdCancelar_Click()
If estadoAbm = 2 Or estadoAbm = 3 Then ' el estado del form es agregar o modificar

    MDIForm1.ActiveForm.Data1.Recordset.CancelUpdate
    
    
    cmdBuscar.Enabled = True
    cmdAgregar.Enabled = True
    'cmdBorrar.Enabled = True
    'cmdClose.Enabled = True
    'cmdModificar.Enabled = True
    
    cmdAgregar.SetFocus
    cmdAgregar.Default = True
    'cmdClose.Cancel = True
    'cmdPrimero.Enabled = True
    'cmdAnterior.Enabled = True
    'cmdSiguiente.Enabled = True
    'cmdUltimo.Enabled = True
           
    Call enabledDesplaz 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
    
    Call enabledDesplaz
Else

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        MDIForm1.ActiveForm.Hide
    
    End If

End If
End Sub



Private Sub cmdImprimir_Click()
''aclare el filtro para imprimir
'CrystalReport1.SelectionFormula = " {pacientes.legajo} = " & Val(Label1.Caption) '& " and {platosmenu.fechaMenu} in Date(" & Year(DTdesde.Value) & ", " & Month(DTdesde.Value) & ", " & Day(DTdesde.Value) & ") to Date(" & Year(DThasta.Value) & ", " & Month(DThasta.Value) & ", " & Day(DThasta.Value) & ") "
'
'CrystalReport1.Destination = crptToWindow
'CrystalReport1.PrintReport

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

Call enabledDesplaz

End Sub

Private Sub cmdUltimo_Click()

MDIForm1.ActiveForm.Data1.Recordset.MoveLast

cmdSiguiente.Enabled = False
cmdUltimo.Enabled = False

cmdAnterior.Enabled = True
cmdPrimero.Enabled = True

End Sub

