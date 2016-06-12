VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_calLegajo 
   Caption         =   "Seleccione un Paciente"
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3465
   Icon            =   "frm_calLegajo.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   825
   ScaleWidth      =   3465
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "db1nueva prueba anterior sin replica.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pacientes"
      Top             =   1560
      Visible         =   0   'False
      Width           =   11775
   End
   Begin VB.Frame Frame8 
      Caption         =   "Nombre:"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmd_cerrar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_calLegajo.frx":0ECA
         Height          =   315
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_calLegajo.frx":15DA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancelar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton cmd_aceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_calLegajo.frx":1CDE
         Height          =   315
         Left            =   2400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_calLegajo.frx":23EE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Agregar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_calLegajo.frx":2AF2
         DataField       =   "Legajo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "nom"
         BoundColumn     =   "Legajo"
         Text            =   "DataCombo1"
      End
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   2
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_calLegajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()
frm_calendario.cargarParametros 1, DataCombo1.BoundText

frm_calendario.Show

Unload Me

End Sub

Private Sub cmd_cerrar_Click()
Unload Me

End Sub

Private Sub Form_Load()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 1230
Me.Width = 3585
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Data1.DatabaseName = Lugar

strQuery = " select * from pacientes where legajo = 1"

With Data1
    .RecordSource = strQuery
    .Refresh
End With
    
End Sub
