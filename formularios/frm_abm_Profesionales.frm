VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_abm_Profesionales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profesionales"
   ClientHeight    =   5640
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frm_abm_Profesionales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6765
   Begin VB.TextBox txtFields 
      DataField       =   "prf_durTurD"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   5760
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Paciente"
      Height          =   495
      Left            =   1920
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label Label1 
         Caption         =   "Label1"
         DataField       =   "prf_codigo"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   840
         TabIndex        =   31
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4560
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7080
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7200
      Top             =   2880
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
   Begin VB.Frame Frame1 
      Caption         =   "Datos Personales"
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Observaciones"
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "Ac&tualizar"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         ToolTipText     =   "Mostrar Todos"
         Top             =   4345
         Width           =   3375
      End
      Begin VB.CommandButton cmdObserva 
         Caption         =   "O&bservaciones>>>"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         ToolTipText     =   "Observaciones"
         Top             =   4600
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_matricula"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   9
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_durTurN"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   14
         Top             =   4035
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "prf_tpoDoc"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_abm_Profesionales.frx":0ECA
         Left            =   4320
         List            =   "frm_abm_Profesionales.frx":0EE3
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   "prf_memo"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   5040
         Width           =   6495
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6240
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "D:\Dietetica\rpts\rep_pacientes.rpt"
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "prf_fNacimiento"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2400
         TabIndex        =   12
         Top             =   3360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66650113
         CurrentDate     =   37867
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_codPostal"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   5
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_nombre"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   2
         Top             =   555
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_apell"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   3
         Top             =   870
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_direccion"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1170
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_ciudad"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   5
         Top             =   1485
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_tel"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   8
         Top             =   2445
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "prf_email"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   11
         Top             =   3060
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_abm_Profesionales.frx":0F0A
         DataField       =   "prf_idProvincia"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   2100
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "DescripProv"
         BoundColumn     =   "idProvincia"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frm_abm_Profesionales.frx":0F1F
         DataField       =   "prf_idSexo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Top             =   3690
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "descripSexo"
         BoundColumn     =   "idSexo"
         Text            =   "DataCombo2"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "prf_nroDoc"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   11
         Mask            =   "###.###.###"
         PromptChar      =   "_"
      End
      Begin VB.Label lblLabels 
         Caption         =   "Matricula:"
         Height          =   255
         Index           =   16
         Left            =   480
         TabIndex        =   38
         Top             =   255
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "minutos"
         Height          =   255
         Index           =   12
         Left            =   3120
         TabIndex        =   34
         Top             =   4050
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Duracion Turno Normal:"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   33
         Top             =   4050
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Tpo.:"
         Height          =   195
         Index           =   13
         Left            =   3795
         TabIndex        =   32
         Top             =   2820
         Width           =   375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Sexo:"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   27
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Provincia:"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   26
         Top             =   2130
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Teléfono:"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   25
         Top             =   2460
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Código Postal:"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   24
         Top             =   1815
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   23
         Top             =   570
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Apellido:"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   22
         Top             =   885
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Dirección:"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   21
         Top             =   1185
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ciudad:"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   20
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Nacimiento:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   3390
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Email:"
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   18
         Top             =   3075
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nro. Documento:"
         Height          =   255
         Index           =   11
         Left            =   480
         TabIndex        =   17
         Top             =   2775
         Width           =   1815
      End
   End
   Begin VB.Frame fme_botones_abm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   615
      TabIndex        =   39
      Top             =   5040
      Width           =   5535
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":0F34
         Height          =   375
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":108D
         Picture         =   "frm_abm_Profesionales.frx":11DF
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Aceptar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":149B
         Height          =   375
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":162F
         Picture         =   "frm_abm_Profesionales.frx":1781
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Cancelar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox Pic_Cancelar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5160
         MouseIcon       =   "frm_abm_Profesionales.frx":1C34
         Picture         =   "frm_abm_Profesionales.frx":1D86
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   65
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
         MouseIcon       =   "frm_abm_Profesionales.frx":2087
         Picture         =   "frm_abm_Profesionales.frx":21D9
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   64
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":2495
         Height          =   375
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":25B6
         Picture         =   "frm_abm_Profesionales.frx":2708
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Modificar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":297B
         Height          =   375
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":2A91
         Picture         =   "frm_abm_Profesionales.frx":2BE3
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Eliminar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":2D72
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":2EBF
         Picture         =   "frm_abm_Profesionales.frx":3011
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Agregar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUltimo 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":344B
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":35F3
         Picture         =   "frm_abm_Profesionales.frx":3745
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Ultimo"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSiguiente 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":3C10
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":3D7D
         Picture         =   "frm_abm_Profesionales.frx":3ECF
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Siguiente"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":4344
         Height          =   375
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":44CC
         Picture         =   "frm_abm_Profesionales.frx":461E
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Buscar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":48FB
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":4A65
         Picture         =   "frm_abm_Profesionales.frx":4BB7
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Anterior"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimero 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":5025
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Profesionales.frx":51CA
         Picture         =   "frm_abm_Profesionales.frx":531C
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Primero"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox Pic_Primero 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         MouseIcon       =   "frm_abm_Profesionales.frx":57D7
         Picture         =   "frm_abm_Profesionales.frx":5929
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   47
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
         MouseIcon       =   "frm_abm_Profesionales.frx":5DE4
         Picture         =   "frm_abm_Profesionales.frx":5F36
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   46
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
         MouseIcon       =   "frm_abm_Profesionales.frx":63A4
         Picture         =   "frm_abm_Profesionales.frx":64F6
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   45
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
         MouseIcon       =   "frm_abm_Profesionales.frx":67D3
         Picture         =   "frm_abm_Profesionales.frx":6925
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   44
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
         MouseIcon       =   "frm_abm_Profesionales.frx":6D9A
         Picture         =   "frm_abm_Profesionales.frx":6EEC
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   43
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
         MouseIcon       =   "frm_abm_Profesionales.frx":73B7
         Picture         =   "frm_abm_Profesionales.frx":7509
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   42
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
         MouseIcon       =   "frm_abm_Profesionales.frx":7943
         Picture         =   "frm_abm_Profesionales.frx":7A95
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   41
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
         MouseIcon       =   "frm_abm_Profesionales.frx":7C24
         Picture         =   "frm_abm_Profesionales.frx":7D76
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   40
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
         MouseIcon       =   "frm_abm_Profesionales.frx":7FE9
         Picture         =   "frm_abm_Profesionales.frx":813B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   56
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
         MouseIcon       =   "frm_abm_Profesionales.frx":825C
         Picture         =   "frm_abm_Profesionales.frx":83AE
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   57
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
         MouseIcon       =   "frm_abm_Profesionales.frx":84C4
         Picture         =   "frm_abm_Profesionales.frx":8616
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   58
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
         MouseIcon       =   "frm_abm_Profesionales.frx":8763
         Picture         =   "frm_abm_Profesionales.frx":88B5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   59
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
         MouseIcon       =   "frm_abm_Profesionales.frx":8A5D
         Picture         =   "frm_abm_Profesionales.frx":8BAF
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   60
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
         MouseIcon       =   "frm_abm_Profesionales.frx":8D1C
         Picture         =   "frm_abm_Profesionales.frx":8E6E
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   61
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
         MouseIcon       =   "frm_abm_Profesionales.frx":8FF6
         Picture         =   "frm_abm_Profesionales.frx":9148
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   62
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
         MouseIcon       =   "frm_abm_Profesionales.frx":92B2
         Picture         =   "frm_abm_Profesionales.frx":9404
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   63
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
         MouseIcon       =   "frm_abm_Profesionales.frx":95A9
         Picture         =   "frm_abm_Profesionales.frx":96FB
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   68
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
         MouseIcon       =   "frm_abm_Profesionales.frx":9854
         Picture         =   "frm_abm_Profesionales.frx":99A6
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   69
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Profesionales.frx":9B3A
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_abm_Profesionales.frx":9C92
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Imprimir"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox Pic_Imprimir_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         MouseIcon       =   "frm_abm_Profesionales.frx":A112
         Picture         =   "frm_abm_Profesionales.frx":A264
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   72
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Imprimir 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         MouseIcon       =   "frm_abm_Profesionales.frx":A3BC
         Picture         =   "frm_abm_Profesionales.frx":A50E
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   70
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5220
      Visible         =   0   'False
      Width           =   6765
   End
   Begin VB.Label lblLabels 
      Caption         =   "Duracion Turno Demanda:"
      Height          =   255
      Index           =   14
      Left            =   3840
      TabIndex        =   37
      Top             =   30
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Caption         =   "minutos"
      Height          =   255
      Index           =   15
      Left            =   6480
      TabIndex        =   36
      Top             =   15
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frm_abm_Profesionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim axp As Recordset
Dim alimen As Recordset

Dim msg As String
Dim tb As Recordset
Dim tb1 As Recordset
Dim cantReg As Integer
'Public estadoAbm As Integer ' define el estado de un formulario de abm
'                             1 = sin cambios; 2 = agregar; 3 = modificar
'el modulo "fSetEnableFields(MDIForm1.ActiveForm, vbFalse)" se debe agregar al proyecto
Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdCancelar_Click

End Sub

Private Sub Label1_Change()

Me.Caption = "Profesionales " & " - Nro. " & Label1.Caption

End Sub



Private Sub MaskEdBox1_GotFocus()

MaskEdBox1.SelStart = 0
MaskEdBox1.SelLength = 50

End Sub

Private Sub MaskEdBox1_LostFocus()
MaskEdBox1.SelStart = 0
MaskEdBox1.SelLength = 0
End Sub

Private Sub Pic_Imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Imprimir

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
    
    If txtFields(9).Text <> "" Then
    
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
          
        'Call enabledDesplaz 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados
        Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)
    
        estadoAbm = 1 ' el estado del form es "sin cambios"
            
        Call enabledDesplaz
        
        Call f_Boton_Zorder
    Else
        MsgBox "El campo matricula no puede estar vacio", vbInformation
    End If
    
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

MDIForm1.ActiveForm.txtFields(0).SetFocus

cmdAceptar.Default = True
cmdCancelar.Cancel = True

Call f_Boton_Zorder

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

Call f_Boton_Zorder

End Sub

Private Sub cmdBuscar_Click()
Dim strQuery As String

strQuery = " select * from profesionales order by prf_apell, prf_nombre"

With Data1
    .RecordSource = strQuery
    .Refresh
End With

'aclare campo por el cual buscar
msg = InputBox("Ingrese apellido del profesional:", "Buscar por Apellido")
   
If msg <> "" Then

    strQuery = " select * from profesionales where prf_apell like '" & msg & "*' order by prf_apell, prf_nombre"
    
    With MDIForm1.ActiveForm.Data1
        .RecordSource = strQuery
        .Refresh
    End With
        
End If

Call enabledDesplaz

Call f_Boton_Zorder

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
           
    Call f_Boton_Zorder
    
Else

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        MDIForm1.ActiveForm.Hide
    
    End If

End If
End Sub



Private Sub cmdImprimir_Click()
Dim strQuery, sMsg As String

'Resets the value of all properties (except DataSource Property) to their default values.
CrystalReport1.Reset

sMsg = MsgBox("¿Desea imprimir todos los registros?" & vbCrLf & vbTab & "- Si, para ver todos" & vbCrLf & vbTab & "- No, para ver solo el registro actual", vbYesNoCancel, "Imprimir")
  
If sMsg = vbYes Then

    CrystalReport1.ReportFileName = App_Path & "\rpts\rep_profesionales_all.rpt"
        
    'CrystalReport1.ParameterFields(4) = "SortField;Legajo;True"
    
    'CrystalReport1.ParameterFields(4) = "SortField;Obra Social;True"
    
    CrystalReport1.ParameterFields(4) = "SortField;ApellyNom;True"
        
    strQuery = ""
            
Else
    
    If sMsg = vbNo Then
    
        CrystalReport1.ReportFileName = App_Path & "\rpts\rep_profesionales_one.rpt"
        
        strQuery = " {profesionales.prf_codigo} = " & Val(Label1.Caption)
    
    End If
    
End If

If Not sMsg = vbCancel Then
    Call f_print(CrystalReport1, strQuery, crptToWindow)
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
cmdBuscar.Enabled = False
cmdAceptar.Visible = True
cmdCancelar.Visible = True
cmdPrimero.Enabled = False
cmdAnterior.Enabled = False
cmdSiguiente.Enabled = False
cmdUltimo.Enabled = False


MDIForm1.ActiveForm.Data1.Recordset.Edit
MDIForm1.ActiveForm.txtFields(0).SetFocus

cmdAceptar.Default = True
cmdCancelar.Cancel = True

estadoAbm = 3 ' el estado es modificar

Call f_Boton_Zorder

End Sub


Private Sub cmdPrimero_Click()

MDIForm1.ActiveForm.Data1.Recordset.MoveFirst

'cmdSiguiente.Enabled = True
'cmdUltimo.Enabled = True
'
'cmdAnterior.Enabled = False
'cmdPrimero.Enabled = False

Call enabledDesplaz

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

'cmdSiguiente.Enabled = False
'cmdUltimo.Enabled = False
'
'cmdAnterior.Enabled = True
'cmdPrimero.Enabled = True

Call enabledDesplaz
End Sub


Private Sub cmdObserva_Click()
If Me.cmdObserva.Caption = "O&bservaciones>>>" Then

    Me.Height = 7515
    
    fme_botones_abm.Top = 6480 '6360
    
    Me.Frame1.Height = 6495
    
    Me.cmdObserva.Caption = "<<<O&bservaciones"

Else

    Me.Height = 6015
        
    fme_botones_abm.Top = 5040 '4800
    
    Me.Frame1.Height = 4935
    
    Me.cmdObserva.Caption = "O&bservaciones>>>"

End If

'If Me.cmdObserva.Caption = "O&bservaciones>>>" Then

'    Me.Height = 9000
'    Me.TabStrip1.Height = 8415
'    Me.cmdAceptar.Top = 7800
'    Me.cmdAgregar.Top = 7800
'    Me.cmdAnterior.Top = 7800
'    Me.cmdBorrar.Top = 7800
'    Me.cmdBuscar.Top = 7800
'    Me.cmdCancelar.Top = 7800
'    Me.cmdImprimir.Top = 7800
'    Me.cmdModificar.Top = 7800
'    Me.cmdPrimero.Top = 7800
'    Me.cmdSiguiente.Top = 7800
'    Me.cmdUltimo.Top = 7800
'    Me.Frame1.Height = 6495
'
'    Me.cmdObserva.Caption = "<<<O&bservaciones"
'
'Else
'
'    Me.Height = 7230
'    Me.TabStrip1.Height = 6615
'    Me.cmdAceptar.Top = 6240
'    Me.cmdAgregar.Top = 6240
'    Me.cmdAnterior.Top = 6240
'    Me.cmdBorrar.Top = 6240
'    Me.cmdBuscar.Top = 6240
'    Me.cmdCancelar.Top = 6240
'    Me.cmdImprimir.Top = 6240
'    Me.cmdModificar.Top = 6240
'    Me.cmdPrimero.Top = 6240
'    Me.cmdSiguiente.Top = 6240
'    Me.cmdUltimo.Top = 6240
'    Me.Frame1.Height = 4935
'
'    Me.cmdObserva.Caption = "O&bservaciones>>>"

'End If

End Sub


Private Sub Command1_Click()
Dim strQuery As String
strQuery = " select * from profesionales order by prf_apell, prf_nombre"

With Data1
    .RecordSource = strQuery
    .Refresh
End With

Call enabledDesplaz
End Sub



Private Sub Command2_Click()
'dbdiet.Execute "insert into histclinicas (legajo) select " & Val(MDIForm1.ActiveForm.Label1.Caption)
txt_memo.Show vbModal

End Sub


Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'Aquí es donde se coloca el código de control de errores
  'Si quiere ignorar los errores, marque como comentario la línea siguiente
  'Si desea detectarlos, agregue código aquí para controlarlos
  MsgBox "El error de datos alcanzó err:" & Error$(DataErr)
  Response = 0  'ignorar el error
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'Esto mostrará la posición del registro actual
  'para dynasets y snapshots
  Data1.Caption = "Registros: " & (Data1.Recordset.RecordCount) 'AbsolutePosition + 1)
  'para el objeto tabla debe establecer la propiedad index cuando
  'se crea el recordset; use la línea siguiente
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  'Aquí es donde se coloca el código de validación
  'Se llama a este evento cuando se produce la siguiente acción
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
 ' Screen.MousePointer = vbHourglass
End Sub

Private Sub DataCombo1_LostFocus()
If DataCombo1.Text = "" Then
    DataCombo1.SetFocus
    MsgBox "Debe Completar la Provincia", vbInformation, "Información"
End If

End Sub

Private Sub DataCombo2_LostFocus()
If DataCombo2.Text = "" Then
    DataCombo2.SetFocus
    MsgBox "Debe Completar el Sexo del Paciente", vbInformation, "Información"
End If

End Sub

Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 5970
Me.Width = 6855
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Call enabledDesplaz

Call f_Boton_Zorder

End Sub

Private Sub Form_Load()

Call f_CargarOrigenDatos

For i = 0 To 9
    txtFields(i).Enabled = False
Next

DTPicker1.Enabled = False
aux = 1

estadoAbm = 1 ' el estado es sim cambios

'-------------------------
''se refresca el data1 para que el metodo enabledDesplaz funcione correctamente con el recordset cargado
'strquery = " select * from profesionales order by prf_apell, prf_nombre"
'
'With Data1
'    .RecordSource = strquery
'    .Refresh
'End With
''--------------------------------

End Sub

Private Sub txtFields_GotFocus(Index As Integer)
For i = 0 To 9
    txtFields(i).SelStart = 0
    txtFields(i).SelLength = 50
Next

End Sub

Private Sub txtFields_Validate(Index As Integer, Cancel As Boolean)

If Index = 9 And txtFields(9).Text = "" Then
    
    MsgBox "El campo matricula no puede estar vacio", vbInformation
    Cancel = True
    
Else
        
    Cancel = False

End If

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Data1.Recordset = Nothing
Set Me.Adodc1.Recordset = Nothing
Set Me.Adodc2.Recordset = Nothing

strQuery = " select * from profesionales order by prf_apell, prf_nombre"
Call f_Data_DatabaseName(Data1, strQuery)

strQuery = "select * from provincia order by descripprov"
Call f_Adodc_ConnectionString(Adodc1, strQuery)

strQuery = "sexo"
Call f_Adodc_ConnectionString(Adodc2, strQuery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo1, Data1, Adodc1, "prf_idProvincia", "idProvincia", "DescripProv")

Call f_Enlaza_ControlData(DataCombo2, Data1, Adodc2, "prf_idSexo", "idSexo", "descripSexo")

'==============================================

End Sub


Private Sub fme_botones_abm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call f_Aceptar

End Sub

Private Sub Pic_Agregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Agregar

End Sub

Private Sub Pic_Anterior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call f_Anterior

End Sub

Private Sub Pic_Borrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Borrar

End Sub

Private Sub Pic_Buscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call f_Buscar

End Sub

Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cancelar

End Sub


Private Sub Pic_Modificar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Modificar

End Sub

Private Sub Pic_Primero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Primero

End Sub

Private Sub Pic_Siguiente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Siguiente

End Sub

Private Sub Pic_Ultimo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Ultimo

End Sub


Sub f_Boton_Zorder()

If Me.cmdImprimir.Enabled = True Then
    Me.Pic_Imprimir.ZOrder 0
Else
    Me.Pic_Imprimir_Gris.ZOrder 0
End If

If Me.cmdPrimero.Enabled = True Then
    Me.Pic_Primero.ZOrder 0
Else
    Me.Pic_Primero_Gris.ZOrder 0
End If

If Me.cmdAnterior.Enabled = True Then
    Me.Pic_Anterior.ZOrder 0
Else
    Me.Pic_Anterior_Gris.ZOrder 0
End If

If Me.cmdBuscar.Enabled = True Then
    Me.Pic_Buscar.ZOrder 0
Else
    Me.Pic_Buscar_Gris.ZOrder 0
End If

If Me.cmdSiguiente.Enabled = True Then
    Me.Pic_Siguiente.ZOrder 0
Else
    Me.Pic_Siguiente_Gris.ZOrder 0
End If

If Me.cmdUltimo.Enabled = True Then
    Me.Pic_Ultimo.ZOrder 0
Else
    Me.Pic_Ultimo_Gris.ZOrder 0
End If

If Me.cmdAgregar.Enabled = True Then
    Me.Pic_Agregar.ZOrder 0
Else
    Me.Pic_Agregar_Gris.ZOrder 0
End If

If Me.cmdBorrar.Enabled = True Then
    Me.Pic_Borrar.ZOrder 0
Else
    Me.Pic_Borrar_Gris.ZOrder 0
End If

If Me.cmdModificar.Enabled = True Then
    Me.Pic_Modificar.ZOrder 0
Else
    Me.Pic_Modificar_Gris.ZOrder 0
End If

If Me.cmdAceptar.Enabled = True Then
    Me.Pic_Aceptar.ZOrder 0
Else
    Me.Pic_Aceptar_Gris.ZOrder 0
End If

If Me.cmdCancelar.Enabled = True Then
    Me.Pic_Cancelar.ZOrder 0
Else
    Me.Pic_Cancelar_Gris.ZOrder 0
End If

Me.cmdImprimir.ZOrder 1
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

Sub f_Primero()

Me.cmdPrimero.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Anterior()
Me.cmdAnterior.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Buscar()
Me.cmdBuscar.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Siguiente()
Me.cmdSiguiente.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Ultimo()
Me.cmdUltimo.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Agregar()

Me.cmdAgregar.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Borrar()

Me.cmdBorrar.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Modificar()

Me.cmdModificar.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Aceptar()

Me.cmdAceptar.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Cancelar()

Me.cmdCancelar.ZOrder 0

Me.cmdImprimir.ZOrder 1
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

Sub f_Imprimir()

Me.cmdImprimir.ZOrder 0

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


