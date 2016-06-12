VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm_Evaluacion_Subjetiva 
   Caption         =   "Evaluacion subjetiva del estado nutricional"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   ControlBox      =   0   'False
   Icon            =   "frm_Evaluacion_Subjetiva.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   11070
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8040
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.TextBox Text1 
      DataField       =   "legajo"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   6000
      TabIndex        =   58
      Text            =   "legajo"
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txt_idHistClinica 
      DataField       =   "idHistClinica"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5160
      TabIndex        =   56
      Text            =   "idHistClinica"
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "db1nueva prueba anterior sin replica.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from histclinicas where legajo = 1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame 
      Caption         =   "Paciente:"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   4935
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   0
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "D:\Dietetica\rpts\rep_histclinicas.rpt"
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   2160
         TabIndex        =   59
         Top             =   120
         Width           =   2655
         Begin VB.CommandButton cmdModificar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Evaluacion_Subjetiva.frx":0ECA
            Height          =   375
            Left            =   1320
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":0FEB
            Picture         =   "frm_Evaluacion_Subjetiva.frx":113D
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Modificar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdAceptar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Evaluacion_Subjetiva.frx":13B0
            Height          =   375
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":1509
            Picture         =   "frm_Evaluacion_Subjetiva.frx":165B
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Aceptar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdCancelar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Evaluacion_Subjetiva.frx":1917
            Height          =   375
            Left            =   2280
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":1AAB
            Picture         =   "frm_Evaluacion_Subjetiva.frx":1BFD
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Cancelar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.PictureBox Pic_Aceptar_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1800
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":20B0
            Picture         =   "frm_Evaluacion_Subjetiva.frx":2202
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   64
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Cancelar_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2280
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":235B
            Picture         =   "frm_Evaluacion_Subjetiva.frx":24AD
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   65
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Modificar_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1320
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":2641
            Picture         =   "frm_Evaluacion_Subjetiva.frx":2793
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   68
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Modificar 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1320
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":28B4
            Picture         =   "frm_Evaluacion_Subjetiva.frx":2A06
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   66
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Aceptar 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1800
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":2C79
            Picture         =   "frm_Evaluacion_Subjetiva.frx":2DCB
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   62
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Cancelar 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2280
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":3087
            Picture         =   "frm_Evaluacion_Subjetiva.frx":31D9
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   63
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdImprimir 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Evaluacion_Subjetiva.frx":34DA
            Height          =   375
            Left            =   720
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_Evaluacion_Subjetiva.frx":3632
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Imprimir"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.PictureBox Pic_Imprimir 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   720
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":3AB2
            Picture         =   "frm_Evaluacion_Subjetiva.frx":3C04
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   69
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Imprimir_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   720
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":4084
            Picture         =   "frm_Evaluacion_Subjetiva.frx":41D6
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   71
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmd_Tipito 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Evaluacion_Subjetiva.frx":432E
            Height          =   315
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_Evaluacion_Subjetiva.frx":4A3E
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Info"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.PictureBox Pic_Tipito_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":4CCE
            Picture         =   "frm_Evaluacion_Subjetiva.frx":4E20
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   74
            Top             =   120
            Width           =   315
         End
         Begin VB.PictureBox Pic_Tipito 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            MouseIcon       =   "frm_Evaluacion_Subjetiva.frx":4F50
            Picture         =   "frm_Evaluacion_Subjetiva.frx":50A2
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   73
            Top             =   120
            Width           =   315
         End
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "RESULTADOS"
      Height          =   855
      Left            =   5760
      TabIndex        =   25
      Top             =   5760
      Width           =   5175
      Begin VB.ComboBox cmb_resultado 
         DataField       =   "resultado"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Evaluacion_Subjetiva.frx":5332
         Left            =   3120
         List            =   "frm_Evaluacion_Subjetiva.frx":533F
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Seleccione un resultado"
         Height          =   195
         Left            =   1200
         TabIndex        =   52
         Top             =   240
         Width           =   1710
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "EXPLORACION FISICA"
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      Top             =   5760
      Width           =   5415
      Begin VB.ComboBox cmb_ascitis 
         DataField       =   "ascitis"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Evaluacion_Subjetiva.frx":5382
         Left            =   3600
         List            =   "frm_Evaluacion_Subjetiva.frx":5392
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox cmb_muscular 
         DataField       =   "muscular"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Evaluacion_Subjetiva.frx":53B5
         Left            =   3600
         List            =   "frm_Evaluacion_Subjetiva.frx":53C5
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmb_tobillo 
         DataField       =   "tobillo"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Evaluacion_Subjetiva.frx":53E8
         Left            =   3600
         List            =   "frm_Evaluacion_Subjetiva.frx":53F8
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cmb_sacro 
         DataField       =   "sacro"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Evaluacion_Subjetiva.frx":541B
         Left            =   3600
         List            =   "frm_Evaluacion_Subjetiva.frx":542B
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox cmb_gsaSubcnea 
         DataField       =   "gsaSubcnea"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Evaluacion_Subjetiva.frx":544E
         Left            =   3600
         List            =   "frm_Evaluacion_Subjetiva.frx":545E
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "Ascitis"
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label16 
         Caption         =   "Edema sacro"
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label15 
         Caption         =   "Edema de tobillo"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label14 
         Caption         =   "Desgaste muscular (cuadriceps, deltoides)"
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "Perdida de grasa subcutanea (tricipital)"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Enfermedad y relacion"
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   10815
      Begin VB.Frame Frame8 
         Caption         =   "Requerimientos metabolicos"
         Height          =   975
         Left            =   2880
         TabIndex        =   41
         Top             =   120
         Width           =   5295
         Begin VB.ComboBox cmb_NvlEstres 
            DataField       =   "nvlEstres"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frm_Evaluacion_Subjetiva.frx":5481
            Left            =   3480
            List            =   "frm_Evaluacion_Subjetiva.frx":5491
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Opt_ConEstres 
            Caption         =   "Con estres"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   43
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt_SinEstres 
            Caption         =   "Sin estres"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Nivel estres"
            Height          =   255
            Left            =   2520
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox txt_diagnostico 
         DataField       =   "diagnostico"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "Diagnostico Principal"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "c. capacidad funcional"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   10815
      Begin VB.ComboBox cmb_CapFunc 
         DataField       =   "CapFunc"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Evaluacion_Subjetiva.frx":54CF
         Left            =   8760
         List            =   "frm_Evaluacion_Subjetiva.frx":54DF
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txt_semDisf 
         DataField       =   "semDisf"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txt_durDisf 
         DataField       =   "durDisf"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Opt_Condisfuncion 
         Caption         =   "Con disfuncion"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Opt_Sindisfuncion 
         Caption         =   "Sin disfuncion"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   8280
         TabIndex        =   39
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Semanas"
         Height          =   255
         Left            =   5520
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Duracion"
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "c. sintomas gastrointestinales"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   10815
      Begin VB.CheckBox Chk_anorexia 
         Caption         =   "Anorexia"
         DataField       =   "ChkAnorexiaSinGastro"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Chk_diarrea 
         Caption         =   "Diarrea"
         DataField       =   "ChkDiarreaSinGastro"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Chk_vomito 
         Caption         =   "Vómito"
         DataField       =   "ChkVomitoSinGastro"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Chk_nauseas 
         Caption         =   "Náuseas"
         DataField       =   "ChkNauseasSintGastro"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Chk_no 
         Caption         =   "No"
         DataField       =   "ChknoSintGastro"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "b. ingestion dietaria"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   10815
      Begin VB.TextBox txt_durCbio 
         DataField       =   "durCbio"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txt_semCbio 
         DataField       =   "semCbio"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmb_tpo 
         DataField       =   "tpoCambio"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Evaluacion_Subjetiva.frx":551D
         Left            =   8760
         List            =   "frm_Evaluacion_Subjetiva.frx":5530
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Opt_ConCambios 
         Caption         =   "Con cambios"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Opt_SinCambios 
         Caption         =   "Sin cambios"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   8280
         TabIndex        =   36
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Semanas"
         Height          =   255
         Left            =   5520
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Duracion"
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "a. perdida de peso"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10815
      Begin VB.TextBox txt_pdaSeisMeses 
         DataField       =   "pdaSeisMeses"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txt_porcent 
         Appearance      =   0  'Flat
         DataField       =   "porcent"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7800
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cmb_cambioPeso 
         DataField       =   "cambioPeso"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Evaluacion_Subjetiva.frx":557D
         Left            =   3600
         List            =   "frm_Evaluacion_Subjetiva.frx":558A
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Porcentaje de perdida de peso"
         Height          =   255
         Left            =   5520
         TabIndex        =   31
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Cambio de peso en las dos ultimas semanas"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Perdida Global en los ultimos seis meses (Kg)"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   3255
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1080
      Top             =   0
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   "OLEDB_Omnia.UDL"
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
   Begin RichTextLib.RichTextBox rtfTxtHc 
      Height          =   7935
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   13996
      _Version        =   393217
      BackColor       =   -2147483644
      Enabled         =   0   'False
      Appearance      =   0
      TextRTF         =   $"frm_Evaluacion_Subjetiva.frx":55AB
   End
   Begin VB.Label ContenedorBotones 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   3120
      TabIndex        =   57
      Top             =   240
      Width           =   5745
   End
End
Attribute VB_Name = "frm_Evaluacion_Subjetiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public estadoAbm As Integer ' define el estado de un formulario de abm
'                             1 = sin cambios; 2 = agregar; 3 = modificar
'el modulo "fSetEnableFields(MDIForm1.ActiveForm, vbFalse)" se debe agregar al proyecto
Dim tb As Recordset
Dim Peso As Single
Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
                            
Private Sub cmdAceptar_Click()

If estadoAbm = 2 Or estadoAbm = 3 Then 'si el estado es agregar o modificar
    
    MDIForm1.ActiveForm.Data1.UpdateRecord
    MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
    
    'condiciones extras
        'If estadoAbm = 2 Then
        '    dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(MDIForm1.ActiveForm.Label1.Caption) & ", codalimento from alimentos where estado = true"
        'End If
        
'    cmdbuscar.Enabled = True
'    cmdAgregar.Enabled = True
'    cmdBorrar.Enabled = True
    'cmdClose.Enabled = True
    cmdModificar.Enabled = True
    
'    cmdAgregar.SetFocus
'    cmdAgregar.Default = True
    cmdCancelar.Cancel = True
    
'    cmdPrimero.Enabled = True
'    cmdAnterior.Enabled = True
'    cmdSiguiente.Enabled = True
'    cmdUltimo.Enabled = True
   
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
        
    DataCombo1.Enabled = True
        
    Call f_Boton_Zorder
    
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

Unload frm_FormulaSintetica
Unload frm_formulaDesarrollada
Unload frm_Adm_Diet

cmdAceptar.Default = True
cmdCancelar.Cancel = True

End Sub


Private Sub cmdCancelar_Click()
If estadoAbm = 2 Or estadoAbm = 3 Then ' el estado del form es agregar o modificar

    MDIForm1.ActiveForm.Data1.Recordset.CancelUpdate
    
    
'    cmdbuscar.Enabled = True
'    cmdAgregar.Enabled = True
'    cmdBorrar.Enabled = True
    'cmdClose.Enabled = True
    cmdModificar.Enabled = True
    
'    cmdAgregar.SetFocus
'    cmdAgregar.Default = True
    'cmdClose.Cancel = True
'    cmdPrimero.Enabled = True
'    cmdAnterior.Enabled = True
'    cmdSiguiente.Enabled = True
'    cmdUltimo.Enabled = True
           
    
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
        
    DataCombo1.Enabled = True
        
    Call f_Boton_Zorder
    
Else

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        Unload Me
        'MDIForm1.ActiveForm.Hide
    
    End If

End If
End Sub



Private Sub cmdImprimir_Click()
Dim strQuery As String

CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_histclinicas_one.rpt"

strQuery = " {histClinicas.idhistclinica} = " & Val(txt_idHistClinica)

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub

Private Sub cmdModificar_Click()

Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

If MDIForm1.ActiveForm.Data1.Recordset.BOF = True Or MDIForm1.ActiveForm.Data1.Recordset.EOF = True Then
    MDIForm1.ActiveForm.Data1.Recordset.MoveFirst
End If

'cmdAgregar.Enabled = False
'cmdBorrar.Enabled = False
'cmdclose.Enabled = False
cmdModificar.Enabled = False
'cmdbuscar.Enabled = False
cmdAceptar.Visible = True
cmdCancelar.Visible = True
'cmdPrimero.Enabled = False
'cmdAnterior.Enabled = False
'cmdSiguiente.Enabled = False
'cmdUltimo.Enabled = False


MDIForm1.ActiveForm.Data1.Recordset.Edit
'MDIForm1.ActiveForm.txtFields(1).SetFocus

cmdAceptar.Default = False 'True
cmdCancelar.Cancel = False 'True

estadoAbm = 3 ' el estado es modificar

Call f_Boton_Zorder

End Sub



Private Sub Chk_no_Click()
If estadoAbm = 3 Then

    If Chk_no.Value Then
        Chk_nauseas.Enabled = False
        Chk_vomito.Enabled = False
        Chk_diarrea.Enabled = False
        Chk_anorexia.Enabled = False
        
        Chk_nauseas.Value = False
        Chk_vomito.Value = False
        Chk_diarrea.Value = False
        Chk_anorexia.Value = False
        
    Else
        Chk_nauseas.Enabled = True
        Chk_vomito.Enabled = True
        Chk_diarrea.Enabled = True
        Chk_anorexia.Enabled = True
    End If

End If

End Sub


Private Sub cmd_tipito_Click()

Unload frmPacientes
frmPacientes.Show
frmPacientes.Data1.Recordset.FindFirst " legajo = " & DataCombo1.BoundText

End Sub

Private Sub DataCombo1_Click(Area As Integer)
Dim strQuery As String

If DataCombo1.Text <> "" Then
       
    strQuery = "select * from histclinicas where legajo = " & DataCombo1.BoundText

    With Data1
        .RecordSource = strQuery
        .Refresh
    End With
          
End If

End Sub

Private Sub Form_Load()
Dim strQuery As String

'Data1.DatabaseName = Lugar

Call f_CargarOrigenDatos

'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 8355
Me.Width = 11190
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

'Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)
'Adodc2.Refresh

'DataCombo1.BoundText = Val(Text1.Text)  'nLegajo
If Val(Text1.Text) > 0 Then

    DataCombo1.BoundText = Val(Text1.Text) 'nLegajo
    
End If

If DataCombo1.Text <> "" Then
    
    strQuery = "select * from histclinicas where legajo = " & DataCombo1.BoundText
    
    With Data1
        .RecordSource = strQuery
        .Refresh
    End With
End If

Call f_Boton_Zorder

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdCancelar_Click

End Sub

Private Sub Frame9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Opt_ConCambios_Click()

If estadoAbm = 3 Then

    txt_durCbio.Enabled = True
    txt_semCbio.Enabled = True
    cmb_tpo.Enabled = True

    dbdiet.Execute " update histclinicas set cambiosIngDiet = true where legajo = " & DataCombo1.BoundText

End If

End Sub

Private Sub Opt_Condisfuncion_Click()
If estadoAbm = 3 Then

    txt_durDisf.Enabled = True
    txt_semDisf.Enabled = True
    cmb_CapFunc.Enabled = True
    
    dbdiet.Execute " update histclinicas set disfuncion = true where legajo = " & DataCombo1.BoundText

End If

End Sub

Private Sub Opt_ConEstres_Click()
If estadoAbm = 3 Then

    cmb_NvlEstres.Enabled = True
    
    dbdiet.Execute " update histclinicas set estres = true where legajo = " & DataCombo1.BoundText

End If

End Sub

Private Sub Opt_SinCambios_Click()
If estadoAbm = 3 Then

    txt_durCbio.Enabled = False
    txt_semCbio.Enabled = False
    cmb_tpo.Enabled = False
    
    txt_durCbio.Text = 0
    txt_semCbio.Text = 0
    cmb_tpo.Text = "Sin Cambios"
    
    dbdiet.Execute " update histclinicas set cambiosIngDiet = false where legajo = " & DataCombo1.BoundText

End If

End Sub

Private Sub Opt_Sindisfuncion_Click()
If estadoAbm = 3 Then

    txt_durDisf.Enabled = False
    txt_semDisf.Enabled = False
    cmb_CapFunc.Enabled = False
    
    txt_durDisf.Text = 0
    txt_semDisf.Text = 0
    cmb_CapFunc.Text = "Sin Disfuncion"
    
    dbdiet.Execute " update histclinicas set disfuncion = false where legajo = " & DataCombo1.BoundText

End If

End Sub

Private Sub Opt_SinEstres_Click()
If estadoAbm = 3 Then

    cmb_NvlEstres.Enabled = False
    
    cmb_NvlEstres.Text = "Sin Estres"
    
    dbdiet.Execute " update histclinicas set estres = false where legajo = " & DataCombo1.BoundText

End If

End Sub

Private Sub Pic_Imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Imprimir

End Sub

Private Sub Pic_Tipito_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Tipito

End Sub

Private Sub txt_durCbio_Validate(Cancel As Boolean)
If Val(txt_durCbio.Text) < 0 Then
    MsgBox "El valor debe ser mayor a cero"
    Cancel = True 'no pierde el enfoque
End If

End Sub

Private Sub txt_durDisf_Validate(Cancel As Boolean)

If Val(txt_durDisf.Text) < 0 Then
    MsgBox "El valor debe ser mayor a cero"
    Cancel = True 'no pierde el enfoque
End If

End Sub

Private Sub txt_idHistClinica_Change()
Dim var As Integer

var = Data1.Recordset.Fields("cambiosIngDiet").Value

If var = -1 Then
    Opt_ConCambios.Value = True
Else
    Opt_SinCambios.Value = True
End If

var = Data1.Recordset.Fields("disfuncion").Value

If var = -1 Then
    Opt_Condisfuncion.Value = True
Else
    Opt_Sindisfuncion.Value = True
End If

var = Data1.Recordset.Fields("estres").Value

If var = -1 Then
    Opt_ConEstres.Value = True
Else
    Opt_SinEstres.Value = True
End If



End Sub

Private Sub txt_pdaSeisMeses_Change()

If DataCombo1.BoundText <> "" Then
    Peso = 0
    
    Set tb = dbdiet.OpenRecordset("pacientes", dbOpenDynaset)
    tb.FindFirst " legajo = " & DataCombo1.BoundText
    Peso = tb.Fields("peso").Value
    tb.Close
    
    txt_porcent.Text = Val(txt_pdaSeisMeses.Text) * 100 / Peso
End If

End Sub

Private Sub txt_pdaSeisMeses_Validate(Cancel As Boolean)


If Val(txt_pdaSeisMeses.Text) < 0 Or Val(txt_pdaSeisMeses.Text) > Peso Then
    MsgBox "El valor debe ser mayor a cero y menor al peso del paciente"
    Cancel = True 'no pierde el enfoque
End If


End Sub

Private Sub txt_porcent_Change()

txt_porcent.Text = Format(txt_porcent.Text, "standard")

End Sub

Private Sub txt_porcent_Validate(Cancel As Boolean)

If Val(txt_porcent.Text) < 0 Or Val(txt_porcent.Text) > 100 Then
    MsgBox "Se debe ingresar un valor entre cero y cien (0-100)"
    Cancel = True 'no pierde el enfoque
End If


End Sub

Private Sub txt_semCbio_Validate(Cancel As Boolean)

If Val(txt_semCbio.Text) < 0 Then
    MsgBox "El valor debe ser mayor a cero"
    Cancel = True 'no pierde el enfoque
End If

End Sub

Private Sub txt_semDisf_Validate(Cancel As Boolean)

If Val(txt_semDisf.Text) < 0 Then
    MsgBox "El valor debe ser mayor a cero"
    Cancel = True 'no pierde el enfoque
End If

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Data1.Recordset = Nothing
Set Me.Adodc1.Recordset = Nothing
Set Me.Adodc2.Recordset = Nothing

'strquery = "select * from histclinicas where legajo = 1"
strQuery = "select * from histclinicas"
Call f_Data_DatabaseName(Data1, strQuery)

'strquery = "ConsultaPrueba3"
strQuery = "select * from pacientes"
Call f_Adodc_ConnectionString(Adodc1, strQuery)

strQuery = "select *, (apell & ', ' & nombre) as nom from pacientes order by apell, nombre"
Call f_Adodc_ConnectionString(Adodc2, strQuery)

'Set DataCombo1.DataSource = Adodc1
'Set DataCombo1.RowSource = Adodc2
'
'DataCombo1.DataField = "legajo"
'DataCombo1.BoundColumn = "legajo"
'DataCombo1.ListField = "nom"
'DataCombo1.Refresh

'Define propiedades de los controles enlazados
'Call f_Enlaza_ControlData(DataCombo1, Adodc1, Adodc2, "Legajo", "Legajo", "nom")
Call f_Enlaza_ControlData(DataCombo1, Adodc2, Adodc2, "Legajo", "Legajo", "nom")
'==============================================

End Sub


Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call f_Aceptar

End Sub


Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cancelar

End Sub


Private Sub Pic_Modificar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Modificar

End Sub

Sub f_Boton_Zorder()

If Me.cmd_Tipito.Enabled = True Then
    Me.Pic_Tipito.ZOrder 0
Else
    Me.Pic_Tipito_Gris.ZOrder 0
End If

If Me.cmdImprimir.Enabled = True Then
    Me.Pic_Imprimir.ZOrder 0
Else
    Me.Pic_Imprimir_Gris.ZOrder 0
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
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Sub f_Modificar()

Me.cmd_Tipito.ZOrder 1
Me.cmdImprimir.ZOrder 1
Me.cmdModificar.ZOrder 0
Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Sub f_Aceptar()

Me.cmd_Tipito.ZOrder 1
Me.cmdImprimir.ZOrder 1
Me.cmdAceptar.ZOrder 0
Me.cmdModificar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Sub f_Cancelar()

Me.cmd_Tipito.ZOrder 1
Me.cmdImprimir.ZOrder 1
Me.cmdCancelar.ZOrder 0
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1

End Sub

Sub f_Imprimir()

Me.cmd_Tipito.ZOrder 1
Me.cmdImprimir.ZOrder 0
Me.cmdCancelar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1

End Sub

Sub f_Tipito()

Me.cmd_Tipito.ZOrder 0
Me.cmdImprimir.ZOrder 1
Me.cmdCancelar.ZOrder 1
Me.cmdModificar.ZOrder 1
Me.cmdAceptar.ZOrder 1


End Sub
