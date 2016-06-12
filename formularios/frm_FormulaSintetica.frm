VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_FormulaSintetica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fórmula Sintética"
   ClientHeight    =   5760
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7740
   ControlBox      =   0   'False
   Icon            =   "frm_FormulaSintetica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7740
   Begin VB.Frame frame10 
      Height          =   5175
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7695
      Begin VB.Frame fme_vct 
         BorderStyle     =   0  'None
         Caption         =   "Valor Calórico Total"
         Height          =   4455
         Left            =   2760
         TabIndex        =   56
         Top             =   480
         Width           =   4695
         Begin VB.CommandButton vct 
            Caption         =   "VCT"
            Height          =   375
            Left            =   3360
            TabIndex        =   9
            ToolTipText     =   "Valor calórico total"
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            Caption         =   "Factor de Actividad"
            Height          =   1695
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1815
            Begin VB.OptionButton Opt_Moderado 
               Caption         =   "Moderado"
               Height          =   195
               Left            =   360
               TabIndex        =   59
               Top             =   720
               Width           =   1335
            End
            Begin VB.OptionButton Opt_Leve 
               Caption         =   "Leve"
               Height          =   195
               Left            =   360
               TabIndex        =   58
               Top             =   360
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton Opt_Intenso 
               Caption         =   "Intenso"
               Height          =   195
               Left            =   360
               TabIndex        =   57
               Top             =   1080
               Width           =   1215
            End
         End
         Begin VB.TextBox rebtxt 
            DataField       =   "reb"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   7
            Text            =   " "
            ToolTipText     =   "Requerimiento energético basal"
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox rcttxt 
            DataField       =   "rct"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   8
            Text            =   " "
            ToolTipText     =   "Requerimiento calórico total"
            Top             =   3000
            Width           =   1455
         End
         Begin MSDataListLib.DataList DataList2 
            Bindings        =   "frm_FormulaSintetica.frx":0ECA
            DataField       =   "idFi"
            DataSource      =   "Data1"
            Height          =   1425
            Left            =   2160
            TabIndex        =   6
            Top             =   600
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2514
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "descripFi"
            BoundColumn     =   "idFi"
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "frm_FormulaSintetica.frx":0EDF
            DataField       =   "idFi"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   2160
            TabIndex        =   60
            Top             =   600
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "descripFi"
            BoundColumn     =   "idFi"
            Text            =   "DataCombo2"
         End
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   330
            Left            =   4440
            Top             =   3960
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
            Caption         =   "Adodc3"
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
         Begin VB.Label Label7 
            Caption         =   "REB"
            Height          =   255
            Left            =   3120
            TabIndex        =   63
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "RCT"
            Height          =   255
            Left            =   3120
            TabIndex        =   62
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Factor de Injuria"
            Height          =   255
            Left            =   2160
            TabIndex        =   61
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame fme_macronutrientes 
         BorderStyle     =   0  'None
         Caption         =   "Cálculo de  proteínas, Hidratos de Carbono y Lípidos"
         Height          =   4455
         Left            =   3120
         TabIndex        =   22
         Top             =   480
         Width           =   4095
         Begin VB.TextBox grprottxt 
            DataField       =   "grProtxKg"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   2640
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox PorcLipTxt 
            DataField       =   "PorcentLip"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   2640
            TabIndex        =   14
            Text            =   " "
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton Calcularcmd 
            Caption         =   "Calcular"
            Height          =   375
            Left            =   2880
            TabIndex        =   15
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Frame Frame7 
            Caption         =   "Informe Final del  VCT   ( "
            Height          =   2655
            Left            =   360
            TabIndex        =   23
            Top             =   1800
            Width           =   3735
            Begin VB.TextBox prot1txt 
               DataField       =   "ProtPorcent"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   240
               TabIndex        =   32
               Text            =   " "
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox prot2txt 
               DataField       =   "ProtKcal"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               TabIndex        =   31
               Text            =   " "
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox prot3txt 
               DataField       =   "ProtG"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   2640
               TabIndex        =   30
               Text            =   " "
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox hc1txt 
               DataField       =   "HCPorcent"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   240
               TabIndex        =   29
               Text            =   " "
               Top             =   1440
               Width           =   615
            End
            Begin VB.TextBox hc2txt 
               DataField       =   "HCKcal"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               TabIndex        =   28
               Text            =   " "
               Top             =   1440
               Width           =   615
            End
            Begin VB.TextBox hc3txt 
               DataField       =   "HCG"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   2640
               TabIndex        =   27
               Text            =   " "
               Top             =   1440
               Width           =   615
            End
            Begin VB.TextBox lip1txt 
               DataField       =   "LipPorcent"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   240
               TabIndex        =   26
               Text            =   " "
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox lip2txt 
               DataField       =   "LipKcal"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               TabIndex        =   25
               Text            =   " "
               Top             =   2280
               Width           =   615
            End
            Begin VB.TextBox lip3txt 
               DataField       =   "LipG"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   2640
               TabIndex        =   24
               Text            =   " "
               Top             =   2280
               Width           =   615
            End
            Begin VB.Label lblRctideal 
               Alignment       =   1  'Right Justify
               Caption         =   "Label26"
               DataField       =   "rctideal"
               DataSource      =   "Data1"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   1920
               TabIndex        =   33
               Top             =   0
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "Proteínas"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label Label15 
               Caption         =   "Hidratos de Carbono"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   1200
               Width           =   3375
            End
            Begin VB.Label Label16 
               Caption         =   "Lípidos"
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   2040
               Width           =   3375
            End
            Begin VB.Label Label17 
               Caption         =   "%"
               Height          =   255
               Left            =   960
               TabIndex        =   43
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label18 
               Caption         =   "Kca"
               Height          =   255
               Left            =   2160
               TabIndex        =   42
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label19 
               Caption         =   "%"
               Height          =   255
               Left            =   960
               TabIndex        =   41
               Top             =   1440
               Width           =   375
            End
            Begin VB.Label Label20 
               Caption         =   "g"
               Height          =   255
               Left            =   3360
               TabIndex        =   40
               Top             =   600
               Width           =   135
            End
            Begin VB.Label Label21 
               Caption         =   "Kca"
               Height          =   255
               Left            =   2160
               TabIndex        =   39
               Top             =   1440
               Width           =   375
            End
            Begin VB.Label Label22 
               Caption         =   "g"
               Height          =   255
               Left            =   3360
               TabIndex        =   38
               Top             =   1440
               Width           =   135
            End
            Begin VB.Label Label23 
               Caption         =   "%"
               Height          =   255
               Left            =   960
               TabIndex        =   37
               Top             =   2280
               Width           =   375
            End
            Begin VB.Label Label24 
               Caption         =   "Kca"
               Height          =   255
               Left            =   2160
               TabIndex        =   36
               Top             =   2280
               Width           =   375
            End
            Begin VB.Label Label25 
               Caption         =   "g"
               Height          =   255
               Left            =   3360
               TabIndex        =   35
               Top             =   2280
               Width           =   135
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               Caption         =   "Kcal.)"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3000
               TabIndex        =   34
               Top             =   0
               Width           =   600
            End
         End
         Begin VB.Label Label9 
            Caption         =   "  Ingresar  gr.  prot./Kg/día"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label13 
            Caption         =   "Ingresar % de Lípidos deseados"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   840
            Width           =   2295
         End
      End
      Begin VB.Frame fme_imc 
         BorderStyle     =   0  'None
         Caption         =   "Índice de Masa Corporal"
         Height          =   4335
         Left            =   3120
         TabIndex        =   49
         Top             =   600
         Width           =   4335
         Begin VB.Frame Frame6 
            Caption         =   "Calcular Peso Ideal"
            Height          =   1695
            Left            =   240
            TabIndex        =   51
            Top             =   1200
            Width           =   3975
            Begin VB.TextBox pesoidealtxt 
               DataField       =   "PesoIdeal"
               DataSource      =   "Data1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   2400
               TabIndex        =   52
               Text            =   " "
               Top             =   1200
               Width           =   1455
            End
            Begin VB.TextBox imcidealtxt 
               DataField       =   "imcIdeal"
               DataSource      =   "Data1"
               Height          =   285
               Left            =   2400
               TabIndex        =   11
               Text            =   " "
               Top             =   720
               Width           =   1455
            End
            Begin VB.CommandButton pesoidealcmd 
               Caption         =   "Calcular"
               Height          =   375
               Left            =   2640
               TabIndex        =   12
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label11 
               Caption         =   "Ingresar IMC ideal"
               Height          =   255
               Left            =   600
               TabIndex        =   54
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label12 
               Caption         =   "Peso Ideal (kg)"
               Height          =   255
               Left            =   600
               TabIndex        =   53
               Top             =   1200
               Width           =   1455
            End
         End
         Begin VB.TextBox imctxt 
            DataField       =   "imc"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   50
            Text            =   " "
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton imccmd 
            Caption         =   "IMC"
            Height          =   375
            Left            =   2880
            TabIndex        =   10
            ToolTipText     =   "Índice de masa corporal"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label imclbl 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            DataField       =   "estImc"
            DataSource      =   "Data1"
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   240
            TabIndex        =   55
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.TextBox Text1 
         DataField       =   "Legajo"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2280
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame9 
         Height          =   1695
         Left            =   120
         TabIndex        =   66
         Top             =   3360
         Width           =   2295
         Begin VB.TextBox edadtxt 
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox tallatxt 
            DataField       =   "Talla"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Text            =   " "
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox pesotxt 
            DataField       =   "Peso"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1080
            TabIndex        =   2
            Text            =   " "
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Edad (años)"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Talla (cm)"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Peso (kg)"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Nombre:"
         Height          =   855
         Left            =   120
         TabIndex        =   64
         Top             =   2520
         Visible         =   0   'False
         Width           =   2295
         Begin VB.PictureBox Pic_Tipito_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1920
            MouseIcon       =   "frm_FormulaSintetica.frx":0EF4
            Picture         =   "frm_FormulaSintetica.frx":1046
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   109
            Top             =   360
            Width           =   315
         End
         Begin VB.PictureBox Pic_Tipito 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1920
            MouseIcon       =   "frm_FormulaSintetica.frx":1176
            Picture         =   "frm_FormulaSintetica.frx":12C8
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   108
            Top             =   360
            Width           =   315
         End
         Begin VB.CommandButton cmd_Tipito 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_FormulaSintetica.frx":1558
            Height          =   315
            Left            =   1920
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_FormulaSintetica.frx":1C68
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Agregar"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frm_FormulaSintetica.frx":1EF8
            DataField       =   "Legajo"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "nom"
            BoundColumn     =   "Legajo"
            Text            =   "DataCombo1"
         End
      End
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "frm_FormulaSintetica.frx":1F0C
         DataField       =   "Legajo"
         DataSource      =   "Data1"
         Height          =   2790
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   4921
         _Version        =   393216
         Appearance      =   0
         ListField       =   "nom"
         BoundColumn     =   "Legajo"
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   4935
         Left            =   2520
         TabIndex        =   71
         Top             =   120
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   8705
         ShowTips        =   0   'False
         HotTracking     =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Valor Calorico Total"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Indice de Masa Corporal"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Macronutrientes"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label26 
         Caption         =   "Paciente:"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   2175
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "D:\Dietetica\rpts\rep_formsintetica.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   0
      Top             =   5880
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
      Top             =   8160
      Width           =   11775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sexo"
      Height          =   1095
      Left            =   3960
      TabIndex        =   17
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton femopt 
         Caption         =   "Femenino"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton mascopt 
         Caption         =   "Masculino"
         Height          =   195
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame fme_botones_abm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   1103
      TabIndex        =   73
      Top             =   5160
      Width           =   5535
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":1F20
         Height          =   375
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":2079
         Picture         =   "frm_FormulaSintetica.frx":21CB
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Aceptar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":2487
         Height          =   375
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":261B
         Picture         =   "frm_FormulaSintetica.frx":276D
         Style           =   1  'Graphical
         TabIndex        =   100
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
         MouseIcon       =   "frm_FormulaSintetica.frx":2C20
         Picture         =   "frm_FormulaSintetica.frx":2D72
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   99
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
         MouseIcon       =   "frm_FormulaSintetica.frx":3073
         Picture         =   "frm_FormulaSintetica.frx":31C5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   98
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":3481
         Height          =   375
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":35A2
         Picture         =   "frm_FormulaSintetica.frx":36F4
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Modificar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":3967
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":3A7D
         Picture         =   "frm_FormulaSintetica.frx":3BCF
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Eliminar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":3D5E
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":3EAB
         Picture         =   "frm_FormulaSintetica.frx":3FFD
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Agregar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUltimo 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":4437
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":45DF
         Picture         =   "frm_FormulaSintetica.frx":4731
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Ultimo"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSiguiente 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":4BFC
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":4D69
         Picture         =   "frm_FormulaSintetica.frx":4EBB
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Siguiente"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":5330
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":54B8
         Picture         =   "frm_FormulaSintetica.frx":560A
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Buscar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":58E7
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":5A51
         Picture         =   "frm_FormulaSintetica.frx":5BA3
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Anterior"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimero 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":6011
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_FormulaSintetica.frx":61B6
         Picture         =   "frm_FormulaSintetica.frx":6308
         Style           =   1  'Graphical
         TabIndex        =   82
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
         MouseIcon       =   "frm_FormulaSintetica.frx":67C3
         Picture         =   "frm_FormulaSintetica.frx":6915
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   81
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
         MouseIcon       =   "frm_FormulaSintetica.frx":6DD0
         Picture         =   "frm_FormulaSintetica.frx":6F22
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   80
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
         MouseIcon       =   "frm_FormulaSintetica.frx":7390
         Picture         =   "frm_FormulaSintetica.frx":74E2
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   79
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
         MouseIcon       =   "frm_FormulaSintetica.frx":77BF
         Picture         =   "frm_FormulaSintetica.frx":7911
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   78
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
         MouseIcon       =   "frm_FormulaSintetica.frx":7D86
         Picture         =   "frm_FormulaSintetica.frx":7ED8
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   77
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
         MouseIcon       =   "frm_FormulaSintetica.frx":83A3
         Picture         =   "frm_FormulaSintetica.frx":84F5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   76
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
         MouseIcon       =   "frm_FormulaSintetica.frx":892F
         Picture         =   "frm_FormulaSintetica.frx":8A81
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   75
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
         MouseIcon       =   "frm_FormulaSintetica.frx":8C10
         Picture         =   "frm_FormulaSintetica.frx":8D62
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   74
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
         MouseIcon       =   "frm_FormulaSintetica.frx":8FD5
         Picture         =   "frm_FormulaSintetica.frx":9127
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   90
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
         MouseIcon       =   "frm_FormulaSintetica.frx":9248
         Picture         =   "frm_FormulaSintetica.frx":939A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   91
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
         MouseIcon       =   "frm_FormulaSintetica.frx":94B0
         Picture         =   "frm_FormulaSintetica.frx":9602
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   92
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
         MouseIcon       =   "frm_FormulaSintetica.frx":974F
         Picture         =   "frm_FormulaSintetica.frx":98A1
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   93
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
         MouseIcon       =   "frm_FormulaSintetica.frx":9A49
         Picture         =   "frm_FormulaSintetica.frx":9B9B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   94
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
         MouseIcon       =   "frm_FormulaSintetica.frx":9D08
         Picture         =   "frm_FormulaSintetica.frx":9E5A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   95
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
         MouseIcon       =   "frm_FormulaSintetica.frx":9FE2
         Picture         =   "frm_FormulaSintetica.frx":A134
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   96
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
         MouseIcon       =   "frm_FormulaSintetica.frx":A29E
         Picture         =   "frm_FormulaSintetica.frx":A3F0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   97
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
         MouseIcon       =   "frm_FormulaSintetica.frx":A595
         Picture         =   "frm_FormulaSintetica.frx":A6E7
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   102
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
         MouseIcon       =   "frm_FormulaSintetica.frx":A840
         Picture         =   "frm_FormulaSintetica.frx":A992
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   103
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_FormulaSintetica.frx":AB26
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_FormulaSintetica.frx":AC7E
         Style           =   1  'Graphical
         TabIndex        =   105
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
         Left            =   0
         MouseIcon       =   "frm_FormulaSintetica.frx":B0FE
         Picture         =   "frm_FormulaSintetica.frx":B250
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   104
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Imprimir_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         MouseIcon       =   "frm_FormulaSintetica.frx":B6D0
         Picture         =   "frm_FormulaSintetica.frx":B822
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   106
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Ingrese el IMC ideal "
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "RCT"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "REB"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   5760
      Width           =   855
   End
End
Attribute VB_Name = "frm_FormulaSintetica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
Dim nLegajo, idFi As Integer

Private Sub Calcularcmd_Click()
Dim GrProtpeso, Proteinas, Kcaprot, porcentprot, rebtxtAux, rcttxtAux
Dim idFa As Integer

'===========================
'valores Factor de Actividad
If Opt_Leve Then

    idFa = 3
    
Else
    
    If Opt_Moderado Then
    
        idFa = 4
    
    Else
    
        idFa = 5
        
    End If
    
End If
'===========================


Set tb = dbdiet.OpenRecordset("fa", dbOpenDynaset)
tb.FindFirst " idfa = " & idFa

If mascopt.Value = True Then
  a = 66.47
  b = 13.75
  c = 5
  d = 6.75
  
  'valores por factor de actividad masculino
  fa = tb.Fields("valorfamasc").Value
  
Else
' valores por sexo femenino para REB
  a = 655.1
  b = 9.56
  c = 1.85
  d = 4.68
  
  'valores por factor de actividad femenino
  fa = tb.Fields("valorfafem").Value
End If

rebtxtAux = a + b * Val(pesoidealtxt.Text) + c * Val(tallatxt.Text) - d * Val(edadtxt.Text)

tb.Close

Set tb = dbdiet.OpenRecordset("fi", dbOpenDynaset)
tb.FindFirst " idfi = " & idFi 'DataCombo2.BoundText

fi = tb.Fields("valorfi").Value

If fi <> 0 Then
  rcttxtAux = Val(rebtxtAux) * fa * fi 'Val(rebtxt.Text) * fa * fi
Else
  rcttxtAux = Val(rebtxtAux) * fa 'Val(rebtxt.Text) * fa
End If
tb.Close

'dbdiet.Execute " update pacientes set rctideal = " & rcttxtAux & " where legajo = " & DataList1.BoundText 'DataCombo1.BoundText
lblRctideal = rcttxtAux

'lblRctideal.Caption = rcttxtAux

GrProtpeso = Val(grprottxt.Text) * Val(pesoidealtxt.Text)
Kcaprot = GrProtpeso * 4
porcentprot = Kcaprot * 100 / rcttxtAux 'Val(rcttxt.Text)
prot1txt = porcentprot
prot2txt = Kcaprot
prot3txt = GrProtpeso

PorcentHc = (100 - Val(PorcLipTxt.Text)) - porcentprot
KcaHc = PorcentHc * rcttxtAux / 100 'Val(rcttxt.Text) / 100
GrHc = KcaHc / 4
hc1txt = PorcentHc
hc2txt = KcaHc
hc3txt = GrHc

Kcalip = Val(PorcLipTxt.Text) * rcttxtAux / 100 'Val(rcttxt.Text) / 100
Grlip = Kcalip / 9
lip1txt = Val(PorcLipTxt.Text)
lip2txt = Kcalip
lip3txt = Grlip

'Data1.Refresh

End Sub



Private Sub cmdImprimir_Click()
Dim strQuery As String

'Resets the value of all properties (except DataSource Property) to their default values.
CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_formSintetica_one.rpt"
              
strQuery = " {pacientes.legajo} = " & nLegajo

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub

Private Sub cmd_tipito_Click()
Unload frmPacientes
frmPacientes.Show
frmPacientes.Data1.Recordset.FindFirst " legajo = " & nLegajo 'DataCombo1.BoundText

End Sub


Private Sub DataCombo1_Click(Area As Integer)
Dim strQuery As String
If DataCombo1.Text <> "" Then
    strQuery = " select * from pacientes where legajo = " & DataCombo1.BoundText

    With Data1
        .RecordSource = strQuery
        .Refresh
    End With
End If

End Sub

Private Sub DataCombo1_LostFocus()
If DataCombo1.Text = "" Then
    DataCombo1.SetFocus
    MsgBox "Debe Completar el Nombre del Paciente", vbInformation, "Información"
End If

End Sub

Private Sub DataCombo2_LostFocus()
If DataCombo2.Text = "" Then
    DataCombo2.SetFocus
    MsgBox "Debe Completar el Factor de Injuria", vbInformation, "Información"
End If

End Sub

Private Sub DataList1_Click()
Dim strQuery As String

If DataList1.Text <> "" Then
          
    nLegajo = DataList1.BoundText
    
    strQuery = " select * from pacientes where legajo = " & nLegajo

    With Data1
        .RecordSource = strQuery
        .Refresh
    End With
    
    idFi = DataList2.BoundText
    
    Me.Caption = " Fórmula Sintética " & " - " & DataList1.Text
End If

End Sub

Private Sub DataList1_LostFocus()

If DataList1.Text = "" Then
   
    DataList1.SetFocus
    MsgBox "Debe seleccionar un Paciente", vbInformation, "Información"

End If

End Sub

Private Sub DataList2_Click()

idFi = DataList2.BoundText

End Sub

Private Sub DataList2_LostFocus()
If DataList2.Text = "" Then
    DataList2.SetFocus
    MsgBox "Debe Completar el Factor de Injuria", vbInformation, "Información"
End If

End Sub

Private Sub edadtxt_Change()
''If Val(pesotxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 Then
''    vct.Enabled = True
''    imccmd.Enabled = True
''    Else
''    vct.Enabled = False
''    imccmd.Enabled = False
''End If
''
''If Val(grprottxt.Text) > 0 And Val(PorcLipTxt.Text) > 0 And Val(PorcLipTxt.Text) < 101 And Val(pesoidealtxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 And Val(rebtxt.Text) > 0 Then
''    Calcularcmd.Enabled = True
''    Else
''    Calcularcmd.Enabled = False
''End If

Call f_BotonEnabled

End Sub

Private Sub edadtxt_GotFocus()
edadtxt.SelStart = 0
edadtxt.SelLength = 50

End Sub

Private Sub edadtxt_KeyUp(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''    If vct.Enabled = True Then
''        'vct.Value = True
''        Call vct_Click
''    End If
''    If imccmd.Enabled = True Then
''        'imccmd.Value = True
''        Call imccmd_Click
''
''    End If
''    If pesoidealcmd.Enabled = True Then
''        'pesoidealcmd.Value = True
''        Call pesoidealcmd_Click
''
''    End If
''    If Calcularcmd.Enabled = True Then
''        'Calcularcmd.Value = True
''        Call Calcularcmd_Click
''
''    End If
''End If

End Sub

Private Sub edadtxt_LostFocus()

Call f_ejecutaCalculos

End Sub

Private Sub femopt_Click()
'dbdiet.Execute " update pacientes set idsexo = 2 where legajo = " & DataList1.BoundText 'DataCombo1.BoundText

End Sub

Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 6165
Me.Width = 7830
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

If DataList1.Text <> "" Then
    
    Me.Caption = " Fórmula Sintética " & " - " & DataList1.Text
    
End If

If estadoAbm = 2 Or estadoAbm = 3 Then

    Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

Else
    
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

End If

Me.Adodc.Refresh
Me.Adodc3.Refresh

End Sub

Private Sub Form_Load()
Dim strQuery As String

nLegajo = 0
idFi = 0

Call f_CargarOrigenDatos

'strquery = "select * from pacientes order by apell, nombre"
'
'With Data1
'    .RecordSource = strquery
'    .Refresh
'End With

'Adodc.ConnectionString = "FILE NAME=" & App.Path & "\Alimentos anterior sin replica.UDL"
'Adodc3.ConnectionString = "FILE NAME=" & App.Path & "\Alimentos anterior sin replica.UDL"
'Data1.DatabaseName = App.Path & "\db1nueva prueba anterior sin replica.mdb"

vct.Enabled = False
imccmd.Enabled = False
pesoidealcmd.Enabled = False
Calcularcmd.Enabled = False

'define los valores de los frame correspondientes para que funcionen con el tabstrip
Call TabStrip1_Click

estadoAbm = 1


Call f_Boton_Zorder

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdCancelar_Click

End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub grprottxt_Change()

Call f_BotonEnabled

''If estadoAbm = 3 Then
''
''    If Val(grprottxt.Text) > 0 And Val(PorcLipTxt.Text) > 0 And Val(PorcLipTxt.Text) < 101 And Val(pesoidealtxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 And Val(rebtxt.Text) > 0 Then
''        Calcularcmd.Enabled = True
''        Else
''        Calcularcmd.Enabled = False
''    End If
''
''End If

End Sub

Private Sub grprottxt_GotFocus()
Calcularcmd.Default = True
grprottxt.SelStart = 0
grprottxt.SelLength = 50

End Sub

Private Sub hc2txt_Change()
hc2txt.Text = Format(hc2txt.Text, "standard")
End Sub

Private Sub imccmd_Click()
imctxt.Text = Val(pesotxt.Text) / ((Val(tallatxt.Text) / 100) ^ 2)

If Val(edadtxt.Text) < 19 Then
    imclbl.Caption = "Al ser menor de 19 años no se indican parámetros"
End If

If Val(edadtxt.Text) >= 19 And Val(edadtxt.Text) <= 24 Then
    If Val(imctxt.Text) >= 19 And Val(imctxt.Text) <= 24 Then
        imclbl.Caption = "El paciente se encuentra dentro de los parámetros normales (19-24)"
        Else
        imclbl.Caption = "El paciente se encuentra fuera de los parámetros normales (19-24)"
    End If
End If

If Val(edadtxt.Text) >= 25 And Val(edadtxt.Text) <= 34 Then
    If Val(imctxt.Text) >= 20 And Val(imctxt.Text) <= 25 Then
        imclbl.Caption = "El paciente se encuentra dentro de los parámetros normales (20-25)"
        Else
        imclbl.Caption = "El paciente se encuentra fuera de los parámetros normales (20-25)"
    End If
End If

If Val(edadtxt.Text) >= 35 And Val(edadtxt.Text) <= 44 Then
    If Val(imctxt.Text) >= 21 And Val(imctxt.Text) <= 26 Then
        imclbl.Caption = "El paciente se encuentra dentro de los parámetros normales (21-26)"
        Else
        imclbl.Caption = "El paciente se encuentra fuera de los parámetros normales (21-26)"
    End If
End If

If Val(edadtxt.Text) >= 45 And Val(edadtxt.Text) <= 54 Then
    If Val(imctxt.Text) >= 22 And Val(imctxt.Text) <= 27 Then
        imclbl.Caption = "El paciente se encuentra dentro de los parámetros normales (22-27)"
        Else
        imclbl.Caption = "El paciente se encuentra fuera de los parámetros normales (22-27)"
    End If
End If

If Val(edadtxt.Text) >= 55 And Val(edadtxt.Text) <= 64 Then
    If Val(imctxt.Text) >= 23 And Val(imctxt.Text) <= 28 Then
        imclbl.Caption = "El paciente se encuentra dentro de los parámetros normales (23-28)"
        Else
        imclbl.Caption = "El paciente se encuentra fuera de los parámetros normales (23-28)"
    End If
End If

If Val(edadtxt.Text) >= 65 Then
    If Val(imctxt.Text) >= 24 And Val(imctxt.Text) <= 29 Then
        imclbl.Caption = "El paciente se encuentra dentro de los parámetros normales (24-29)"
        Else
        imclbl.Caption = "El paciente se encuentra fuera de los parámetros normales (24-29)"
    End If
End If

 
 
End Sub

Private Sub imcidealtxt_Change()

Call f_BotonEnabled

''If estadoAbm = 3 Then
''
''    If Val(tallatxt.Text) > 0 And Val(imcidealtxt.Text) > 0 Then
''        pesoidealcmd.Enabled = True
''        Else
''        pesoidealcmd.Enabled = False
''    End If
''
''End If

End Sub

Private Sub imcidealtxt_GotFocus()
pesoidealcmd.Default = True
imcidealtxt.SelStart = 0
imcidealtxt.SelLength = 50

End Sub

Private Sub imctxt_Change()

imctxt.Text = Format(imctxt.Text, "standard")

End Sub

Private Sub lblRctideal_Change()
lblRctideal.Caption = Format(lblRctideal.Caption, "standard")
End Sub

Private Sub lip2txt_Change()
lip2txt.Text = Format(lip2txt.Text, "standard")
End Sub

Private Sub mascopt_Click()

'dbdiet.Execute " update pacientes set idsexo = 1 where legajo = " & DataList1.BoundText 'DataCombo1.BoundText

End Sub

Private Sub pesoidealcmd_Click()

pesoidealtxt.Text = ((Val(tallatxt.Text) / 100) ^ 2) * Val(imcidealtxt.Text)

End Sub

Private Sub pesoidealtxt_Change()

Call f_BotonEnabled

''If estadoAbm = 3 Then
''
''    If Val(grprottxt.Text) > 0 And Val(PorcLipTxt.Text) > 0 And Val(PorcLipTxt.Text) < 101 And Val(pesoidealtxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 And Val(rebtxt.Text) > 0 Then
''        Calcularcmd.Enabled = True
''        Else
''        Calcularcmd.Enabled = False
''    End If
''
''End If

End Sub

Private Sub pesotxt_Change()
''If Val(pesotxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 Then
''    vct.Enabled = True
''    imccmd.Enabled = True
''    Else
''    vct.Enabled = False
''    imccmd.Enabled = False
''End If

Call f_BotonEnabled

End Sub

Private Sub pesotxt_GotFocus()
pesotxt.SelStart = 0
pesotxt.SelLength = 50

End Sub

Private Sub pesotxt_KeyPress(KeyAscii As Integer)
'''Text1.Text = KeyAscii
''If KeyAscii = 13 Then
''    If vct.Enabled = True Then
''        'vct.Value = True
''        Call vct_Click
''    End If
''    If imccmd.Enabled = True Then
''        'imccmd.Value = True
''        Call imccmd_Click
''
''    End If
''    If pesoidealcmd.Enabled = True Then
''        'pesoidealcmd.Value = True
''        Call pesoidealcmd_Click
''    End If
''    If Calcularcmd.Enabled = True Then
''        'Calcularcmd.Value = True
''        Call Calcularcmd_Click
''    End If
''End If

End Sub

Private Sub pesotxt_LostFocus()

Call f_ejecutaCalculos

End Sub

Private Sub Pic_Imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Imprimir

End Sub

Private Sub Pic_Tipito_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Tipito

End Sub

Private Sub PorcLipTxt_Change()
    
Call f_BotonEnabled

''If estadoAbm = 3 Then
''
''    If Val(grprottxt.Text) > 0 And Val(PorcLipTxt.Text) > 0 And Val(PorcLipTxt.Text) < 101 And Val(pesoidealtxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 And Val(rebtxt.Text) > 0 Then
''        Calcularcmd.Enabled = True
''        Else
''        Calcularcmd.Enabled = False
''    End If
''
''End If

End Sub

Private Sub prothcgr_Click()
Load frm_formulaDesarrollada
End Sub

Private Sub salir_Click()
End
End Sub

Private Sub PorcLipTxt_GotFocus()
Calcularcmd.Default = True
PorcLipTxt.SelStart = 0
PorcLipTxt.SelLength = 50

End Sub

Private Sub prot2txt_Change()
prot2txt.Text = Format(prot2txt.Text, "standard")
End Sub

Private Sub rcttxt_Change()

rcttxt.Text = Format(rcttxt.Text, "standard")

End Sub

Private Sub rebtxt_Change()

''If estadoAbm = 3 Then
''
''    If Val(grprottxt.Text) > 0 And Val(PorcLipTxt.Text) > 0 And Val(PorcLipTxt.Text) < 101 And Val(pesoidealtxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 And Val(rebtxt.Text) > 0 Then
''        Calcularcmd.Enabled = True
''        Else
''        Calcularcmd.Enabled = False
''    End If
''
''End If

Call f_BotonEnabled

'rebtxt.Text = Format(rebtxt.Text, "standard")

End Sub

Private Sub TabStrip1_Click()
'define los valores de los frame correspondientes para que funcionen con el tabstrip
Dim a As String
a = TabStrip1.SelectedItem.Index

Select Case a
    Case Is = 1
    
        fme_vct.ZOrder 0
        fme_imc.ZOrder 1
        fme_Macronutrientes.ZOrder 1
        
    Case Is = 2
    
        fme_vct.ZOrder 1
        fme_imc.ZOrder 0
        fme_Macronutrientes.ZOrder 1
        
    Case Is = 3
    
        fme_vct.ZOrder 1
        fme_imc.ZOrder 1
        fme_Macronutrientes.ZOrder 0
        
End Select

End Sub

Private Sub tallatxt_Change()
''If Val(pesotxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 Then
''    imccmd.Enabled = True
''    vct.Enabled = True
''    Else
''    imccmd.Enabled = False
''    vct.Enabled = False
''End If
''
''
''If Val(tallatxt.Text) > 0 And Val(imcidealtxt.Text) > 0 Then
''    pesoidealcmd.Enabled = True
''    Else
''    pesoidealcmd.Enabled = False
''End If
''
''If Val(grprottxt.Text) > 0 And Val(PorcLipTxt.Text) > 0 And Val(PorcLipTxt.Text) < 101 And Val(pesoidealtxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 And Val(rebtxt.Text) > 0 Then
''    Calcularcmd.Enabled = True
''    Else
''    Calcularcmd.Enabled = False
''End If

Call f_BotonEnabled

End Sub

Private Sub tallatxt_GotFocus()
tallatxt.SelStart = 0
tallatxt.SelLength = 50

End Sub

Private Sub tallatxt_KeyUp(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''    If vct.Enabled = True Then
''        'vct.Value = True
''        Call vct_Click
''    End If
''    If imccmd.Enabled = True Then
''        'imccmd.Value = True
''        Call imccmd_Click
''
''    End If
''    If pesoidealcmd.Enabled = True Then
''        'pesoidealcmd.Value = True
''        Call pesoidealcmd_Click
''
''    End If
''    If Calcularcmd.Enabled = True Then
''        'Calcularcmd.Value = True
''        Call Calcularcmd_Click
''
''    End If
''End If


End Sub

Private Sub tallatxt_LostFocus()

Call f_ejecutaCalculos

End Sub

Private Sub Text1_Change()
'Set tb = dbdiet.OpenRecordset("
Dim var 'As Integer

'establece que factor de injuria que corresponde en el botón de opción de FA

var = Data1.Recordset.Fields("idfa").Value

Select Case var
    Case Is = 3
        Opt_Leve.Value = True
    Case Is = 4
        Opt_Moderado.Value = True
    Case Is = 5
        Opt_Intenso.Value = True
End Select

'establece sexo que corresponde en el botón de opción de sexo
var = Data1.Recordset.Fields("idsexo").Value

If var = 1 Then
    mascopt.Value = True
Else
    femopt.Value = True
End If

edadtxt.Text = Format(Now() - Data1.Recordset.Fields("fnacimiento").Value, "yy")
End Sub

'Private Sub Text2_Click()
'Text2.Text = nLegajo 'DataCombo1.BoundText
'End Sub

Private Sub txt_idFa_Change()

'Select Case txt_idFa
'    Case Is = 3
'        LeveOpt.Value = True
'    Case Is = 4
'        ModeradoOpt.Value = True
'    Case Is = 5
'        IntensoOpt.Value = True
'End Select

End Sub

Private Sub vct_Click()
Dim idFa As Long
Dim fa, fi

rebtxt.Text = ""
rcttxt.Text = ""

'===========================
'valores Factor de Actividad
If Opt_Leve Then

    idFa = 3
    
Else
    
    If Opt_Moderado Then
    
        idFa = 4
    
    Else
    
        idFa = 5
        
    End If
    
End If
'===========================

Set tb = dbdiet.OpenRecordset("fa", dbOpenDynaset)
tb.FindFirst " idfa = " & idFa

If mascopt.Value = True Then
  a = 66.47
  b = 13.75
  c = 5
  d = 6.75
  
  'valores por factor de actividad masculino
  fa = tb.Fields("valorfamasc").Value
  
Else
' valores por sexo femenino para REB
  a = 655.1
  b = 9.56
  c = 1.85
  d = 4.68
  
  'valores por factor de actividad femenino
  fa = tb.Fields("valorfafem").Value
End If

'calcula el REB
rebtxt.Text = a + b * Val(pesotxt.Text) + c * Val(tallatxt.Text) - d * Val(edadtxt.Text)

tb.Close

Set tb = dbdiet.OpenRecordset("fi", dbOpenDynaset)
tb.FindFirst " idfi = " & idFi 'DataCombo2.BoundText

fi = tb.Fields("valorfi").Value

If fi <> 0 Then
  rcttxt.Text = Val(rebtxt.Text) * fa * fi
Else
  rcttxt.Text = Val(rebtxt.Text) * fa
End If
tb.Close

End Sub


Private Sub cmdAceptar_Click()
Dim sMsg As String

If estadoAbm = 2 Or estadoAbm = 3 Then 'si el estado es agregar o modificar
    
    sMsg = vbYes
    
    If Len(f_validaPendientes) > 0 Then
        
        sMsg = MsgBox("Falta calcular los siguientes campos: " & f_validaPendientes & "¿Esta seguro que desea continuar?", vbYesNo)
            
    End If
    
    If sMsg = vbYes Then
        
        'por las dudas fuerzo el calculo de todas las funciones que se pueda
        Call f_ejecutaCalculos
        
        Data1.UpdateRecord
        MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
        
    '    'condiciones extras
    '    If estadoAbm = 2 Then
    '        dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(MDIForm1.ActiveForm.Label1.Caption) & ", codalimento from alimentos where estado = true"
    '        dbdiet.Execute "insert into histclinicas (legajo) select " & Val(MDIForm1.ActiveForm.Label1.Caption) '& ", codalimento from alimentos where estado = true"
    '    End If
    '
    '    cmdBuscar.Enabled = True
    '    cmdAgregar.Enabled = True
    '    'cmdBorrar.Enabled = True
    '    'cmdClose.Enabled = True
        cmdModificar.Enabled = True
    '
    '    cmdAgregar.SetFocus
        
        cmdAceptar.Default = True
        cmdCancelar.Cancel = True
        
        'cmdPrimero.Enabled = True
        'cmdAnterior.Enabled = True
        'cmdSiguiente.Enabled = True
        'cmdUltimo.Enabled = True
        'Call enabledDesplaz 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados
        Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)
    
        estadoAbm = 1 ' el estado del form es "sin cambios"
        
        DataList1.Enabled = True
        DataList2.Enabled = False
        
        vct.Enabled = False
        imccmd.Enabled = False
        pesoidealcmd.Enabled = False
        Calcularcmd.Enabled = False
            
        Call f_Boton_Zorder
        
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

MDIForm1.ActiveForm.txtFields(1).SetFocus

Unload frm_FormulaSintetica
Unload frm_formulaDesarrollada
Unload frm_Adm_Diet
Unload frm_Evaluacion_Subjetiva

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
        
            strQuery = "select * from alimenxpaciente where legajo = " & Val(Label1.Caption) & " and cantidad <> 0"
                    
            Set tb = dbdiet.OpenRecordset(strQuery)
            strQuery = "select * from menu where legajo = " & Val(Label1.Caption)
            
            Set tb1 = dbdiet.OpenRecordset(strQuery)
            If tb.RecordCount = 0 And tb1.RecordCount = 0 Then
                Data1.Recordset.Delete
                Data1.Recordset.MovePrevious
                dbdiet.Execute "delete from alimenxpaciente where legajo = " & Val(Label1.Caption)
                dbdiet.Execute "delete from menu where legajo = " & Val(Label1.Caption)
                dbdiet.Execute "delete from platosmenu where legajo = " & Val(Label1.Caption)
            Else
                MsgBox "No se puede eliminar el registro actual porque puede afectar la integridad del Sistema", , "Información"
            End If
            tb.Close
            tb1.Close
                            
            Call f_Boton_Zorder
            
    Else
        cmdAgregar.SetFocus
    End If
End If

End Sub

Private Sub cmdBuscar_Click()
Dim strQuery As String

strQuery = " select * from pacientes order by apell, nombre"

With Data1
    .RecordSource = strQuery
    .Refresh
End With

'aclare campo por el cual buscar
msg = InputBox("Ingrese apellido del paciente:", "Buscar por Apellido")
    
If msg <> "" Then

    strQuery = " select * from pacientes where apell like '" & msg & "*' order by apell, nombre"
    
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

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        MDIForm1.ActiveForm.Data1.Recordset.CancelUpdate
    
    End If
    
    'cmdBuscar.Enabled = True
    'cmdAgregar.Enabled = True
    'cmdBorrar.Enabled = True
    'cmdClose.Enabled = True
    cmdModificar.Enabled = True
    
    'cmdAgregar.SetFocus
    cmdAceptar.Default = True
    'cmdClose.Cancel = True
    'cmdPrimero.Enabled = True
    'cmdAnterior.Enabled = True
    'cmdSiguiente.Enabled = True
    'cmdUltimo.Enabled = True
           
    'Call enabledDesplaz 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
    
    DataList1.Enabled = True
    DataList2.Enabled = False
    
    vct.Enabled = False
    imccmd.Enabled = False
    pesoidealcmd.Enabled = False
    Calcularcmd.Enabled = False
    
    Call f_Boton_Zorder
    
Else

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        MDIForm1.ActiveForm.Hide
    
    End If
  
    'por las dudas fuerzo el calculo de todas las funciones que se pueda
    Call f_ejecutaCalculos

End If

End Sub



'Private Sub cmdImprimir_Click()
''aclare el filtro para imprimir
'CrystalReport1.SelectionFormula = " {pacientes.legajo} = " & Val(Label1.Caption) '& " and {platosmenu.fechaMenu} in Date(" & Year(DTdesde.Value) & ", " & Month(DTdesde.Value) & ", " & Day(DTdesde.Value) & ") to Date(" & Year(DThasta.Value) & ", " & Month(DThasta.Value) & ", " & Day(DThasta.Value) & ") "
'
'CrystalReport1.Destination = crptToWindow
'CrystalReport1.PrintReport
'
'End Sub

Private Sub cmdModificar_Click()

Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

pesoidealtxt.Enabled = False
imctxt.Enabled = False
rebtxt.Enabled = False
rcttxt.Enabled = False
prot1txt.Enabled = False
prot2txt.Enabled = False
prot3txt.Enabled = False
hc1txt.Enabled = False
hc2txt.Enabled = False
hc3txt.Enabled = False
lip1txt.Enabled = False
lip2txt.Enabled = False
lip3txt.Enabled = False

DataList1.Enabled = False
DataList2.Enabled = True

'If MDIForm1.ActiveForm.Data1.Recordset.BOF = True Or MDIForm1.ActiveForm.Data1.Recordset.EOF = True Then
'    MDIForm1.ActiveForm.Data1.Recordset.MoveFirst
'End If
'
'cmdAgregar.Enabled = False
'cmdBorrar.Enabled = False
''cmdclose.Enabled = False
cmdModificar.Enabled = False
'cmdBuscar.Enabled = False
cmdAceptar.Visible = True
cmdCancelar.Visible = True
'cmdPrimero.Enabled = False
'cmdAnterior.Enabled = False
'cmdSiguiente.Enabled = False
'cmdUltimo.Enabled = False


Data1.Recordset.Edit

'MDIForm1.ActiveForm.txtFields(1).SetFocus
'
cmdAceptar.Default = True
cmdCancelar.Cancel = True
'
estadoAbm = 3 ' el estado es modificar

Call f_BotonEnabled

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

Function f_validaPendientes() As String

f_validaPendientes = "" '& vbCrLf

If rcttxt.Text = "" Then
    f_validaPendientes = f_validaPendientes & vbTab & " - RCT " & vbCrLf
End If

If rebtxt.Text = "" Then
    f_validaPendientes = f_validaPendientes & vbTab & " - REB " & vbCrLf
End If

If imctxt.Text = "" Then
    f_validaPendientes = f_validaPendientes & vbTab & " - IMC " & vbCrLf
End If

If imcidealtxt.Text = "" Then
    f_validaPendientes = f_validaPendientes & vbTab & " - IMC ideal " & vbCrLf
End If

If pesoidealtxt.Text = "" Then
    f_validaPendientes = f_validaPendientes & vbTab & " - Peso ideal " & vbCrLf
End If

If lblRctideal.Caption = "" Then
    f_validaPendientes = f_validaPendientes & vbTab & " - RCT ideal " & vbCrLf
End If

End Function

Sub f_ejecutaCalculos()

If Val(pesotxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 Then
    Call vct_Click
    Call imccmd_Click
End If

If Val(tallatxt.Text) > 0 And Val(imcidealtxt.Text) > 0 Then
    Call pesoidealcmd_Click
End If

If Val(grprottxt.Text) > 0 And Val(PorcLipTxt.Text) > 0 And Val(PorcLipTxt.Text) < 101 And Val(pesoidealtxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 And Val(rebtxt.Text) > 0 Then
    Call Calcularcmd_Click
End If

''If vct.Enabled = True Then
''    Call vct_Click
''End If
''
''If imccmd.Enabled = True Then
''    Call imccmd_Click
''End If
''
''If pesoidealcmd.Enabled = True Then
''    Call pesoidealcmd_Click
''End If
''
''If Calcularcmd.Enabled = True Then
''    Call Calcularcmd_Click
''End If

End Sub

Sub f_BotonEnabled()

If estadoAbm = 3 Then

    If Val(pesotxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 Then
        imccmd.Enabled = True
        vct.Enabled = True
        Else
        imccmd.Enabled = False
        vct.Enabled = False
    End If
    
    
    If Val(tallatxt.Text) > 0 And Val(imcidealtxt.Text) > 0 Then
        pesoidealcmd.Enabled = True
        Else
        pesoidealcmd.Enabled = False
    End If
    
    If Val(grprottxt.Text) > 0 And Val(PorcLipTxt.Text) > 0 And Val(PorcLipTxt.Text) < 101 And Val(pesoidealtxt.Text) > 0 And Val(tallatxt.Text) > 0 And Val(edadtxt.Text) > 0 And Val(rebtxt.Text) > 0 Then
        Calcularcmd.Enabled = True
        Else
        Calcularcmd.Enabled = False
    End If

End If

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Data1.Recordset = Nothing
Set Me.Adodc.Recordset = Nothing
Set Me.Adodc3.Recordset = Nothing

'strquery = "select * from Pacientes"
strQuery = "select * from pacientes order by apell, nombre"
Call f_Data_DatabaseName(Data1, strQuery)

strQuery = "select *, (apell & ', ' & nombre) as nom from pacientes order by apell, nombre"
Call f_Adodc_ConnectionString(Adodc, strQuery)

strQuery = "select * from fi"
Call f_Adodc_ConnectionString(Adodc3, strQuery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo1, Data1, Adodc, "Legajo", "Legajo", "nom")

Call f_Enlaza_ControlData(DataCombo2, Data1, Adodc3, "idFi", "idFi", "descripFi")

Call f_Enlaza_ControlData(DataList1, Data1, Adodc, "Legajo", "Legajo", "nom")

Call f_Enlaza_ControlData(DataList2, Data1, Adodc3, "idFi", "idFi", "descripFi")
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

Sub f_Tipito()

Me.cmd_Tipito.ZOrder 0

End Sub
