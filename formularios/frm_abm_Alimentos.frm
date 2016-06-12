VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_abm_Alimentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alimentos"
   ClientHeight    =   5430
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6525
   Icon            =   "frm_abm_Alimentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6525
   Begin VB.Frame Frame2 
      Caption         =   "Nro."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Label lbl_CodAlimento 
         Caption         =   "label1"
         DataField       =   "CodAlimento"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   110
      Top             =   4440
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5280
      Top             =   120
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
      Caption         =   "Detalles"
      Height          =   4815
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   6495
      Begin VB.Frame fme_Macronutrientes 
         Caption         =   "Macronutrientes"
         Height          =   1335
         Left            =   360
         TabIndex        =   26
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox txtFields 
            DataField       =   "Lip"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   3
            Left            =   2040
            MaxLength       =   7
            TabIndex        =   7
            Top             =   870
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Prot"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   2
            Left            =   2040
            MaxLength       =   7
            TabIndex        =   6
            Top             =   555
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "HC"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   1
            Left            =   2040
            MaxLength       =   7
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   23
            Left            =   2880
            TabIndex        =   92
            Top             =   570
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Proteínas:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   91
            Top             =   570
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   24
            Left            =   2880
            TabIndex        =   90
            Top             =   900
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   22
            Left            =   2880
            TabIndex        =   89
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Lípidos:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   69
            Top             =   900
            Width           =   570
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Hidratos de Carbono:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   1500
         End
      End
      Begin VB.Frame fme_Minerales 
         Caption         =   "Minerales"
         Height          =   1335
         Left            =   360
         TabIndex        =   79
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox txtFields 
            DataField       =   "Sodio"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   5
            Left            =   720
            MaxLength       =   7
            TabIndex        =   9
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Calcio"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   6
            Left            =   720
            MaxLength       =   7
            TabIndex        =   10
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Hierro"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   7
            Left            =   3720
            MaxLength       =   7
            TabIndex        =   11
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Fosforo"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   8
            Left            =   3720
            MaxLength       =   7
            TabIndex        =   12
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Potasio"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   9
            Left            =   720
            MaxLength       =   7
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "mg."
            Height          =   195
            Index           =   33
            Left            =   4560
            TabIndex        =   102
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "mg."
            Height          =   195
            Index           =   32
            Left            =   4560
            TabIndex        =   101
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "mg."
            Height          =   195
            Index           =   31
            Left            =   1560
            TabIndex        =   100
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "mg."
            Height          =   195
            Index           =   30
            Left            =   1560
            TabIndex        =   99
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "mg."
            Height          =   195
            Index           =   29
            Left            =   1560
            TabIndex        =   98
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Sodio:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   84
            Top             =   600
            Width           =   450
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Calcio:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   83
            Top             =   960
            Width           =   480
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Hierro:"
            Height          =   195
            Index           =   9
            Left            =   3000
            TabIndex        =   82
            Top             =   240
            Width           =   465
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fósforo:"
            Height          =   195
            Index           =   10
            Left            =   3000
            TabIndex        =   81
            Top             =   600
            Width           =   570
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Potasio:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   570
         End
      End
      Begin VB.Frame fme_Otros 
         Caption         =   "Otros"
         Height          =   1335
         Left            =   360
         TabIndex        =   77
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox txtFields 
            DataField       =   "Fibra"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   4
            Left            =   2040
            MaxLength       =   7
            TabIndex        =   23
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Glucosa"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   19
            Left            =   2040
            MaxLength       =   7
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   40
            Left            =   2880
            TabIndex        =   109
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   28
            Left            =   2880
            TabIndex        =   97
            Top             =   600
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fibra:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   390
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Glucosa:"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   630
         End
      End
      Begin VB.Frame fme_Acidos_Grasos 
         Caption         =   "Acidos_Grasos"
         Height          =   1335
         Left            =   360
         TabIndex        =   85
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox txtFields 
            DataField       =   "AGS"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   16
            Left            =   2040
            MaxLength       =   7
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "AGMI"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   17
            Left            =   2040
            MaxLength       =   7
            TabIndex        =   20
            Top             =   555
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "AGPI"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   18
            Left            =   2040
            MaxLength       =   7
            TabIndex        =   21
            Top             =   870
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   26
            Left            =   2880
            TabIndex        =   96
            Top             =   570
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Monoinsaturados:"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   95
            Top             =   570
            Width           =   1260
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   27
            Left            =   2880
            TabIndex        =   94
            Top             =   900
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "gr."
            Height          =   195
            Index           =   25
            Left            =   2880
            TabIndex        =   93
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Saturados:"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Polinsaturados:"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   86
            Top             =   900
            Width           =   1080
         End
      End
      Begin VB.Frame fme_Vitaminas 
         Caption         =   "Vitaminas"
         Height          =   1335
         Left            =   360
         TabIndex        =   70
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox txtFields 
            DataField       =   "VitA"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   10
            Left            =   720
            MaxLength       =   7
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "VitB1"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   11
            Left            =   720
            MaxLength       =   7
            TabIndex        =   14
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "VitB2"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   12
            Left            =   720
            MaxLength       =   7
            TabIndex        =   15
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "VitNiacina"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   13
            Left            =   3720
            MaxLength       =   7
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "VitC"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   14
            Left            =   3720
            MaxLength       =   7
            TabIndex        =   17
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtFields 
            DataField       =   "VitE"
            DataSource      =   "Data1"
            Height          =   285
            Index           =   15
            Left            =   3720
            MaxLength       =   7
            TabIndex        =   18
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "U.I."
            Height          =   195
            Index           =   39
            Left            =   4560
            TabIndex        =   108
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "U.I."
            Height          =   195
            Index           =   38
            Left            =   1560
            TabIndex        =   107
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "mg."
            Height          =   195
            Index           =   37
            Left            =   4560
            TabIndex        =   106
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "mg."
            Height          =   195
            Index           =   36
            Left            =   4560
            TabIndex        =   105
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "mg."
            Height          =   195
            Index           =   35
            Left            =   1560
            TabIndex        =   104
            Top             =   960
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "mg."
            Height          =   195
            Index           =   34
            Left            =   1560
            TabIndex        =   103
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "A:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   150
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "B1:"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   75
            Top             =   600
            Width           =   240
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "B2:"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   74
            Top             =   960
            Width           =   240
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Niacina:"
            Height          =   195
            Index           =   15
            Left            =   3000
            TabIndex        =   73
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "C:"
            Height          =   195
            Index           =   16
            Left            =   3000
            TabIndex        =   72
            Top             =   600
            Width           =   150
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "E:"
            Height          =   195
            Index           =   17
            Left            =   3000
            TabIndex        =   71
            Top             =   960
            Width           =   150
         End
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frm_abm_Alimentos.frx":0ECA
         DataField       =   "idUnidad"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "DescripUnidad"
         BoundColumn     =   "idUnidad"
         Text            =   "DataCombo2"
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   2160
         TabIndex        =   32
         Top             =   240
         Width           =   3375
         Begin VB.CheckBox Check1 
            Caption         =   "Incluir en Tabla de Alimentos"
            DataField       =   "estado"
            DataSource      =   "Data1"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   33
            ToolTipText     =   "Incluir el alimento en la Fórmula Desarrollada"
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmd_Actualizar 
         Caption         =   "Ac&tualizar"
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         ToolTipText     =   "Mostrar Todos"
         Top             =   4440
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_abm_Alimentos.frx":0EDF
         DataField       =   "idCategoria"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Decripcion"
         BoundColumn     =   "idCategoria"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtFields 
         DataField       =   "DescripAlimento"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1320
         Width           =   3375
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   4080
         Top             =   1680
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   661
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
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   1815
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3201
         ShowTips        =   0   'False
         HotTracking     =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   5
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Macronutrientes"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Minerales"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Vitaminas"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Acidos Grasos"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Otros..."
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblLabels 
         Caption         =   "Unidad:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   67
         Top             =   1695
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidades cada 100 grs.   _____________________"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6360
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lblLabels 
         Caption         =   "Descripción:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Grupos de alimentos:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   975
         Width           =   1815
      End
   End
   Begin VB.Frame fme_botones_abm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   495
      TabIndex        =   24
      Top             =   4800
      Width           =   5535
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":0EF4
         Height          =   375
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":104D
         Picture         =   "frm_abm_Alimentos.frx":119F
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Aceptar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":145B
         Height          =   375
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":15EF
         Picture         =   "frm_abm_Alimentos.frx":1741
         Style           =   1  'Graphical
         TabIndex        =   60
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
         MouseIcon       =   "frm_abm_Alimentos.frx":1BF4
         Picture         =   "frm_abm_Alimentos.frx":1D46
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   59
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
         MouseIcon       =   "frm_abm_Alimentos.frx":2047
         Picture         =   "frm_abm_Alimentos.frx":2199
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   58
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":2455
         Height          =   375
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":2576
         Picture         =   "frm_abm_Alimentos.frx":26C8
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Modificar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":293B
         Height          =   375
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":2A51
         Picture         =   "frm_abm_Alimentos.frx":2BA3
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Eliminar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":2D32
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":2E7F
         Picture         =   "frm_abm_Alimentos.frx":2FD1
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Agregar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUltimo 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":340B
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":35B3
         Picture         =   "frm_abm_Alimentos.frx":3705
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ultimo"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSiguiente 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":3BD0
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":3D3D
         Picture         =   "frm_abm_Alimentos.frx":3E8F
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Siguiente"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":4304
         Height          =   375
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":448C
         Picture         =   "frm_abm_Alimentos.frx":45DE
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Buscar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":48BB
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":4A25
         Picture         =   "frm_abm_Alimentos.frx":4B77
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Anterior"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimero 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":4FE5
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_Alimentos.frx":518A
         Picture         =   "frm_abm_Alimentos.frx":52DC
         Style           =   1  'Graphical
         TabIndex        =   42
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
         MouseIcon       =   "frm_abm_Alimentos.frx":5797
         Picture         =   "frm_abm_Alimentos.frx":58E9
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   41
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
         MouseIcon       =   "frm_abm_Alimentos.frx":5DA4
         Picture         =   "frm_abm_Alimentos.frx":5EF6
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   40
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
         MouseIcon       =   "frm_abm_Alimentos.frx":6364
         Picture         =   "frm_abm_Alimentos.frx":64B6
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   39
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
         MouseIcon       =   "frm_abm_Alimentos.frx":6793
         Picture         =   "frm_abm_Alimentos.frx":68E5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   38
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
         MouseIcon       =   "frm_abm_Alimentos.frx":6D5A
         Picture         =   "frm_abm_Alimentos.frx":6EAC
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   37
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
         MouseIcon       =   "frm_abm_Alimentos.frx":7377
         Picture         =   "frm_abm_Alimentos.frx":74C9
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   36
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
         MouseIcon       =   "frm_abm_Alimentos.frx":7903
         Picture         =   "frm_abm_Alimentos.frx":7A55
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   35
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
         MouseIcon       =   "frm_abm_Alimentos.frx":7BE4
         Picture         =   "frm_abm_Alimentos.frx":7D36
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   34
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
         MouseIcon       =   "frm_abm_Alimentos.frx":7FA9
         Picture         =   "frm_abm_Alimentos.frx":80FB
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   50
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
         MouseIcon       =   "frm_abm_Alimentos.frx":821C
         Picture         =   "frm_abm_Alimentos.frx":836E
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   51
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
         MouseIcon       =   "frm_abm_Alimentos.frx":8484
         Picture         =   "frm_abm_Alimentos.frx":85D6
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   52
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
         MouseIcon       =   "frm_abm_Alimentos.frx":8723
         Picture         =   "frm_abm_Alimentos.frx":8875
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   53
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
         MouseIcon       =   "frm_abm_Alimentos.frx":8A1D
         Picture         =   "frm_abm_Alimentos.frx":8B6F
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   54
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
         MouseIcon       =   "frm_abm_Alimentos.frx":8CDC
         Picture         =   "frm_abm_Alimentos.frx":8E2E
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   55
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
         MouseIcon       =   "frm_abm_Alimentos.frx":8FB6
         Picture         =   "frm_abm_Alimentos.frx":9108
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   56
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
         MouseIcon       =   "frm_abm_Alimentos.frx":9272
         Picture         =   "frm_abm_Alimentos.frx":93C4
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   57
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
         MouseIcon       =   "frm_abm_Alimentos.frx":9569
         Picture         =   "frm_abm_Alimentos.frx":96BB
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   62
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
         MouseIcon       =   "frm_abm_Alimentos.frx":9814
         Picture         =   "frm_abm_Alimentos.frx":9966
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   63
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_Alimentos.frx":9AFA
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_abm_Alimentos.frx":9C52
         Style           =   1  'Graphical
         TabIndex        =   65
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
         MouseIcon       =   "frm_abm_Alimentos.frx":A0D2
         Picture         =   "frm_abm_Alimentos.frx":A224
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   64
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
         MouseIcon       =   "frm_abm_Alimentos.frx":A6A4
         Picture         =   "frm_abm_Alimentos.frx":A7F6
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   66
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "D:\Dietetica\Database\db1nueva prueba anterior sin replica.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from alimentos order by idcategoria, descripalimento"
      Top             =   5085
      Visible         =   0   'False
      Width           =   6525
   End
End
Attribute VB_Name = "frm_abm_Alimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msg As String
Dim tb As Recordset
Dim tb1 As Recordset
Dim nCodAlimento As Integer

'Public estadoAbm As Integer ' define el estado de un formulario de abm
'                             1 = sin cambios; 2 = agregar; 3 = modificar
'el modulo "fSetEnableFields(MDIForm1.ActiveForm, vbFalse)" se debe agregar al proyecto
'Dim Titulo As String 'titulo del form
Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
                            
Private Sub cmdAceptar_Click()
Dim nCodAlimento_add As String

nCodAlimento_add = nCodAlimento

If estadoAbm = 2 Or estadoAbm = 3 Then 'si el estado es agregar o modificar

    If DataCombo1.Text = "" Then
        MsgBox "Debe completar la categoría del alimento"
        DataCombo1.SetFocus
    Else
        If txtFields(2).Text = "" Then
            MsgBox "Debe completar el nombre del alimento"
            txtFields(2).SetFocus
        Else
            strQuery = "select * from alimentos where idcategoria = " & DataCombo1.BoundText & " and descripalimento like '" & txtFields(0).Text & "'"
            Set tb = dbdiet.OpenRecordset(strQuery)
            'verifico que la descripcion no exista
            If tb.RecordCount = 0 Then
                
                MDIForm1.ActiveForm.Data1.UpdateRecord
                MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
        
                tb.Close
                
                
''                'actualiza la tabla AlimenxPaciente según corresponda con el control CheckBox
''                If Check1.Value = 1 Then
''                    dbdiet.Execute " insert into alimenxpaciente (legajo, codalimento) select legajo, " & nCodAlimento & " from pacientes "
''                    'dbdiet.Execute " update alimentos set estado = true where codalimento = " & nCodAlimento
''                Else
''                    dbdiet.Execute " delete from alimenxpaciente where codalimento = " & nCodAlimento
''                End If
                            
                'aux = 1
            Else
                If estadoAbm = 3 Then
                
                    Data1.UpdateRecord
                    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
                    
''                    If Check1.Value = 1 Then
''                        dbdiet.Execute " insert into alimenxpaciente (legajo, codalimento) select legajo, " & nCodAlimento & " from pacientes "
''                        'dbdiet.Execute " update alimentos set estado = true where codalimento = " & nCodAlimento
''                    Else
''                        dbdiet.Execute " delete from alimenxpaciente where codalimento = " & nCodAlimento
''                    End If
                Else
                    MsgBox "El alimento ingresado ya fue incluído dentro de la categoría " & DataCombo1.Text
                    
                    txtFields(0).SetFocus
                    
                    Exit Sub
                    
                End If
            
            End If
        
            'condiciones extras
                'If estadoAbm = 2 Then
                '    dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(MDIForm1.ActiveForm.lbl_CodAlimento.Caption) & ", codalimento from alimentos where estado = true"
                'End If
                
            cmdBuscar.Enabled = True
            cmdAgregar.Enabled = True
            'cmdBorrar.Enabled = True
            'cmdClose.Enabled = True
            'cmdModificar.Enabled = True
            
            cmdAgregar.SetFocus
            cmdAgregar.Default = True
            cmdCancelar.Cancel = True
            
            'cmdPrimero.Enabled = True
            'cmdAnterior.Enabled = True
            'cmdSiguiente.Enabled = True
            'cmdUltimo.Enabled = True
            
            Call enabledDesplaz 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados
            Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)
        
            estadoAbm = 1 ' el estado del form es "sin cambios"
            
            Call cmd_Actualizar_Click
            
            Data1.Recordset.FindFirst " codalimento = " & nCodAlimento_add
            
            Call enabledDesplaz
            
            Call f_Boton_Zorder
            
        End If
        
    End If

Else
    
    If Not MDIForm1.ActiveForm Is Nothing Then
    
        Unload Me
            
    End If
    
End If

End Sub

Private Sub cmdAgregar_Click()
Dim c

c = DataCombo1.BoundText

Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

estadoAbm = 2 ' el estado es agregar

MDIForm1.ActiveForm.Data1.Recordset.AddNew

DataCombo1.BoundText = c

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

DataCombo1.SetFocus

'Unload frm_formulaDesarrollada

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

If DataCombo1.Text <> "" And txtFields(2).Text <> "" Then

    If MDIForm1.ActiveForm.Data1.Recordset.RecordCount > 0 And MDIForm1.ActiveForm.Data1.Recordset.EOF = False And MDIForm1.ActiveForm.Data1.Recordset.BOF = False Then
        msg = MsgBox("¿Desea Eliminar el registro actual?", vbYesNo, "Eliminar")
        
        If msg = vbYes Then
            'verifica que se pueda eliminar sin problemas y no perder integridad
            
            strQuery = "select * from alimenxpaciente where codalimento = " & nCodAlimento '& " and cantidad <> 0"
            Set tb = dbdiet.OpenRecordset(strQuery)
            strQuery = "select * from ingredientesplatos where codalimento = " & nCodAlimento
            Set tb1 = dbdiet.OpenRecordset(strQuery)
            If tb.RecordCount = 0 And tb1.RecordCount = 0 Then
                       
''                dbdiet.Execute "delete from alimenxpaciente where codalimento = " & nCodAlimento
''                dbdiet.Execute "delete from ingredientesplatos where codalimento = " & nCodAlimento
''                dbdiet.Execute "delete from menu where idalimento = " & nCodAlimento
                
                Data1.Recordset.Delete
                Data1.Recordset.MovePrevious
            Else
                
                MsgBox "No se puede eliminar el registro actual porque puede afectar la integridad del Sistema", vbInformation, "Información"
                
                Dim sMsg As String
                '=======================================
                sMsg = ""
                
                sMsg = sMsg & f_sNoDelete(tb, "Legajo")
                
                MsgBox "El alimento que intenta eliminar esta asignado en la Formula Desarrolla en los siguientes legajos de pacientes: " & vbCrLf & vbTab & sMsg, vbInformation
                '=======================================
                sMsg = ""
                
                sMsg = sMsg & f_sNoDelete(tb1, "idPlato")
                
                MsgBox "El alimento que intenta eliminar esta asignado como ingrediente en los siguientes codigos de platos: " & vbCrLf & vbTab & sMsg, vbInformation
                '=======================================
                                
            End If
            tb.Close
            tb1.Close
                        
            Call f_Boton_Zorder
            
        Else
            cmdAgregar.SetFocus
        End If
    End If
End If

End Sub

Private Sub cmdBuscar_Click()
Dim strQuery As String

'strQuery = " select * from alimentos order by idcategoria, descripalimento"
strQuery = "SELECT * FROM alimentos, Categoria WHERE alimentos.idCategoria = Categoria.idCategoria ORDER BY decripcion, descripalimento"

With Data1
    .RecordSource = strQuery
    .Refresh
End With

'aclare campo por el cual buscar
msg = InputBox("Ingrese descripción del alimento:", "Buscar por Descripción")

If msg <> "" Then

    'strQuery = " select * from alimentos where descripalimento like '" & msg & "*' order by idcategoria, descripalimento"
    strQuery = "SELECT * FROM alimentos, Categoria WHERE alimentos.idCategoria = Categoria.idCategoria AND descripalimento LIKE '*" & msg & "*' ORDER BY decripcion, descripalimento"
    
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
    
    Call enabledDesplaz
        
    Call f_Boton_Zorder
    
Else

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        Unload Me
    
    End If

End If
End Sub



Private Sub cmdImprimir_Click()
Dim strQuery As String
'aclare el filtro para imprimir
msg = MsgBox("¿Desea imprimir todos los registros?", vbYesNo, "Imprimir")

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_alimentos_one.rpt"
    
If msg = vbYes Then
    
    strQuery = ""
    'CrystalReport1.SelectionFormula = "" ' {alimentos.codalimento} " '& nCodAlimento '& " and {platosmenu.fechaMenu} in Date(" & Year(DTdesde.Value) & ", " & Month(DTdesde.Value) & ", " & Day(DTdesde.Value) & ") to Date(" & Year(DThasta.Value) & ", " & Month(DThasta.Value) & ", " & Day(DThasta.Value) & ") "
    'CrystalReport1.Destination = crptToWindow
    'CrystalReport1.PrintReport
Else
    
    strQuery = " {alimentos.codalimento} = " & nCodAlimento
    'CrystalReport1.SelectionFormula = " {alimentos.codalimento} = " & nCodAlimento '& " and {platosmenu.fechaMenu} in Date(" & Year(DTdesde.Value) & ", " & Month(DTdesde.Value) & ", " & Day(DTdesde.Value) & ") to Date(" & Year(DThasta.Value) & ", " & Month(DThasta.Value) & ", " & Day(DThasta.Value) & ") "
    'CrystalReport1.Destination = crptToWindow
    'CrystalReport1.PrintReport
End If

Call f_print(CrystalReport1, strQuery, crptToWindow)

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
DataCombo1.SetFocus

cmdAceptar.Default = True
cmdCancelar.Cancel = True

estadoAbm = 3 ' el estado es modificar

Call f_Boton_Zorder

End Sub

Private Sub cmdPrimero_Click()

MDIForm1.ActiveForm.Data1.Recordset.MoveFirst

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

Call enabledDesplaz

End Sub



Private Sub cmd_Actualizar_Click()
Dim strQuery As String
'strQuery = " select * from alimentos order by idcategoria, descripalimento"
strQuery = "SELECT * FROM alimentos, Categoria WHERE alimentos.idCategoria = Categoria.idCategoria ORDER BY decripcion, descripalimento"

With Data1
    .RecordSource = strQuery
    .Refresh
End With

Call enabledDesplaz

End Sub

Private Sub Command2_Click()

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
    MsgBox "Debe Completar la Categoría", vbInformation, "Información"
End If

End Sub

Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 5805
Me.Width = 6615
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Call enabledDesplaz

Call f_Boton_Zorder

Call TabStrip1_Click

End Sub

Private Sub Form_Load()

'Data1.DatabaseName = Lugar

Call f_CargarOrigenDatos

For i = 0 To 19
    txtFields(i).Enabled = False
Next

aux = 1

'Titulo = Me.Caption

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdCancelar_Click

End Sub

Private Sub lbl_CodAlimento_Change()

nCodAlimento = Val(lbl_CodAlimento.Caption)

Me.Caption = "Alimentos - Nro. " & Val(lbl_CodAlimento.Caption)

End Sub

Private Sub Pic_Imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Imprimir

End Sub

Private Sub TabStrip1_Click()
'define los valores de los frame correspondientes para que funcionen con el tabstrip
Dim a As String
a = TabStrip1.SelectedItem.Index

Select Case a
    Case Is = 1
    
        fme_Macronutrientes.ZOrder 0
        fme_Minerales.ZOrder 1
        fme_Vitaminas.ZOrder 1
        fme_Acidos_Grasos.ZOrder 1
        fme_Otros.ZOrder 1
        
    Case Is = 2
    
        fme_Macronutrientes.ZOrder 1
        fme_Minerales.ZOrder 0
        fme_Vitaminas.ZOrder 1
        fme_Acidos_Grasos.ZOrder 1
        fme_Otros.ZOrder 1
        
    Case Is = 3
    
        fme_Macronutrientes.ZOrder 1
        fme_Minerales.ZOrder 1
        fme_Vitaminas.ZOrder 0
        fme_Acidos_Grasos.ZOrder 1
        fme_Otros.ZOrder 1
    
    Case Is = 4
    
        fme_Macronutrientes.ZOrder 1
        fme_Minerales.ZOrder 1
        fme_Vitaminas.ZOrder 1
        fme_Acidos_Grasos.ZOrder 0
        fme_Otros.ZOrder 1
        
    Case Is = 5
    
        fme_Macronutrientes.ZOrder 1
        fme_Minerales.ZOrder 1
        fme_Vitaminas.ZOrder 1
        fme_Acidos_Grasos.ZOrder 1
        fme_Otros.ZOrder 0
        
End Select


End Sub

Private Sub txtFields_GotFocus(Index As Integer)
For i = 0 To 19
    txtFields(i).SelStart = 0
    txtFields(i).SelLength = 50
Next

Select Case Index

    Case Is = 1
        Set Me.TabStrip1.SelectedItem = Me.TabStrip1.Tabs(1)
    Case Is = 9
        Set Me.TabStrip1.SelectedItem = Me.TabStrip1.Tabs(2)
    Case Is = 10
        Set Me.TabStrip1.SelectedItem = Me.TabStrip1.Tabs(3)
    Case Is = 16
        Set Me.TabStrip1.SelectedItem = Me.TabStrip1.Tabs(4)
    Case Is = 19
        Set Me.TabStrip1.SelectedItem = Me.TabStrip1.Tabs(5)
        
End Select

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Data1.Recordset = Nothing
Set Me.Adodc1.Recordset = Nothing
Set Me.Adodc2.Recordset = Nothing

strQuery = "SELECT * FROM alimentos, Categoria WHERE alimentos.idCategoria = Categoria.idCategoria ORDER BY decripcion, descripalimento"
Call f_Data_DatabaseName(Data1, strQuery)

strQuery = "select * from Categoria"
Call f_Adodc_ConnectionString(Adodc1, strQuery)

strQuery = "select * from unidades"
Call f_Adodc_ConnectionString(Adodc2, strQuery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo1, Data1, Adodc1, "alimentos.idCategoria", "idCategoria", "Decripcion")

Call f_Enlaza_ControlData(DataCombo2, Data1, Adodc2, "idUnidad", "idUnidad", "DescripUnidad")
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


Function f_sNoDelete(tb As Recordset, sNameCodigo As String) As String      '(param_Array_RecordSet() As Recordset)

Dim nCount, i As Integer
Dim lFirst As Boolean

lFirst = True

f_sNoDelete = ""

nCount = f_Cant_Registros(tb)

tb.MoveFirst
For i = 1 To nCount
    
    If lFirst Then
    
        f_sNoDelete = f_sNoDelete & tb.Fields(sNameCodigo).Value
        lFirst = False
        
    Else
    
        f_sNoDelete = f_sNoDelete & ", " & tb.Fields(sNameCodigo).Value
        
    End If
    
    tb.MoveNext
Next

End Function
