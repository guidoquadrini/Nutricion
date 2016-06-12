VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_formulaDesarrollada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fórmula Desarrollada"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frm_formulaDesarrollada.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10965
   Begin MSAdodcLib.Adodc adodc_Totales_Nutrientes 
      Height          =   375
      Left            =   7320
      Top             =   5400
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "adodc_Totales_Nutrientes"
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   1200
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
      Caption         =   "Adodc6"
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
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   315
      Left            =   9840
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame 
      Caption         =   "Tabla de Alimentos:"
      Height          =   5775
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.Frame fme_tabla 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   4095
         Left            =   2640
         TabIndex        =   32
         Top             =   1440
         Width           =   8055
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frm_formulaDesarrollada.frx":0ECA
            Height          =   3375
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   5953
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   -2147483633
            BorderStyle     =   0
            HeadLines       =   3
            RowHeight       =   16
            RowDividerStyle =   0
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Detalle"
            ColumnCount     =   12
            BeginProperty Column00 
               DataField       =   "tmp_Legajo"
               Caption         =   "Legajo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "tmp_CodAlimento"
               Caption         =   "CodAlimento"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "idcategoria"
               Caption         =   "idcategoria"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "DescripAlimento"
               Caption         =   "Descripcion Alimento"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "tmp_Cantidad"
               Caption         =   "Cantidad"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "Hidratos de Carbono"
               Caption         =   "Hidratos de Carbono (gr.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "Proteínas"
               Caption         =   "Proteinas (gr.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "Lípidos"
               Caption         =   "Lipidos (gr.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "Kcal"
               Caption         =   "Kcal"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "HC1"
               Caption         =   "HC1"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "Prot1"
               Caption         =   "Prot1"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "Lip1"
               Caption         =   "Lip1"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column03 
                  Alignment       =   3
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column04 
                  ColumnAllowSizing=   0   'False
               EndProperty
               BeginProperty Column05 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column06 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column07 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column08 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column09 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column10 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            DataField       =   "sumhc"
            DataSource      =   "Adodc4"
            Height          =   255
            Left            =   3000
            TabIndex        =   43
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Label1"
            DataField       =   "sumprot"
            DataSource      =   "Adodc4"
            Height          =   255
            Left            =   4440
            TabIndex        =   42
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Label1"
            DataField       =   "sumlip"
            DataSource      =   "Adodc4"
            Height          =   255
            Left            =   5520
            TabIndex        =   41
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Label1"
            DataField       =   "sumkcal"
            DataSource      =   "Adodc4"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6600
            TabIndex        =   40
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Label lbl 
            BackColor       =   &H8000000A&
            Caption         =   "Label1"
            DataField       =   "rctideal"
            DataSource      =   "Adodc5"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   3
            Left            =   6600
            TabIndex        =   39
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label lbl 
            BackColor       =   &H8000000A&
            Caption         =   "Label1"
            DataField       =   "lipg"
            DataSource      =   "Adodc5"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   5520
            TabIndex        =   38
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label lbl 
            BackColor       =   &H8000000A&
            Caption         =   "Label1"
            DataField       =   "protg"
            DataSource      =   "Adodc5"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   37
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label lbl 
            BackColor       =   &H8000000A&
            Caption         =   "Label1"
            DataField       =   "hcg"
            DataSource      =   "Adodc5"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   36
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Totales ........................................"
            Height          =   255
            Left            =   360
            TabIndex        =   35
            Top             =   3480
            Width           =   2415
         End
         Begin VB.Label Label6 
            Caption         =   "Totales Ideales ............................"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   3840
            Width           =   2415
         End
      End
      Begin VB.Frame fme_Valores_Nutrientes 
         BorderStyle     =   0  'None
         Caption         =   "Valores Nutrientes"
         Height          =   4095
         Left            =   2640
         TabIndex        =   44
         Top             =   1440
         Width           =   8055
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "frm_formulaDesarrollada.frx":0EDF
            Height          =   3615
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   6376
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   -2147483633
            BorderStyle     =   0
            HeadLines       =   3
            RowHeight       =   16
            RowDividerStyle =   0
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Detalle"
            ColumnCount     =   17
            BeginProperty Column00 
               DataField       =   "DescripAlimento"
               Caption         =   "Descripcion Alimento"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "tmp_Fibra"
               Caption         =   "Fibra (gr.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "tmp_Sodio"
               Caption         =   "Sodio (mg.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "tmp_Calcio"
               Caption         =   "Calcio (mg.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "tmp_Hierro"
               Caption         =   "Hierro (mg.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "tmp_Fosforo"
               Caption         =   "Fosforo (mg.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "tmp_Potasio"
               Caption         =   "Potasio (mg.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "tmp_Glucosa"
               Caption         =   "Glucosa (gr.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "tmp_VitA"
               Caption         =   "VitA (U.I.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "tmp_VitB1"
               Caption         =   "VitB1 (mg.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "tmp_VitB2"
               Caption         =   "VitB2 (mg.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "tmp_VitNiacina"
               Caption         =   "VitNiacina (mg.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column12 
               DataField       =   "tmp_VitC"
               Caption         =   "VitC (mg.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column13 
               DataField       =   "tmp_VitE"
               Caption         =   "VitE (U.I.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column14 
               DataField       =   "tmp_AGS"
               Caption         =   "AGS (gr.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column15 
               DataField       =   "tmp_AGMI"
               Caption         =   "AGMI (gr.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column16 
               DataField       =   "tmp_AGPI"
               Caption         =   "AGPI (gr.)"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column04 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column05 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column06 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column07 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column09 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column10 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column11 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column12 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column13 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column14 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column15 
                  Locked          =   -1  'True
               EndProperty
               BeginProperty Column16 
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Fme_Totales_Nutrientes 
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   2640
         TabIndex        =   46
         Top             =   1440
         Width           =   8055
         Begin VB.Frame fme_Minerales 
            Caption         =   "Minerales"
            Height          =   1335
            Left            =   75
            TabIndex        =   83
            Top             =   0
            Width           =   7935
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "mg."
               Height          =   195
               Left            =   7080
               TabIndex        =   98
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "mg."
               Height          =   195
               Left            =   7080
               TabIndex        =   97
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "mg."
               Height          =   195
               Left            =   3000
               TabIndex        =   96
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "mg."
               Height          =   195
               Left            =   3000
               TabIndex        =   95
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "mg."
               Height          =   195
               Left            =   3000
               TabIndex        =   94
               Top             =   960
               Width           =   255
            End
            Begin VB.Label lbl_Fosforo 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_Fosforo"
               DataField       =   "sum_Fosforo"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   5760
               TabIndex        =   93
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl_Hierro 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_Hierro"
               DataField       =   "sum_Hierro"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   5760
               TabIndex        =   92
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl_Calcio 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_Calcio"
               DataField       =   "sum_Calcio"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   91
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl_Sodio 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_Sodio"
               DataField       =   "sum_Sodio"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   90
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl_Potasio 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_Potasio"
               DataField       =   "sum_Potasio"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   89
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Sodio:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   88
               Top             =   600
               Width           =   555
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Calcio:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   87
               Top             =   960
               Width           =   600
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Hierro:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   4200
               TabIndex        =   86
               Top             =   240
               Width           =   585
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Fósforo:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   4200
               TabIndex        =   85
               Top             =   600
               Width           =   705
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Potasio:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   705
            End
         End
         Begin VB.Frame fme_Vitaminas 
            Caption         =   "Vitaminas"
            Height          =   1335
            Left            =   75
            TabIndex        =   50
            Top             =   1320
            Width           =   7935
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "U.I."
               Height          =   195
               Left            =   7080
               TabIndex        =   77
               Top             =   960
               Width           =   255
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "mg."
               Height          =   195
               Left            =   7080
               TabIndex        =   76
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "mg."
               Height          =   195
               Left            =   7080
               TabIndex        =   75
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "mg."
               Height          =   195
               Left            =   3000
               TabIndex        =   74
               Top             =   960
               Width           =   255
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "mg."
               Height          =   195
               Left            =   3000
               TabIndex        =   73
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "U.I."
               Height          =   195
               Left            =   3000
               TabIndex        =   72
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_VitE"
               DataField       =   "sum_VitE"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   5760
               TabIndex        =   68
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl_VitC 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_VitC"
               DataField       =   "sum_VitC"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   5760
               TabIndex        =   67
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl_VitNiacina 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_VitNiacina"
               DataField       =   "sum_VitNiacina"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   5760
               TabIndex        =   66
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl_VitB2 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_VitB2"
               DataField       =   "sum_VitB2"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   65
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl_VitB1 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_VitB1"
               DataField       =   "sum_VitB1"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   64
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl_VitA 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_VitA"
               DataField       =   "sum_VitA"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   63
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "E:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   4200
               TabIndex        =   56
               Top             =   960
               Width           =   195
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "C:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   4200
               TabIndex        =   55
               Top             =   600
               Width           =   195
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Niacina:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   4200
               TabIndex        =   54
               Top             =   240
               Width           =   720
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "B2:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   53
               Top             =   960
               Width           =   300
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "B1:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   120
               TabIndex        =   52
               Top             =   600
               Width           =   300
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "A:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   195
            End
         End
         Begin VB.Frame fme_Acidos_Grasos 
            Caption         =   "Acidos Grasos"
            Height          =   1335
            Left            =   75
            TabIndex        =   57
            Top             =   2640
            Width           =   3855
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "gr."
               Height          =   195
               Left            =   3000
               TabIndex        =   80
               Top             =   960
               Width           =   180
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "gr."
               Height          =   195
               Left            =   3000
               TabIndex        =   79
               Top             =   600
               Width           =   180
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "gr."
               Height          =   195
               Left            =   3000
               TabIndex        =   78
               Top             =   240
               Width           =   180
            End
            Begin VB.Label lbl_AGPI 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_AGPI"
               DataField       =   "sum_AGPI"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   71
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl_AGMI 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_AGMI"
               DataField       =   "sum_AGMI"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   70
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl_AGS 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_AGS"
               DataField       =   "sum_AGS"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   69
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Polinsaturados:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   20
               Left            =   120
               TabIndex        =   60
               Top             =   960
               Width           =   1320
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Monoinsaturados:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   120
               TabIndex        =   59
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Saturados:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Width           =   930
            End
         End
         Begin VB.Frame fme_Otros 
            Caption         =   "Otros..."
            Height          =   1335
            Left            =   4155
            TabIndex        =   47
            Top             =   2640
            Width           =   3855
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "gr."
               Height          =   195
               Left            =   3000
               TabIndex        =   82
               Top             =   600
               Width           =   180
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "gr."
               Height          =   195
               Left            =   3000
               TabIndex        =   81
               Top             =   240
               Width           =   180
            End
            Begin VB.Label lbl_Fibra 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_Fibra"
               DataField       =   "sum_Fibra"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   62
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl_Glucosa 
               Alignment       =   1  'Right Justify
               Caption         =   "lbl_Glucosa"
               DataField       =   "sum_Glucosa"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               DataSource      =   "adodc_Totales_Nutrientes"
               Height          =   255
               Left            =   1680
               TabIndex        =   61
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Glucosa:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Width           =   765
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               Caption         =   "Fibra:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   48
               Top             =   600
               Width           =   495
            End
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   4575
         Left            =   2520
         TabIndex        =   31
         Top             =   1080
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8070
         ShowTips        =   0   'False
         HotTracking     =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fórmula Desarrollada"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Tabla de nutrientes"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Totales de nutrientes"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame 
         Caption         =   "Paciente:"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10695
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frm_formulaDesarrollada.frx":0EF4
            DataField       =   "Legajo"
            DataSource      =   "Adodc6"
            Height          =   315
            Left            =   2640
            TabIndex        =   6
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "nom"
            BoundColumn     =   "Legajo"
            Text            =   ""
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   495
            Left            =   5520
            TabIndex        =   8
            Top             =   120
            Width           =   2295
            Begin VB.CommandButton cmd_cerrar 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_formulaDesarrollada.frx":0F09
               Height          =   375
               Left            =   1800
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frm_formulaDesarrollada.frx":109D
               Picture         =   "frm_formulaDesarrollada.frx":11EF
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Cancelar"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.CommandButton cmd_aceptar 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_formulaDesarrollada.frx":16A2
               Height          =   375
               Left            =   1320
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frm_formulaDesarrollada.frx":17FB
               Picture         =   "frm_formulaDesarrollada.frx":194D
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Aceptar"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.PictureBox Pic_Aceptar 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1320
               MouseIcon       =   "frm_formulaDesarrollada.frx":1C09
               Picture         =   "frm_formulaDesarrollada.frx":1D5B
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   10
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Cerrar 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1800
               MouseIcon       =   "frm_formulaDesarrollada.frx":2017
               Picture         =   "frm_formulaDesarrollada.frx":2169
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   9
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Aceptar_Gris 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1320
               MouseIcon       =   "frm_formulaDesarrollada.frx":246A
               Picture         =   "frm_formulaDesarrollada.frx":25BC
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   13
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Cerrar_Gris 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1800
               MouseIcon       =   "frm_formulaDesarrollada.frx":2715
               Picture         =   "frm_formulaDesarrollada.frx":2867
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   14
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton cmd_Tipito 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_formulaDesarrollada.frx":29FB
               Height          =   315
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frm_formulaDesarrollada.frx":310B
               Style           =   1  'Graphical
               TabIndex        =   28
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
               MouseIcon       =   "frm_formulaDesarrollada.frx":339B
               Picture         =   "frm_formulaDesarrollada.frx":34ED
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   30
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
               MouseIcon       =   "frm_formulaDesarrollada.frx":361D
               Picture         =   "frm_formulaDesarrollada.frx":376F
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   29
               Top             =   120
               Width           =   315
            End
         End
      End
      Begin VB.PictureBox CrystalReport1 
         Height          =   480
         Left            =   120
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   99
         Top             =   4800
         Width           =   1200
      End
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "frm_formulaDesarrollada.frx":39FF
         DataField       =   "idCategoria"
         DataSource      =   "Adodc3"
         Height          =   4155
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   7329
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ListField       =   "Decripcion"
         BoundColumn     =   "idCategoria"
      End
      Begin VB.Label Label7 
         Caption         =   "Grupos de alimentos:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   9960
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   9960
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3120
      Top             =   2040
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9960
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3240
      Top             =   1080
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
   Begin VB.Frame Frame 
      Caption         =   "Grupos de alimentos"
      Height          =   855
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   -120
      Visible         =   0   'False
      Width           =   2415
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frm_formulaDesarrollada.frx":3A14
         DataField       =   "idCategoria"
         DataSource      =   "Adodc3"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "Decripcion"
         BoundColumn     =   "idCategoria"
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4335
      TabIndex        =   15
      Top             =   5760
      Width           =   2295
      Begin VB.CommandButton cmd_Cancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_formulaDesarrollada.frx":3A29
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_formulaDesarrollada.frx":3BA1
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Deshacer cambios"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmd_guardar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_formulaDesarrollada.frx":4021
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_formulaDesarrollada.frx":417A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Guardar cambios"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox Pic_Guardar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         MouseIcon       =   "frm_formulaDesarrollada.frx":4436
         Picture         =   "frm_formulaDesarrollada.frx":4588
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   16
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Cancelar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1200
         MouseIcon       =   "frm_formulaDesarrollada.frx":4844
         Picture         =   "frm_formulaDesarrollada.frx":4996
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   17
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Guardar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         MouseIcon       =   "frm_formulaDesarrollada.frx":4E16
         Picture         =   "frm_formulaDesarrollada.frx":4F68
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   24
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Cancelar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1200
         MouseIcon       =   "frm_formulaDesarrollada.frx":50C1
         Picture         =   "frm_formulaDesarrollada.frx":5213
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   22
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_salir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_formulaDesarrollada.frx":538B
         Height          =   375
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_formulaDesarrollada.frx":551F
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Salir"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox Pic_Salir 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         MouseIcon       =   "frm_formulaDesarrollada.frx":5820
         Picture         =   "frm_formulaDesarrollada.frx":5972
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   18
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Salir_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         MouseIcon       =   "frm_formulaDesarrollada.frx":5C73
         Picture         =   "frm_formulaDesarrollada.frx":5DC5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   23
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_formulaDesarrollada.frx":5F59
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_formulaDesarrollada.frx":60B1
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   120
         MouseIcon       =   "frm_formulaDesarrollada.frx":6531
         Picture         =   "frm_formulaDesarrollada.frx":6683
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   25
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Imprimir_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         MouseIcon       =   "frm_formulaDesarrollada.frx":6B03
         Picture         =   "frm_formulaDesarrollada.frx":6C55
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   27
         Top             =   120
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_formulaDesarrollada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
Dim bandera As Integer
Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
                            
Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim aa As Integer
aa = 0
If bandera = 1 Then
    
    If Not Adodc1.Recordset.BOF = True And Not Adodc1.Recordset.EOF = True Then
    
        aa = Adodc1.Recordset.Fields("idcategoria").Value
        
        DataCombo2.BoundText = aa
        DataList1.BoundText = aa
    
    End If
    
End If
End Sub

Private Sub cmd_Cancelar_Click()

MousePointer = vbHourglass

Call f_DeshacerCambios

Call f_Boton_Zorder

MousePointer = vbDefault

End Sub

Private Sub cmd_cerrar_Click()

'Unload Me
frm_formulaDesarrollada.Hide

End Sub

Private Sub cmdAgregar_Click()

End Sub

Private Sub cmdBorrar_Click()

End Sub

Private Sub cmdImprimir_Click()

frm_formulaDesarrollada_Print.Show vbModal

''Dim strquery, sMsg As String
''
'''Resets the value of all properties (except DataSource Property) to their default values.
''CrystalReport1.Reset
''
''sMsg = MsgBox("¿Desea imprimir informe detallado?", vbYesNoCancel, "Imprimir")
''
''If sMsg = vbYes Then
''
''    CrystalReport1.ReportFileName = App_Path & "\rpts\rep_formdesarrollada_one.rpt"
''
''    'CrystalReport1.ParameterFields(4) = "SortField;Legajo;True"
''
''    'CrystalReport1.ParameterFields(4) = "SortField;Obra Social;True"
''
''    'CrystalReport1.ParameterFields(4) = "SortField;ApellyNom;True"
''
''    strquery = " {pacientes.legajo} = " & DataCombo1.BoundText
''
''Else
''
''    If sMsg = vbNo Then
''
''        CrystalReport1.ReportFileName = App_Path & "\rpts\rep_formdesarrolladaSinCant_one.rpt"
''
''        strquery = " {pacientes.legajo} = " & DataCombo1.BoundText
''
''    End If
''
''End If
''
''If Not sMsg = vbCancel Then
''    Call f_print(CrystalReport1, strquery, crptToWindow)
''End If

End Sub

Private Sub cmdSiguiente_Click()

End Sub

Private Sub cmd_tipito_Click()
Unload frmPacientes
frmPacientes.Show
frmPacientes.Data1.Recordset.FindFirst " legajo = " & DataCombo1.BoundText

End Sub

Private Sub cmd_aceptar_Click()
Dim strQuery As String

MousePointer = vbHourglass

If DataCombo1.Text <> "" Then
    
    'inserta los datos que estan cargados en la tabla alimenxpaciente
    dbdiet.Execute "insert into alimenxpaciente_tmp select legajo as tmp_legajo, codalimento as tmp_codalimento, cantidad as tmp_cantidad, hc as tmp_hc, prot as tmp_prot, lip as tmp_lip, kcal as tmp_kcal, Fibra AS tmp_Fibra, Sodio AS tmp_Sodio, Calcio AS tmp_Calcio, Hierro AS tmp_Hierro, Fosforo AS tmp_Fosforo, Potasio AS tmp_Potasio, Glucosa AS tmp_Glucosa, VitA AS tmp_VitA, VitB1 AS tmp_VitB1, VitB2 AS tmp_VitB2, VitNiacina AS tmp_VitNiacina, VitC AS tmp_VitC, VitE AS tmp_VitE, AGS AS tmp_AGS, AGMI AS tmp_AGMI, AGPI AS tmp_AGPI from alimenxpaciente where alimenxpaciente.legajo = " & DataCombo1.BoundText
    
    'inserta los datos restantes para completar la tabla
    dbdiet.Execute "insert into alimenxpaciente_tmp (tmp_legajo, tmp_codalimento) select " & DataCombo1.BoundText & ", codalimento from alimentos where estado = true"
    
    'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
    dbdiet.Execute "insert into alimenxpaciente_tmp (tmp_legajo, tmp_codalimento) select " & DataCombo1.BoundText & ", codalimento from alimentos where estado = true"
          
    'strquery = "select * from consultaprueba3 where alimenxpaciente.legajo = " & DataCombo1.BoundText
    strQuery = "select * from csl_alimentosxpacientes where alimenxpaciente_tmp.tmp_legajo = " & DataCombo1.BoundText '& " and alimentos.estado = 1"
    
    With Adodc1
        .RecordSource = strQuery
        .Refresh
    End With
    
    With DataGrid1
        .ReBind
        .Refresh
    End With
               
    Call DatagridRefresh(Adodc1, DataGrid1)
    
    strQuery = "select rctideal, hcg, protg, lipg from pacientes where legajo = " & DataCombo1.BoundText
        
    With Adodc5
        .RecordSource = strQuery
        .Refresh
    End With

    adodc_Totales_Nutrientes.Refresh
        
    For i = 0 To 3
        lbl(i).Caption = Format(Val(lbl(i).Caption), "standard")
    Next

    Frame(0).Enabled = False
    DataList1.Enabled = True
    DataGrid1.Enabled = True
        
    DataList1.BackColor = &H80000005
    DataGrid1.BackColor = &H80000005
    DataGrid2.BackColor = &H80000005
    
    cmd_salir.Enabled = True
    cmdImprimir.Enabled = True
    'cmdImprimir.Enabled = False
    cmd_Aceptar.Enabled = False
    cmd_Cerrar.Enabled = False
    cmd_Tipito.Enabled = False
    'Frame1.Visible = True
    Frame1.Enabled = True
    
    Call calculaTotales
               
    estadoAbm = 1
        
    Call f_Boton_Zorder
    
End If

MousePointer = vbDefault

End Sub





Private Sub DataCombo1_Change()

If DataCombo1.Text <> "" Then
    Me.Caption = " Fórmula Desarrollada " & " - Nro. " & DataCombo1.BoundText & " - " & DataCombo1.Text
End If

End Sub

Private Sub DataCombo1_LostFocus()
If DataCombo1.Text = "" Then
    DataCombo1.SetFocus
    MsgBox "Debe Completar el Nombre del Paciente", vbInformation, "Información"
End If
End Sub

Private Sub DataCombo2_Click(Area As Integer)
Dim aa As Integer
aa = 0

strQuery = " select * from consultaprueba3 where idcategoria = " & DataCombo2.BoundText '& " and legajo = " & DataCombo1.BoundText
Set tb = dbdiet.OpenRecordset(strQuery)
If tb.RecordCount <> 0 Then
    bandera = 0
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find "idcategoria = " & DataCombo2.BoundText
    'MsgBox Adodc1.Recordset.AbsolutePosition
'    DataGrid1.SetFocus
    
    aa = Adodc1.Recordset.Fields("idcategoria").Value
    DataCombo2.BoundText = aa
    bandera = 1
Else
    MsgBox "No existe alimento en la categoría seleccionada"
    
End If
tb.Close
'DataCombo2.SetFocus
End Sub

Private Sub DataCombo2_LostFocus()

If DataCombo2.Text = "" Then
    DataCombo2.SetFocus
    MsgBox "Debe Completar la Categoría ", vbInformation, "Información"
End If

End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
Dim hc, prot, lip, Kcal
Dim strQuery As String

MousePointer = vbHourglass

'"hidratos de carbono" --> columna 5
'"proteínas" ------------> columna 6
'"lípidos" --------------> columna 7
'"kcal" -----------------> columna 8
    
If Adodc1.Recordset.RecordCount > 0 Then

    If DataGrid1.Columns(4) <> "" Then
        
        '=====================Calculo de macronutrientes=======================================
        hc = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("hc1").Value / 100
        prot = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("prot1").Value / 100
        lip = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("lip1").Value / 100
        Kcal = hc * 4 + prot * 4 + lip * 9
       
        DataGrid1.Columns(5).Text = hc
        DataGrid1.Columns(6).Text = prot
        DataGrid1.Columns(7).Text = lip
        DataGrid1.Columns(8).Text = Kcal
        
        Adodc1.Recordset.Update "hidratos de carbono", hc
        Adodc1.Recordset.Update "proteínas", prot
        Adodc1.Recordset.Update "lípidos", lip
        Adodc1.Recordset.Update "kcal", Kcal
        '======================================================================================
        
        '=====================Calculo del resto de nutrientes=======================================
        Fibra = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("Fibra").Value / 100
        Sodio = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("Sodio").Value / 100
        Calcio = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("Calcio").Value / 100
        Hierro = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("Hierro").Value / 100
        Fosforo = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("Fosforo").Value / 100
        Potasio = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("Potasio").Value / 100
        Glucosa = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("Glucosa").Value / 100
        VitA = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("VitA").Value / 100
        VitB1 = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("VitB1").Value / 100
        VitB2 = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("VitB2").Value / 100
        VitNiacina = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("VitNiacina").Value / 100
        VitC = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("VitC").Value / 100
        VitE = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("VitE").Value / 100
        AGS = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("AGS").Value / 100
        AGMI = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("AGMI").Value / 100
        AGPI = Adodc1.Recordset.Fields("tmp_cantidad").Value * Adodc1.Recordset.Fields("AGPI").Value / 100
              
        DataGrid2.Columns(1).Text = Fibra
        DataGrid2.Columns(2).Text = Sodio
        DataGrid2.Columns(3).Text = Calcio
        DataGrid2.Columns(4).Text = Hierro
        DataGrid2.Columns(5).Text = Fosforo
        DataGrid2.Columns(6).Text = Potasio
        DataGrid2.Columns(7).Text = Glucosa
        DataGrid2.Columns(8).Text = VitA
        DataGrid2.Columns(9).Text = VitB1
        DataGrid2.Columns(10).Text = VitB2
        DataGrid2.Columns(11).Text = VitNiacina
        DataGrid2.Columns(12).Text = VitC
        DataGrid2.Columns(13).Text = VitE
        DataGrid2.Columns(14).Text = AGS
        DataGrid2.Columns(15).Text = AGMI
        DataGrid2.Columns(16).Text = AGPI
                        
        Adodc1.Recordset.Update "tmp_Fibra", Fibra
        Adodc1.Recordset.Update "tmp_Sodio", Sodio
        Adodc1.Recordset.Update "tmp_Calcio", Calcio
        Adodc1.Recordset.Update "tmp_Hierro", Hierro
        Adodc1.Recordset.Update "tmp_Fosforo", Fosforo
        Adodc1.Recordset.Update "tmp_Potasio", Potasio
        Adodc1.Recordset.Update "tmp_Glucosa", Glucosa
        Adodc1.Recordset.Update "tmp_VitA", VitA
        Adodc1.Recordset.Update "tmp_VitB1", VitB1
        Adodc1.Recordset.Update "tmp_VitB2", VitB2
        Adodc1.Recordset.Update "tmp_VitNiacina", VitNiacina
        Adodc1.Recordset.Update "tmp_VitC", VitC
        Adodc1.Recordset.Update "tmp_VitE", VitE
        Adodc1.Recordset.Update "tmp_AGS", AGS
        Adodc1.Recordset.Update "tmp_AGMI", AGMI
        Adodc1.Recordset.Update "tmp_AGPI", AGPI
        '======================================================================================
        
               
        Call calculaTotales
                
        adodc_Totales_Nutrientes.Refresh
                
        Adodc1.Recordset.MoveNext
        
        If Adodc1.Recordset.EOF = True Then
            Adodc1.Recordset.MoveLast
        End If
        
    Else
        DataGrid1.Columns(4) = 0
        Adodc1.Recordset.Fields("hidratos de carbono").Value = 0
        Adodc1.Recordset.Fields("proteínas").Value = 0
        Adodc1.Recordset.Fields("lípidos").Value = 0
        Adodc1.Recordset.Fields("kcal").Value = 0
    
    End If

End If

MousePointer = vbDefault

End Sub

Private Sub DataGrid1_Change()
    
estadoAbm = 3

Me.cmd_Guardar.Enabled = True
Me.cmd_Cancelar.Enabled = True

Call f_Boton_Zorder

End Sub

Private Sub DataGrid1_LostFocus()

Call DataGrid1_AfterColUpdate(4)

End Sub

Private Sub DataList1_Click()
Dim aa As Integer
aa = 0

'strquery = "select * from consultaprueba3 where alimenxpaciente.legajo = " & DataCombo1.BoundText '& " and alimentos.estado = 1"
strQuery = "select * from csl_alimentosxpacientes where alimenxpaciente_tmp.tmp_legajo = " & DataCombo1.BoundText '& " and alimentos.estado = 1"

Set tb = dbdiet.OpenRecordset(strQuery)
'If tb.RecordCount <> 0 Then
bandera = 0
    'Adodc1.Recordset.MoveFirst
    'Adodc1.Recordset.Find "idcategoria = " & DataList1.BoundText
tb.MoveFirst
tb.FindFirst "idcategoria = " & DataList1.BoundText
    'MsgBox Adodc1.Recordset.AbsolutePosition
'    DataGrid1.SetFocus
If tb.NoMatch = False Then
    ss = tb.AbsolutePosition
        'tb.Seek
        'Adodc1.Recordset.AbsolutePosition = tb.AbsolutePosition + 1
    DataGrid1.Bookmark = tb.AbsolutePosition + 1
        'aa = Adodc1.Recordset.Fields("idcategoria").Value
        'DataCombo2.BoundText = aa
        'DataList1.BoundText = aa
    bandera = 1
    tb.Close
Else
    MsgBox "No existe alimento en la categoría seleccionada"
    
End If

End Sub

Private Sub DataList1_LostFocus()
Dim aa As Integer
aa = 0
Adodc3.Refresh
If bandera = 1 Then
    aa = Adodc1.Recordset.Fields("idcategoria").Value

        DataCombo2.BoundText = aa
        DataList1.BoundText = aa
End If

End Sub

Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 6750
Me.Width = 11055 ' 11265
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Call TabStrip1_Click

Adodc2.Refresh
Adodc6.Refresh

End Sub

Private Sub Form_Load()
Dim strQuery As String
Dim i

If DataCombo1.Text <> "" Then

    Me.Caption = " Fórmula Desarrollada " & " - " & DataCombo1.Text

End If

Call DatagridWidth

Call f_CargarOrigenDatos

'DataCombo1.BoundText = 1

For i = 0 To 2
        Frame(i).ZOrder 0
Next


If DataCombo1.BoundText <> "" Then

    strQuery = "select * from csl_alimentosxpacientes where alimenxpaciente_tmp.tmp_legajo = " & DataCombo1.BoundText '& " and alimentos.estado = 1"

    With Adodc1
        .RecordSource = strQuery
        .Refresh
    End With

    With DataGrid1
        .ReBind
        .Refresh
    End With
        
    strQuery = "select rctideal, hcg, protg, lipg from pacientes where legajo = " & DataCombo1.BoundText
    
    With Adodc5
        .RecordSource = strQuery
        .Refresh
    End With
    
    For i = 0 To 3
        lbl(i).Caption = Format(Val(lbl(i).Caption), "standard")
    Next
    
    Call calculaTotales
    
    adodc_Totales_Nutrientes.Refresh
           
End If

Frame(0).Enabled = True
DataList1.Enabled = False
DataGrid1.Enabled = False

bandera = 1

estadoAbm = 1

Call f_Boton_Zorder

End Sub
Private Sub calculaTotales()
Dim strQuery As String

'strquery = " select sum(alimenxpaciente.hc) as sumhc, sum(alimenxpaciente.prot) as sumprot, sum(alimenxpaciente.lip) as sumlip, sum(alimenxpaciente.kcal) as sumkcal from alimenxpaciente where legajo = " & DataCombo1.BoundText
strQuery = " select sum(alimenxpaciente_tmp.tmp_hc) as sumhc, sum(alimenxpaciente_tmp.tmp_prot) as sumprot, sum(alimenxpaciente_tmp.tmp_lip) as sumlip, sum(alimenxpaciente_tmp.tmp_kcal) as sumkcal from alimenxpaciente_tmp where tmp_legajo = " & DataCombo1.BoundText

With Adodc4
    .RecordSource = strQuery
    .Refresh
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Frame(0).Enabled = False Then

    Call f_finalizaOperacion

End If

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Label1_Change()
Label1.Caption = Format(Label1.Caption, "standard")
End Sub

Private Sub Label2_Change()
Label2.Caption = Format(Label2.Caption, "standard")
End Sub

Private Sub Label3_Change()
Label3.Caption = Format(Label3.Caption, "standard")
End Sub

Private Sub Label4_Change()
Label4.Caption = Format(Label4.Caption, "standard")

End Sub

Sub DatagridWidth()

DataGrid1.Columns("LEGAJO").Width = 915.0237
DataGrid1.Columns("codAlimento").Width = 915.0237
DataGrid1.Columns("IDCATEGORIA").Width = 915.0237
DataGrid1.Columns("Descripcion Alimento").Width = 1739.906
DataGrid1.Columns("Cantidad").Width = 915.0237
DataGrid1.Columns("Hidratos de Carbono (gr.)").Width = 1470.047
DataGrid1.Columns("Proteinas (gr.)").Width = 1094.74
DataGrid1.Columns("LIPIDOS (gr.)").Width = 1065.26
DataGrid1.Columns("Kcal").Width = 1065.26
DataGrid1.Columns("HC1").Width = 1065.26
DataGrid1.Columns("PROT1").Width = 1065.26
DataGrid1.Columns("LIP1").Width = 1065.26

DataGrid2.Columns(0).Width = 1739.906
DataGrid2.Columns(1).Width = 750.0473
DataGrid2.Columns(2).Width = 750.0473
DataGrid2.Columns(3).Width = 750.0473
DataGrid2.Columns(4).Width = 750.0473
DataGrid2.Columns(5).Width = 750.0473
DataGrid2.Columns(6).Width = 750.0473
DataGrid2.Columns(7).Width = 750.0473
DataGrid2.Columns(8).Width = 750.0473
DataGrid2.Columns(9).Width = 750.0473
DataGrid2.Columns(10).Width = 750.0473
DataGrid2.Columns(11).Width = 750.0473
DataGrid2.Columns(12).Width = 750.0473
DataGrid2.Columns(13).Width = 750.0473
DataGrid2.Columns(14).Width = 750.0473
DataGrid2.Columns(15).Width = 750.0473
DataGrid2.Columns(16).Width = 750.0473

End Sub

Private Sub cmd_guardar_Click()
   
MousePointer = vbHourglass

Call f_GuardarCambios

Call f_Boton_Zorder

MousePointer = vbDefault

End Sub


Private Sub cmdAnterior_Click()
'If MDIForm1.ActiveForm.Data1.Recordset.BOF = False Then
    MDIForm1.ActiveForm.Adodc2.Recordset.MovePrevious
'Else
'    MDIForm1.ActiveForm.Data1.Recordset.MoveLast
'End If

If MDIForm1.ActiveForm.Adodc2.Recordset.AbsolutePosition = 0 Then

    cmdAnterior.Enabled = False
    cmdPrimero.Enabled = False
    
Else
    
    cmdSiguiente.Enabled = True
    cmdUltimo.Enabled = True

End If

End Sub


Private Sub cmd_salir_Click()

MousePointer = vbHourglass

Call f_finalizaOperacion
    
Call f_Boton_Zorder

MousePointer = vbDefault

End Sub

Sub f_GuardarCambios()
Dim strMsg As String

If estadoAbm = 3 Then

    strMsg = MsgBox("¿Esta seguro que desea guardar los cambios realizados?", vbYesNo)

    If strMsg = vbYes Then
        
        MousePointer = vbHourglass
           
        dbdiet.Execute "delete * from alimenxpaciente where alimenxpaciente.legajo = " & DataCombo1.BoundText
        
        'dbdiet.Execute "insert into alimenxpaciente select tmp_legajo as legajo, tmp_codalimento as codalimento, tmp_cantidad as cantidad, tmp_hc as hc, tmp_prot as prot, tmp_lip as lip, tmp_kcal as kcal from alimenxpaciente_tmp where alimenxpaciente_tmp.tmp_cantidad > 0"
        dbdiet.Execute ("csl_tmp_a_AlimenXPacientes")
        
        'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
        dbdiet.Execute ("csl_tmp_a_AlimenXPacientes")
        'dbdiet.Execute "insert into alimenxpaciente select tmp_legajo as legajo, tmp_codalimento as codalimento, tmp_cantidad as cantidad, tmp_hc as hc, tmp_prot as prot, tmp_lip as lip, tmp_kcal as kcal from alimenxpaciente_tmp where alimenxpaciente_tmp.tmp_cantidad > 0"
                
        estadoAbm = 1
        
        Me.cmd_Guardar.Enabled = False
        Me.cmd_Cancelar.Enabled = False

        
        Me.cmd_Guardar.Enabled = False
        Me.cmd_Cancelar.Enabled = False

        MousePointer = vbDefault
        
    End If

Else

    strMsg = MsgBox("No se han realizado cambios", vbInformation)

End If

End Sub

Sub f_DeshacerCambios()

Dim strMsg As String

If estadoAbm = 3 Then
    strMsg = MsgBox("¿Esta seguro que desea deshacer los cambios realizados?", vbYesNo)
Else
    strMsg = MsgBox("No se han realizado cambios", vbInformation)
End If

If strMsg = vbYes Then
    MousePointer = vbHourglass
    
    If estadoAbm = 3 Then
        
        dbdiet.Execute "delete * from alimenxpaciente_tmp"
            
        dbdiet.Execute "insert into alimenxpaciente_tmp select legajo as tmp_legajo, codalimento as tmp_codalimento, cantidad as tmp_cantidad, hc as tmp_hc, prot as tmp_prot, lip as tmp_lip, kcal as tmp_kcal from alimenxpaciente where alimenxpaciente.legajo = " & DataCombo1.BoundText '& " and alimenxpaciente.codalimento = " & fechaMenu
    
        dbdiet.Execute "insert into alimenxpaciente_tmp (tmp_legajo, tmp_codalimento) select " & DataCombo1.BoundText & ", codalimento from alimentos where estado = true"
                               
        'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
        dbdiet.Execute "insert into alimenxpaciente_tmp (tmp_legajo, tmp_codalimento) select " & DataCombo1.BoundText & ", codalimento from alimentos where estado = true"
        
    End If
    
    Call DatagridRefresh(Adodc1, DataGrid1)
    
    adodc_Totales_Nutrientes.Refresh
            
    estadoAbm = 1
    
    Me.cmd_Guardar.Enabled = False
    Me.cmd_Cancelar.Enabled = False
    
    'Me.cmd_Guardar.Enabled = False
    'Me.cmd_cancelar.Enabled = False

    MousePointer = vbDefault
End If

End Sub

Sub f_finalizaOperacion()
Dim strMsg As String

strMsg = vbNo

If estadoAbm = 3 Then
    strMsg = MsgBox("¿Esta seguro que desea finalizar la operacion?" & vbCrLf & vbTab & "- Se perderan los cambios realizados", vbYesNo)
Else
    strMsg = MsgBox("¿Esta seguro que desea finalizar la operacion?", vbYesNo)
End If

If strMsg = vbYes Then
    
    dbdiet.Execute "delete * from alimenxpaciente_tmp"
    
    'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
    dbdiet.Execute "delete * from alimenxpaciente_tmp"
    
    Me.adodc_Totales_Nutrientes.Refresh
            
    cmd_salir.Enabled = False
    cmd_Guardar.Enabled = False
    cmd_Cancelar.Enabled = False
    cmdImprimir.Enabled = False
    
    Frame(0).Enabled = True
    DataList1.Enabled = False
    DataGrid1.Enabled = False
        
    DataList1.BackColor = &H8000000F
    DataGrid1.BackColor = &H8000000F
    DataGrid2.BackColor = &H8000000F
    
    cmd_Aceptar.Enabled = True
    cmd_Cerrar.Enabled = True
    cmd_Tipito.Enabled = True
    'Frame1.Visible = False
    Frame1.Enabled = False
    
    estadoAbm = 1 ' el estado del form es "sin cambios"

    Call DatagridRefresh(Adodc1, DataGrid1)
           
End If

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Adodc1.Recordset = Nothing
Set Me.Adodc2.Recordset = Nothing
Set Me.Adodc3.Recordset = Nothing
Set Me.Adodc4.Recordset = Nothing
Set Me.Adodc5.Recordset = Nothing
Set Me.Adodc6.Recordset = Nothing
Set Me.adodc_Totales_Nutrientes.Recordset = Nothing

strQuery = "csl_alimentosxpacientes"
Call f_Adodc_ConnectionString(Adodc1, strQuery)

strQuery = "select *, (apell & ', ' & nombre) as nom from pacientes order by apell, nombre"
Call f_Adodc_ConnectionString(Adodc2, strQuery)

strQuery = "select * from Categoria"
Call f_Adodc_ConnectionString(Adodc3, strQuery)
           
strQuery = "select sum(alimenxpaciente.hc) as sumhc, sum(alimenxpaciente.prot) as sumprot, sum(alimenxpaciente.lip) as sumlip, sum(alimenxpaciente.kcal) as sumkcal from alimenxpaciente"
Call f_Adodc_ConnectionString(Adodc4, strQuery)

strQuery = "select rctideal, hcg, protg, lipg from pacientes"
Call f_Adodc_ConnectionString(Adodc5, strQuery)

strQuery = "select * from Pacientes"
Call f_Adodc_ConnectionString(Adodc6, strQuery)

strQuery = "SELECT * FROM csl_Totales_Nutrientes"
Call f_Adodc_ConnectionString(adodc_Totales_Nutrientes, strQuery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo1, Adodc6, Adodc2, "Legajo", "Legajo", "nom")

Call f_Enlaza_ControlData(DataCombo2, Adodc3, Adodc3, "idCategoria", "idCategoria", "Decripcion")

Call f_Enlaza_ControlData(DataList1, Adodc3, Adodc3, "idCategoria", "idCategoria", "Decripcion")
'==============================================

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

If Me.cmd_Guardar.Enabled = True Then
    Me.Pic_Guardar.ZOrder 0
Else
    Me.Pic_Guardar_Gris.ZOrder 0
End If

If Me.cmd_Cancelar.Enabled = True Then
    Me.Pic_Cancelar.ZOrder 0
Else
    Me.Pic_Cancelar_Gris.ZOrder 0
End If

If Me.cmd_salir.Enabled = True Then
    Me.Pic_Salir.ZOrder 0
Else
    Me.Pic_Salir_Gris.ZOrder 0
End If

If Me.cmd_Aceptar.Enabled = True Then
    Me.Pic_Aceptar.ZOrder 0
Else
    Me.Pic_Aceptar_Gris.ZOrder 0
End If

If Me.cmd_Cerrar.Enabled = True Then
    Me.Pic_Cerrar.ZOrder 0
Else
    Me.Pic_Cerrar_Gris.ZOrder 0
End If

Me.cmdImprimir.ZOrder 1
Me.cmd_Aceptar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_Cerrar.ZOrder 1
Me.cmd_Guardar.ZOrder 1
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Aceptar()

Me.cmd_Aceptar.ZOrder 0
Me.cmd_Cerrar.ZOrder 1

End Sub

Sub f_Cerrar()

Me.cmd_Aceptar.ZOrder 1
Me.cmd_Cerrar.ZOrder 0

End Sub

Sub f_Guardar()

Me.cmdImprimir.ZOrder 1
Me.cmd_Guardar.ZOrder 0
Me.cmd_Cancelar.ZOrder 1
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Cancelar()

Me.cmdImprimir.ZOrder 1
Me.cmd_Guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 0
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Salir()

Me.cmdImprimir.ZOrder 1
Me.cmd_Guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_salir.ZOrder 0

End Sub

Sub f_Imprimir()

Me.cmdImprimir.ZOrder 0
Me.cmd_Guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Tipito()

Me.cmd_Tipito.ZOrder 0

End Sub

Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Aceptar

End Sub

Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cancelar

End Sub

Private Sub Pic_Cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cerrar


End Sub

Private Sub Pic_Guardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Guardar


End Sub

Private Sub Pic_Imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Imprimir

End Sub

Private Sub Pic_Salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Salir

End Sub

Private Sub Pic_Tipito_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Tipito

End Sub

Private Sub TabStrip1_Click()
'define los valores de los frame correspondientes para que funcionen con el tabstrip
Dim a As String
a = TabStrip1.SelectedItem.Index

Select Case a
    Case Is = 1
    
        Me.fme_tabla.ZOrder 0
        Me.fme_Valores_Nutrientes.ZOrder 1
        Me.Fme_Totales_Nutrientes.ZOrder 1
                
    Case Is = 2
    
        Me.fme_tabla.ZOrder 1
        Me.fme_Valores_Nutrientes.ZOrder 0
        Me.Fme_Totales_Nutrientes.ZOrder 1
        
    Case Is = 3
    
        Me.fme_tabla.ZOrder 1
        Me.fme_Valores_Nutrientes.ZOrder 1
        Me.Fme_Totales_Nutrientes.ZOrder 0
        
End Select

End Sub
