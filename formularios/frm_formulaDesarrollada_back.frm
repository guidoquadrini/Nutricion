VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_formulaDesarrollada_back 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fórmula Desarrollada"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   FillStyle       =   0  'Solid
   Icon            =   "frm_formulaDesarrollada_back.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   10710
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   2160
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
      Connect         =   "FILE NAME=Alimentos anterior sin replica.UDL"
      OLEDBString     =   ""
      OLEDBFile       =   "Alimentos anterior sin replica.UDL"
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Pacientes"
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
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame 
      Caption         =   "Tabla de Alimentos:"
      Height          =   5295
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin VB.Frame Frame 
         Caption         =   "Paciente:"
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   10335
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frm_formulaDesarrollada_back.frx":0ECA
            DataField       =   "Legajo"
            DataSource      =   "Adodc6"
            Height          =   315
            Left            =   2640
            TabIndex        =   18
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
            TabIndex        =   20
            Top             =   120
            Width           =   2295
            Begin VB.CommandButton cmd_cerrar 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_formulaDesarrollada_back.frx":0EDF
               Height          =   375
               Left            =   1800
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frm_formulaDesarrollada_back.frx":1073
               Picture         =   "frm_formulaDesarrollada_back.frx":11C5
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Cancelar"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.CommandButton cmd_aceptar 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_formulaDesarrollada_back.frx":1678
               Height          =   375
               Left            =   1320
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frm_formulaDesarrollada_back.frx":17D1
               Picture         =   "frm_formulaDesarrollada_back.frx":1923
               Style           =   1  'Graphical
               TabIndex        =   23
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
               MouseIcon       =   "frm_formulaDesarrollada_back.frx":1BDF
               Picture         =   "frm_formulaDesarrollada_back.frx":1D31
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   22
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
               MouseIcon       =   "frm_formulaDesarrollada_back.frx":1FED
               Picture         =   "frm_formulaDesarrollada_back.frx":213F
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   21
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
               MouseIcon       =   "frm_formulaDesarrollada_back.frx":2440
               Picture         =   "frm_formulaDesarrollada_back.frx":2592
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   25
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
               MouseIcon       =   "frm_formulaDesarrollada_back.frx":26EB
               Picture         =   "frm_formulaDesarrollada_back.frx":283D
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   26
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton cmd_Tipito 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_formulaDesarrollada_back.frx":29D1
               Height          =   315
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frm_formulaDesarrollada_back.frx":30E1
               Style           =   1  'Graphical
               TabIndex        =   40
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
               MouseIcon       =   "frm_formulaDesarrollada_back.frx":3371
               Picture         =   "frm_formulaDesarrollada_back.frx":34C3
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   42
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
               MouseIcon       =   "frm_formulaDesarrollada_back.frx":35F3
               Picture         =   "frm_formulaDesarrollada_back.frx":3745
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   41
               Top             =   120
               Width           =   315
            End
         End
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   120
         Top             =   4800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "D:\Dietetica\rpts\rep_formdesarrollada.rpt"
         PrintFileLinesPerPage=   60
      End
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "frm_formulaDesarrollada_back.frx":39D5
         DataField       =   "idCategoria"
         DataSource      =   "Adodc3"
         Height          =   3180
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   5609
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         ListField       =   "Decripcion"
         BoundColumn     =   "idCategoria"
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_formulaDesarrollada_back.frx":39EA
         Height          =   3375
         Left            =   2520
         TabIndex        =   1
         Top             =   960
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
      Begin VB.Label Label7 
         Caption         =   "Grupos de alimentos:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         DataField       =   "sumhc"
         DataSource      =   "Adodc4"
         Height          =   255
         Left            =   5520
         TabIndex        =   11
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         DataField       =   "sumprot"
         DataSource      =   "Adodc4"
         Height          =   255
         Left            =   6960
         TabIndex        =   10
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Label1"
         DataField       =   "sumlip"
         DataSource      =   "Adodc4"
         Height          =   255
         Left            =   8040
         TabIndex        =   9
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Label1"
         DataField       =   "sumkcal"
         DataSource      =   "Adodc4"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9120
         TabIndex        =   8
         Top             =   4440
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
         Left            =   9120
         TabIndex        =   7
         Top             =   4800
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
         Left            =   8040
         TabIndex        =   6
         Top             =   4800
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
         Left            =   6960
         TabIndex        =   5
         Top             =   4800
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
         Left            =   5520
         TabIndex        =   4
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Totales ........................................"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   4440
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Totales Ideales ............................"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   4800
         Width           =   2415
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
      Connect         =   "FILE NAME=Alimentos anterior sin replica.UDL"
      OLEDBString     =   ""
      OLEDBFile       =   "Alimentos anterior sin replica.UDL"
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select rctideal, hcg, protg, lipg from pacientes"
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
      Connect         =   "FILE NAME=Alimentos anterior sin replica.UDL"
      OLEDBString     =   ""
      OLEDBFile       =   "Alimentos anterior sin replica.UDL"
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frm_formulaDesarrollada_back.frx":39FF
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
      Connect         =   "FILE NAME=Alimentos anterior sin replica.UDL"
      OLEDBString     =   ""
      OLEDBFile       =   "Alimentos anterior sin replica.UDL"
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Categoria"
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
      Connect         =   "FILE NAME=Alimentos anterior sin replica.UDL"
      OLEDBString     =   ""
      OLEDBFile       =   "Alimentos anterior sin replica.UDL"
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "csl_alimentosxpacientes"
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
      Connect         =   "FILE NAME=Alimentos anterior sin replica.UDL"
      OLEDBString     =   ""
      OLEDBFile       =   "Alimentos anterior sin replica.UDL"
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *, (apell & "", "" & nombre) as nom from pacientes order by apell, nombre"
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
      TabIndex        =   12
      Top             =   -120
      Visible         =   0   'False
      Width           =   2415
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frm_formulaDesarrollada_back.frx":3AAF
         DataField       =   "idCategoria"
         DataSource      =   "Adodc3"
         Height          =   315
         Left            =   120
         TabIndex        =   13
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
      Left            =   4508
      TabIndex        =   27
      Top             =   5280
      Width           =   2295
      Begin VB.CommandButton cmd_Cancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_formulaDesarrollada_back.frx":3AC4
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_formulaDesarrollada_back.frx":3C3C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Deshacer cambios"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmd_guardar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_formulaDesarrollada_back.frx":40BC
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_formulaDesarrollada_back.frx":4215
         Style           =   1  'Graphical
         TabIndex        =   33
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
         MouseIcon       =   "frm_formulaDesarrollada_back.frx":44D1
         Picture         =   "frm_formulaDesarrollada_back.frx":4623
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   28
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
         MouseIcon       =   "frm_formulaDesarrollada_back.frx":48DF
         Picture         =   "frm_formulaDesarrollada_back.frx":4A31
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   29
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
         MouseIcon       =   "frm_formulaDesarrollada_back.frx":4EB1
         Picture         =   "frm_formulaDesarrollada_back.frx":5003
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   36
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
         MouseIcon       =   "frm_formulaDesarrollada_back.frx":515C
         Picture         =   "frm_formulaDesarrollada_back.frx":52AE
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   34
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_salir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_formulaDesarrollada_back.frx":5426
         Height          =   375
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_formulaDesarrollada_back.frx":55BA
         Style           =   1  'Graphical
         TabIndex        =   32
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
         MouseIcon       =   "frm_formulaDesarrollada_back.frx":58BB
         Picture         =   "frm_formulaDesarrollada_back.frx":5A0D
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   30
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
         MouseIcon       =   "frm_formulaDesarrollada_back.frx":5D0E
         Picture         =   "frm_formulaDesarrollada_back.frx":5E60
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   35
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_formulaDesarrollada_back.frx":5FF4
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_formulaDesarrollada_back.frx":614C
         Style           =   1  'Graphical
         TabIndex        =   38
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
         MouseIcon       =   "frm_formulaDesarrollada_back.frx":65CC
         Picture         =   "frm_formulaDesarrollada_back.frx":671E
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   37
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
         MouseIcon       =   "frm_formulaDesarrollada_back.frx":6B9E
         Picture         =   "frm_formulaDesarrollada_back.frx":6CF0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   39
         Top             =   120
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_formulaDesarrollada_back"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
Dim bandera As Integer

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

Private Sub Command1_Click()
'''Dim sIniPath As String
''''sIniPath = "c:\gustavo\sin.ini"
''''If Trim$(Command$) <> "" Then
''''    sIniPath = Command$
''''    MsgBox "Verdadero sIniPath: " & sIniPath
''''Else
''''    sIniPath = Command$
''''    MsgBox "Falso sIniPath: " & sIniPath
''''End If
'''Dim fileSys As New FileSystemObject, fil As File
'''Dim drv As Drive
'''Set drv = fileSys.GetDrive(fileSys.GetDriveName("c:"))
''''Set fil = fileSys.GetFile("c:\dietetica\Alimentos anterior sin replica.udl")
'''If fileSys.FileExists("c:\gustavo\sin.txt") = True Then
'''    Set fil = fileSys.GetFile("c:\gustavo\sin.txt")
'''    MsgBox "fil.Size: " & fil.Size & "drv.DriveLetter: " & drv.DriveLetter & " drv.DriveType: " & drv.DriveType & " drv.SerialNumber: " & drv.SerialNumber & "drv.VolumeName: " & drv.VolumeName
'''    MsgBox fil.OpenAsTextStream.ReadAll
'''    MsgBox fil.DateCreated
'''End If
'''
''''MsgBox fileSys.GetDriveName

'dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(MDIForm1.ActiveForm.Label1.Caption) & ", codalimento from alimentos where estado = true"

Call DatagridRefresh(Adodc1, DataGrid1)

End Sub


Private Sub cmd_tipito_Click()
Unload frmPacientes
frmPacientes.Show
frmPacientes.Data1.Recordset.FindFirst " legajo = " & DataCombo1.BoundText

End Sub

Private Sub cmd_aceptar_Click()
Dim strquery As String

MousePointer = vbHourglass

If DataCombo1.Text <> "" Then

    dbdiet.Execute "insert into alimenxpaciente_tmp select legajo as tmp_legajo, codalimento as tmp_codalimento, cantidad as tmp_cantidad, hc as tmp_hc, prot as tmp_prot, lip as tmp_lip, kcal as tmp_kcal from alimenxpaciente where alimenxpaciente.legajo = " & DataCombo1.BoundText '& " and alimenxpaciente.codalimento = " & fechaMenu
    
    dbdiet.Execute "insert into alimenxpaciente_tmp (tmp_legajo, tmp_codalimento) select " & DataCombo1.BoundText & ", codalimento from alimentos where estado = true"
    
    'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
    dbdiet.Execute "insert into alimenxpaciente_tmp (tmp_legajo, tmp_codalimento) select " & DataCombo1.BoundText & ", codalimento from alimentos where estado = true"
          
    'strquery = "select * from consultaprueba3 where alimenxpaciente.legajo = " & DataCombo1.BoundText
    strquery = "select * from csl_alimentosxpacientes where alimenxpaciente_tmp.tmp_legajo = " & DataCombo1.BoundText '& " and alimentos.estado = 1"
    
    With Adodc1
        .RecordSource = strquery
        .Refresh
    End With
    
    With DataGrid1
        .ReBind
        .Refresh
    End With
               
    Call DatagridRefresh(Adodc1, DataGrid1)
    
    strquery = "select rctideal, hcg, protg, lipg from pacientes where legajo = " & DataCombo1.BoundText
        
    With Adodc5
        .RecordSource = strquery
        .Refresh
    End With

    For i = 0 To 3
        lbl(i).Caption = Format(Val(lbl(i).Caption), "standard")
    Next

    Frame(0).Enabled = False
    DataList1.Enabled = True
    DataGrid1.Enabled = True
        
    DataList1.BackColor = &H80000005
    DataGrid1.BackColor = &H80000005
    
    cmd_salir.Enabled = True
    cmdImprimir.Enabled = True
    'cmdImprimir.Enabled = False
    cmd_aceptar.Enabled = False
    cmd_cerrar.Enabled = False
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
    Me.Caption = " Fórmula Desarrollada " & " - " & DataCombo1.Text
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

strquery = " select * from consultaprueba3 where idcategoria = " & DataCombo2.BoundText '& " and legajo = " & DataCombo1.BoundText
Set tb = dbdiet.OpenRecordset(strquery)
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

Private Sub Datagrid1_AfterColUpdate(ByVal ColIndex As Integer)
Dim hc, prot, lip, Kcal
Dim strquery As String

'"hidratos de carbono" --> columna 5
'"proteínas" ------------> columna 6
'"lípidos" --------------> columna 7
'"kcal" -----------------> columna 8
    
If Adodc1.Recordset.RecordCount > 0 Then

    If DataGrid1.Columns(4) <> "" Then
    
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
            
    ''    estadoAbm = 3
               
        Call calculaTotales
        
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

End Sub

Private Sub DataGrid1_Change()
    
estadoAbm = 3

Me.cmd_guardar.Enabled = True
Me.cmd_Cancelar.Enabled = True

Call f_Boton_Zorder

End Sub

Private Sub DataGrid1_LostFocus()

Call Datagrid1_AfterColUpdate(4)

End Sub

Private Sub DataList1_Click()
Dim aa As Integer
aa = 0

'strquery = "select * from consultaprueba3 where alimenxpaciente.legajo = " & DataCombo1.BoundText '& " and alimentos.estado = 1"
strquery = "select * from csl_alimentosxpacientes where alimenxpaciente_tmp.tmp_legajo = " & DataCombo1.BoundText '& " and alimentos.estado = 1"

Set tb = dbdiet.OpenRecordset(strquery)
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
Me.Height = 6255
Me.Width = 10800 ' 11265
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

End Sub

Private Sub Form_Load()
Dim strquery As String
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

    strquery = "select * from csl_alimentosxpacientes where alimenxpaciente_tmp.tmp_legajo = " & DataCombo1.BoundText '& " and alimentos.estado = 1"

    With Adodc1
        .RecordSource = strquery
        .Refresh
    End With

    With DataGrid1
        .ReBind
        .Refresh
    End With
        
    strquery = "select rctideal, hcg, protg, lipg from pacientes where legajo = " & DataCombo1.BoundText
    
    With Adodc5
        .RecordSource = strquery
        .Refresh
    End With
    
    For i = 0 To 3
        lbl(i).Caption = Format(Val(lbl(i).Caption), "standard")
    Next
    
    Call calculaTotales

End If

Frame(0).Enabled = True
DataList1.Enabled = False
DataGrid1.Enabled = False

bandera = 1

estadoAbm = 1

Call f_Boton_Zorder

End Sub
Private Sub calculaTotales()
Dim strquery As String

'strquery = " select sum(alimenxpaciente.hc) as sumhc, sum(alimenxpaciente.prot) as sumprot, sum(alimenxpaciente.lip) as sumlip, sum(alimenxpaciente.kcal) as sumkcal from alimenxpaciente where legajo = " & DataCombo1.BoundText
strquery = " select sum(alimenxpaciente_tmp.tmp_hc) as sumhc, sum(alimenxpaciente_tmp.tmp_prot) as sumprot, sum(alimenxpaciente_tmp.tmp_lip) as sumlip, sum(alimenxpaciente_tmp.tmp_kcal) as sumkcal from alimenxpaciente_tmp where tmp_legajo = " & DataCombo1.BoundText

With Adodc4
    .RecordSource = strquery
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
        
        dbdiet.Execute "insert into alimenxpaciente select tmp_legajo as legajo, tmp_codalimento as codalimento, tmp_cantidad as cantidad, tmp_hc as hc, tmp_prot as prot, tmp_lip as lip, tmp_kcal as kcal from alimenxpaciente_tmp where alimenxpaciente_tmp.tmp_cantidad > 0"
                
        'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
        dbdiet.Execute "insert into alimenxpaciente select tmp_legajo as legajo, tmp_codalimento as codalimento, tmp_cantidad as cantidad, tmp_hc as hc, tmp_prot as prot, tmp_lip as lip, tmp_kcal as kcal from alimenxpaciente_tmp where alimenxpaciente_tmp.tmp_cantidad > 0"
                
        estadoAbm = 1
        
        Me.cmd_guardar.Enabled = False
        Me.cmd_Cancelar.Enabled = False

        
        Me.cmd_guardar.Enabled = False
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
    
    estadoAbm = 1
    
    Me.cmd_guardar.Enabled = False
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
    
    cmd_salir.Enabled = False
    cmd_guardar.Enabled = False
    cmd_Cancelar.Enabled = False
    cmdImprimir.Enabled = False
    
    Frame(0).Enabled = True
    DataList1.Enabled = False
    DataGrid1.Enabled = False
        
    DataList1.BackColor = &H8000000F
    DataGrid1.BackColor = &H8000000F
    
    cmd_aceptar.Enabled = True
    cmd_cerrar.Enabled = True
    cmd_Tipito.Enabled = True
    'Frame1.Visible = False
    Frame1.Enabled = False
    
    estadoAbm = 1 ' el estado del form es "sin cambios"

    Call DatagridRefresh(Adodc1, DataGrid1)
           
End If

End Sub

Sub f_CargarOrigenDatos()
Dim strquery As String
strquery = ""

strquery = "csl_alimentosxpacientes"
Call f_Adodc_ConnectionString(Adodc1, strquery)

strquery = "select *, (apell & ', ' & nombre) as nom from pacientes order by apell, nombre"
Call f_Adodc_ConnectionString(Adodc2, strquery)

strquery = "select * from Categoria"
Call f_Adodc_ConnectionString(Adodc3, strquery)
           
strquery = "select sum(alimenxpaciente.hc) as sumhc, sum(alimenxpaciente.prot) as sumprot, sum(alimenxpaciente.lip) as sumlip, sum(alimenxpaciente.kcal) as sumkcal from alimenxpaciente"
Call f_Adodc_ConnectionString(Adodc4, strquery)

strquery = "select rctideal, hcg, protg, lipg from pacientes"
Call f_Adodc_ConnectionString(Adodc5, strquery)

strquery = "select * from Pacientes"
Call f_Adodc_ConnectionString(Adodc6, strquery)

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

If Me.cmd_guardar.Enabled = True Then
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

If Me.cmd_aceptar.Enabled = True Then
    Me.Pic_Aceptar.ZOrder 0
Else
    Me.Pic_Aceptar_Gris.ZOrder 0
End If

If Me.cmd_cerrar.Enabled = True Then
    Me.Pic_Cerrar.ZOrder 0
Else
    Me.Pic_Cerrar_Gris.ZOrder 0
End If

Me.cmdImprimir.ZOrder 1
Me.cmd_aceptar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_cerrar.ZOrder 1
Me.cmd_guardar.ZOrder 1
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Aceptar()

Me.cmd_aceptar.ZOrder 0
Me.cmd_cerrar.ZOrder 1

End Sub

Sub f_Cerrar()

Me.cmd_aceptar.ZOrder 1
Me.cmd_cerrar.ZOrder 0

End Sub

Sub f_Guardar()

Me.cmdImprimir.ZOrder 1
Me.cmd_guardar.ZOrder 0
Me.cmd_Cancelar.ZOrder 1
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Cancelar()

Me.cmdImprimir.ZOrder 1
Me.cmd_guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 0
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Salir()

Me.cmdImprimir.ZOrder 1
Me.cmd_guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_salir.ZOrder 0

End Sub

Sub f_Imprimir()

Me.cmdImprimir.ZOrder 0
Me.cmd_guardar.ZOrder 1
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
