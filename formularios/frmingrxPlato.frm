VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmingrxPlato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingredientes por Platos"
   ClientHeight    =   4455
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmingrxPlato.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6135
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5400
      Top             =   840
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
   Begin VB.Frame fme_body 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   5895
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   4200
         Top             =   1560
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   2400
         Top             =   1560
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
      Begin MSAdodcLib.Adodc ado1_datalist1 
         Height          =   330
         Left            =   240
         Top             =   1560
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
         Caption         =   "ado1_datalist1"
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
      Begin MSDataListLib.DataList DataList1 
         Height          =   1035
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1826
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   1035
         Left            =   2400
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1826
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid Datagrid1 
         Bindings        =   "frmingrxPlato.frx":0ECA
         Height          =   1935
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "tmp_Descrip_Categoria"
            Caption         =   "Grupo alimentario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0 ""$"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "tmp_Descrip_Alimento"
            Caption         =   "Ingrediente"
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
            DataField       =   "tmp_Porcion"
            Caption         =   "Cantidad por pocion (gr/cc)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "tmp_idPlato"
            Caption         =   "idPlato"
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
            DataField       =   "tmp_CodAlimento"
            Caption         =   "codAlimento"
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
               ColumnAllowSizing=   -1  'True
               Button          =   -1  'True
               ColumnWidth     =   2145.26
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   -1  'True
               Button          =   -1  'True
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
               ColumnWidth     =   1065.26
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   495
         Left            =   1935
         TabIndex        =   12
         Top             =   2640
         Width           =   2295
         Begin VB.CommandButton cmd_Imprimir 
            Appearance      =   0  'Flat
            DisabledPicture =   "frmingrxPlato.frx":0EDF
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmingrxPlato.frx":1037
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Imprimir"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmd_Salir 
            Appearance      =   0  'Flat
            DisabledPicture =   "frmingrxPlato.frx":14B7
            Height          =   375
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmingrxPlato.frx":164B
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmd_Guardar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frmingrxPlato.frx":194C
            Enabled         =   0   'False
            Height          =   375
            Left            =   720
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmingrxPlato.frx":1AA5
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Guardar cambios"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmd_Cancelar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frmingrxPlato.frx":1D61
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmingrxPlato.frx":1ED9
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Deshacer cambios"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.PictureBox Pic_Salir_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1800
            MouseIcon       =   "frmingrxPlato.frx":2359
            Picture         =   "frmingrxPlato.frx":24AB
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   21
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
            MouseIcon       =   "frmingrxPlato.frx":263F
            Picture         =   "frmingrxPlato.frx":2791
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   22
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
            MouseIcon       =   "frmingrxPlato.frx":2909
            Picture         =   "frmingrxPlato.frx":2A5B
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   23
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Salir 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1800
            MouseIcon       =   "frmingrxPlato.frx":2BB4
            Picture         =   "frmingrxPlato.frx":2D06
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   18
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
            MouseIcon       =   "frmingrxPlato.frx":3007
            Picture         =   "frmingrxPlato.frx":3159
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   19
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Guardar 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   720
            MouseIcon       =   "frmingrxPlato.frx":35D9
            Picture         =   "frmingrxPlato.frx":372B
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   20
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Imprimir 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            MouseIcon       =   "frmingrxPlato.frx":39E7
            Picture         =   "frmingrxPlato.frx":3B39
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   24
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
            MouseIcon       =   "frmingrxPlato.frx":3FB9
            Picture         =   "frmingrxPlato.frx":410B
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   17
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total de Kcal:"
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   29
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label lbl_Kcal 
         Alignment       =   1  'Right Justify
         Caption         =   "lbl_Kcal"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4680
         TabIndex        =   28
         Top             =   2160
         Width           =   975
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   5880
         Y1              =   2520
         Y2              =   2520
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalles"
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   120
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame fme_header 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   5655
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frmingrxPlato.frx":4263
            DataField       =   "idPlato"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   855
            TabIndex        =   2
            Top             =   165
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "NombrePlato"
            BoundColumn     =   "idPlato"
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
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   495
            Left            =   3975
            TabIndex        =   4
            Top             =   0
            Width           =   1695
            Begin VB.CommandButton cmd_Aceptar_header 
               Appearance      =   0  'Flat
               DisabledPicture =   "frmingrxPlato.frx":4278
               Height          =   375
               Left            =   720
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frmingrxPlato.frx":43D1
               Picture         =   "frmingrxPlato.frx":4523
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Aceptar"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.CommandButton cmd_Cancelar_header 
               Appearance      =   0  'Flat
               DisabledPicture =   "frmingrxPlato.frx":47DF
               Height          =   375
               Left            =   1200
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frmingrxPlato.frx":4973
               Picture         =   "frmingrxPlato.frx":4AC5
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Cancelar"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.PictureBox Pic_Cancelar_header 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1200
               MouseIcon       =   "frmingrxPlato.frx":4F78
               Picture         =   "frmingrxPlato.frx":50CA
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   9
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Aceptar_header 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   720
               MouseIcon       =   "frmingrxPlato.frx":53CB
               Picture         =   "frmingrxPlato.frx":551D
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   10
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Cancelar_Gris_header 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1200
               MouseIcon       =   "frmingrxPlato.frx":57D9
               Picture         =   "frmingrxPlato.frx":592B
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   8
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Aceptar_gris_header 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   720
               MouseIcon       =   "frmingrxPlato.frx":5ABF
               Picture         =   "frmingrxPlato.frx":5C11
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   6
               Top             =   120
               Width           =   375
            End
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Plato:"
            Height          =   195
            Index           =   1
            Left            =   375
            TabIndex        =   3
            Top             =   225
            Width           =   405
         End
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6000
         Y1              =   960
         Y2              =   960
      End
   End
End
Attribute VB_Name = "frmingrxPlato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msg As String
Dim tb As Recordset
Dim lDelete As Boolean
'Public estadoAbm As Integer ' define el estado de un formulario de abm
'                             1 = sin cambios; 2 = agregar; 3 = modificar
'el modulo "fSetEnableFields(MDIForm1.ActiveForm, vbFalse)" se debe agregar al proyecto
Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
                            
Private Sub cmd_Aceptar_header_Click()

MousePointer = vbHourglass

Call f_Carga_Datagrid

Call f_TotalKcal

fme_header.Enabled = False
fme_body.Enabled = True

Me.cmd_Salir.SetFocus

Call f_Boton_Zorder

MousePointer = vbDefault

End Sub

Private Sub cmd_Cancelar_Click()

MousePointer = vbHourglass

Call f_DeshacerCambios

Call f_TotalKcal

lDelete = False

Call f_Boton_Zorder

MousePointer = vbDefault

End Sub

Private Sub cmd_Cancelar_header_Click()

Unload Me

End Sub

Private Sub cmd_guardar_Click()

MousePointer = vbHourglass

'actualiza los registros de la DB adelantando un registro
Adodc3.Recordset.MoveNext

If Adodc3.Recordset.EOF = True Then
    Adodc3.Recordset.MoveLast
End If
'--------------------------------------------------------

Call f_Valida_Datos

Call f_GuardarCambios

Call f_TotalKcal

lDelete = False

Call f_Boton_Zorder

MousePointer = vbDefault

End Sub



Private Sub cmd_Imprimir_Click()
Dim strQuery As String
'aclare el filtro para imprimir
msg = MsgBox("¿Desea imprimir todos los registros?", vbYesNo, "Imprimir")
  
CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_alimxplato_one.rpt"

If msg = vbYes Then
    
    strQuery = ""
    
Else
    
    strQuery = " {ingredientesplatos.idplato} = " & DataCombo1.BoundText
    
End If

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub

Private Sub cmd_salir_Click()
MousePointer = vbHourglass

Call f_finalizaOperacion

lDelete = False

Call f_Boton_Zorder

MousePointer = vbDefault

End Sub




Private Sub DataCombo1_Click(Area As Integer)

Me.Caption = "Ingredientes por Platos - Nro. " & DataCombo1.BoundText

End Sub


Private Sub Datagrid1_AfterDelete()

lDelete = True

estadoAbm = 3

Me.cmd_Guardar.Enabled = True
Me.cmd_Cancelar.Enabled = True

Call f_Boton_Zorder

End Sub

Private Sub Datagrid1_AfterUpdate()

Call f_TotalKcal

End Sub

Private Sub Datagrid1_BeforeUpdate(Cancel As Integer)
Dim strQuery As String

If lDelete = False Then
    
    If Adodc3.Recordset.Fields("tmp_idPlato").Value > 0 And Adodc3.Recordset.Fields("tmp_CodAlimento").Value > 0 And Adodc3.Recordset.Fields("tmp_Porcion").Value > 0 Then
        
        Call f_Valida_Datos
                        
    Else
        
        MsgBox "Los datos ingresados no son correctos, Verifique.", vbInformation
        
        Cancel = True
        
    End If

End If

lDelete = False

End Sub

Private Sub Datagrid1_ButtonClick(ByVal ColIndex As Integer)

Select Case ColIndex
    Case Is = 0
    
            With DataList1
                
                .Left = Datagrid1.Columns(ColIndex).Left '+ 100
                .Width = Datagrid1.Columns(ColIndex).Width '+ 50
                        
                .Top = Datagrid1.Top + Datagrid1.RowTop(Datagrid1.Row) + Datagrid1.RowHeight
                
                'en caso de error continuo con la siguiente intruccion
                'ya que cuando estoy agregando un registro la siguiente
                'intruccion provoca un error
                On Error Resume Next
                .BoundText = f_Devuelve_idCategoria(Datagrid1.Columns(4).Value) ' devuelve el idCategoria para el codAlimento dado
                
                .Visible = True
                .SetFocus
            End With
    
    Case Is = 1
            
            With DataList2
            
                .Left = Datagrid1.Columns(ColIndex).Left '+ 100
                .Width = Datagrid1.Columns(ColIndex).Width '+ 50
            
                .Top = Datagrid1.Top + Datagrid1.RowTop(Datagrid1.Row) + Datagrid1.RowHeight
                                                
                'en caso de error continuo con la siguiente intruccion
                'ya que cuando estoy agregando un registro la siguiente
                'intruccion provoca un error
                On Error Resume Next
                .BoundText = Datagrid1.Columns(4).Value 'CodAlimento
                                
                .Visible = True
                .SetFocus
            End With

End Select

End Sub



Private Sub DataGrid1_Change()

If Datagrid1.Columns(0) = "" Then

    strQuery = " SELECT * FROM alimentos ORDER BY DescripAlimento "
    
    With Adodc2
        .RecordSource = strQuery
        .Refresh
    End With

End If

estadoAbm = 3

Me.cmd_Guardar.Enabled = True
Me.cmd_Cancelar.Enabled = True

Call f_Boton_Zorder

End Sub


Private Sub Datagrid1_Click()

Me.DataList1.Visible = False
Me.DataList2.Visible = False

End Sub

Private Sub DataList1_Click()
Dim strQuery As String

'Grupo Alimentario
Datagrid1.Columns(0).Value = DataList1.Text

'ingrediente
Datagrid1.Columns(1).Value = ""

'codAlimento
Datagrid1.Columns(4).Value = Null

strQuery = " SELECT * FROM alimentos WHERE idcategoria = " & DataList1.BoundText

With Adodc2
    .RecordSource = strQuery
    .Refresh
End With

estadoAbm = 3

Me.cmd_Guardar.Enabled = True
Me.cmd_Cancelar.Enabled = True

Call f_Boton_Zorder

End Sub

Private Sub DataList1_LostFocus()

DataList1.Visible = False
'Datagrid1.SetFocus

End Sub

Private Sub DataList2_Click()
Dim strQuery As String

'Ingrediente
Datagrid1.Columns(1).Value = DataList2.Text

'idPlato
Datagrid1.Columns(3).Value = DataCombo1.BoundText

'codalimento
Datagrid1.Columns(4).Value = DataList2.BoundText

'actualizo el valor correspondiente del grupo alimentario segun el alimento seleccionado
    strQuery = " SELECT * FROM Alimentos, categoria WHERE CodAlimento = " & DataList2.BoundText & _
                " AND Alimentos.idCategoria = Categoria.idCategoria "
    
    Set tb = dbdiet.OpenRecordset(strQuery)
    
    'Grupo Alimentario
    Datagrid1.Columns(0).Value = tb.Fields("Decripcion").Value
    
    tb.Close
'--------------------------------------------------------------

estadoAbm = 3

Me.cmd_Guardar.Enabled = True
Me.cmd_Cancelar.Enabled = True

Call f_Boton_Zorder

End Sub

Private Sub DataList2_DblClick()
DataList2.Visible = False
Datagrid1.SetFocus
End Sub

Private Sub DataList2_LostFocus()

DataList2.Visible = False
Datagrid1.SetFocus

End Sub

Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 4830
Me.Width = 6225
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

End Sub

Private Sub DataCombo1_LostFocus()
If DataCombo1.Text = "" Then
    DataCombo1.SetFocus
    MsgBox "Debe Completar el Nombre del Plato", vbInformation, "Información"
End If

End Sub



Private Sub Form_Load()
Call f_Boton_Zorder

Call f_CargarOrigenDatos

Call f_DatagridWidth

lbl_Kcal.Caption = ""

estadoAbm = 1

lDelete = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Call cmdCancelar_Click

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub


Private Sub Pic_Aceptar_header_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Aceptar_header

End Sub

Private Sub Pic_Cancelar_header_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cancelar_header

End Sub

Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cancelar

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


Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Adodc1.Recordset = Nothing
Set Me.Adodc2.Recordset = Nothing
Set Me.Adodc3.Recordset = Nothing
Set Me.ado1_datalist1.Recordset = Nothing

strQuery = "select * from platos order by nombreplato"
Call f_Adodc_ConnectionString(Adodc1, strQuery)

strQuery = "SELECT alimentos.codalimento, alimentos.idcategoria, alimentos.descripalimento, " & _
            " alimentos.hc, alimentos.prot, alimentos.lip, alimentos.estado, " & _
            " FROM alimentos, categoria " & _
            " WHERE alimentos.idcategoria = categoria.idcategoria " & _
            " ORDER BY alimentos.descripalimento"
            
strQuery = " SELECT CodAlimento, idCategoria, DescripAlimento " & _
            " FROM alimentos ORDER BY DescripAlimento "

Call f_Adodc_ConnectionString(Adodc2, strQuery)

strQuery = " SELECT * " & _
            " FROM Categoria, alimentos, IngredientesPlatos_tmp " & _
            " WHERE IngredientesPlatos_tmp.tmp_codAlimento = alimentos.codAlimento " & _
            " AND alimentos.idcategoria = Categoria.idcategoria " & _
            " AND IngredientesPlatos_tmp.tmp_idplato = " & DataCombo1.BoundText & _
            " ORDER BY tmp_idplato, IngredientesPlatos_tmp.tmp_Descrip_Categoria, IngredientesPlatos_tmp.tmp_Descrip_Alimento"

'lo en lazo primero ya que utilizo el valor datacombo1.boundtext para la
'siguiente instruccion SQL
Call f_Enlaza_ControlData(DataCombo1, Adodc1, Adodc1, "idPlato", "idPlato", "NombrePlato")

strQuery = " SELECT * " & _
            " FROM IngredientesPlatos_tmp " & _
            " WHERE IngredientesPlatos_tmp.tmp_idplato = " & DataCombo1.BoundText & _
            " ORDER BY IngredientesPlatos_tmp.tmp_Descrip_Categoria, IngredientesPlatos_tmp.tmp_Descrip_Alimento"

Call f_Adodc_ConnectionString(Adodc3, strQuery)
Call DatagridRefresh(Adodc3, Datagrid1)

strQuery = " SELECT * FROM categoria ORDER BY decripcion "

Call f_Adodc_ConnectionString(ado1_datalist1, strQuery)

'Define propiedades de los controles enlazados

Call f_Enlaza_ControlData(DataList1, ado1_datalist1, ado1_datalist1, "idCategoria", "idCategoria", "Decripcion")

Call f_Enlaza_ControlData(DataList2, Adodc2, Adodc2, "CodAlimento", "CodAlimento", "descripalimento")
'==============================================

End Sub


Function f_Devuelve_idCategoria(nCodAlimento As Long) As Long
Dim strQuery As String

strQuery = " SELECT * FROM Alimentos WHERE CodAlimento = " & nCodAlimento
    
Set tb = dbdiet.OpenRecordset(strQuery)
    
'Grupo Alimentario
f_Devuelve_idCategoria = tb.Fields("idCategoria").Value

tb.Close

End Function

Sub f_Boton_Zorder()

If Me.cmd_Aceptar_header.Enabled = True Then
    Me.Pic_Aceptar_header.ZOrder 0
Else
    Me.Pic_Aceptar_gris_header.ZOrder 0
End If

If Me.cmd_Cancelar_header.Enabled = True Then
    Me.Pic_Cancelar_header.ZOrder 0
Else
    Me.Pic_Cancelar_Gris_header.ZOrder 0
End If

If Me.cmd_Imprimir.Enabled = True Then
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

If Me.cmd_Salir.Enabled = True Then
    Me.Pic_Salir.ZOrder 0
Else
    Me.Pic_Salir_Gris.ZOrder 0
End If

Me.cmd_Aceptar_header.ZOrder 1
Me.cmd_Cancelar_header.ZOrder 1
Me.cmd_Imprimir.ZOrder 1
Me.cmd_Guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_Salir.ZOrder 1

End Sub

Sub f_Aceptar_header()

Me.cmd_Aceptar_header.ZOrder 0
Me.cmd_Cancelar_header.ZOrder 1
Me.cmd_Imprimir.ZOrder 1
Me.cmd_Guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_Salir.ZOrder 1

End Sub

Sub f_Cancelar_header()

Me.cmd_Aceptar_header.ZOrder 1
Me.cmd_Cancelar_header.ZOrder 0
Me.cmd_Imprimir.ZOrder 1
Me.cmd_Guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_Salir.ZOrder 1

End Sub

Sub f_Imprimir()

Me.cmd_Imprimir.ZOrder 0
Me.cmd_Guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_Salir.ZOrder 1

End Sub

Sub f_Guardar()

Me.cmd_Imprimir.ZOrder 1
Me.cmd_Guardar.ZOrder 0
Me.cmd_Cancelar.ZOrder 1
Me.cmd_Salir.ZOrder 1

End Sub

Sub f_Cancelar()

Me.cmd_Imprimir.ZOrder 1
Me.cmd_Guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 0
Me.cmd_Salir.ZOrder 1

End Sub

Sub f_Salir()

Me.cmd_Imprimir.ZOrder 1
Me.cmd_Guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_Salir.ZOrder 0

End Sub

Sub f_Carga_Datagrid()
Dim strQuery As String

strQuery = " SELECT * FROM IngredientesPlatos WHERE idPlato = " & DataCombo1.BoundText

Set tb = dbdiet.OpenRecordset(strQuery)

If tb.RecordCount <> 0 Then
    
    dbdiet.Execute "insert into IngredientesPlatos_tmp (tmp_idPlato, tmp_CodAlimento, tmp_Porcion, tmp_Descrip_Categoria, tmp_Descrip_Alimento) select idPlato, CodAlimento, Porcion, Descrip_Categoria, Descrip_Alimento from IngredientesPlatos where idPlato = " & DataCombo1.BoundText
    
    'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
    dbdiet.Execute "insert into IngredientesPlatos_tmp (tmp_idPlato, tmp_CodAlimento, tmp_Porcion, tmp_Descrip_Categoria, tmp_Descrip_Alimento) select idPlato, CodAlimento, Porcion, Descrip_Categoria, Descrip_Alimento from IngredientesPlatos where idPlato = " & DataCombo1.BoundText
    
End If
    
strQuery = " SELECT * " & _
            " FROM IngredientesPlatos_tmp " & _
            " WHERE IngredientesPlatos_tmp.tmp_idplato = " & DataCombo1.BoundText & _
            " ORDER BY IngredientesPlatos_tmp.tmp_Descrip_Categoria, IngredientesPlatos_tmp.tmp_Descrip_Alimento"

Call f_Adodc_ConnectionString(Adodc3, strQuery)
Call DatagridRefresh(Adodc3, Datagrid1)

tb.Close

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
    
    dbdiet.Execute "delete * from IngredientesPlatos_tmp"
    
    'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
    dbdiet.Execute "delete * from IngredientesPlatos_tmp"
 
    estadoAbm = 1 ' el estado del form es "sin cambios"

    Call DatagridRefresh(Adodc3, Datagrid1)
    
    fme_header.Enabled = True
    fme_body.Enabled = False
    
    Me.cmd_Guardar.Enabled = False
    Me.cmd_Cancelar.Enabled = False
    
    Me.cmd_Cancelar_header.SetFocus
    
    lbl_Kcal.Caption = ""
    
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
        
    If estadoAbm = 3 Then
        
        dbdiet.Execute "delete * from IngredientesPlatos_tmp"
               
        dbdiet.Execute "insert into IngredientesPlatos_tmp (tmp_idPlato, tmp_CodAlimento, tmp_Porcion, tmp_Descrip_Categoria, tmp_Descrip_Alimento) select idPlato, CodAlimento, Porcion, Descrip_Categoria, Descrip_Alimento from IngredientesPlatos where idPlato = " & DataCombo1.BoundText
                                             
        'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
        dbdiet.Execute "insert into IngredientesPlatos_tmp (tmp_idPlato, tmp_CodAlimento, tmp_Porcion, tmp_Descrip_Categoria, tmp_Descrip_Alimento) select idPlato, CodAlimento, Porcion, Descrip_Categoria, Descrip_Alimento from IngredientesPlatos where idPlato = " & DataCombo1.BoundText
        
    End If
    
    Call DatagridRefresh(Adodc3, Datagrid1)
    
    estadoAbm = 1
    
    Me.cmd_Guardar.Enabled = False
    Me.cmd_Cancelar.Enabled = False
       
End If

End Sub

Sub f_GuardarCambios()
Dim strMsg As String

If estadoAbm = 3 Then
    
    strMsg = MsgBox("¿Esta seguro que desea guardar los cambios realizados?", vbYesNo)

    If strMsg = vbYes Then
                                      
        dbdiet.Execute "delete * from IngredientesPlatos where idPlato = " & DataCombo1.BoundText
        
        'se ejecuta la consulta que transfiere los registros de la tabla "turnos_tmp" a la tabla "turnos"
        dbdiet.Execute ("csl_tmp_a_IngredientesPlatos")
        
        'lo hago dos veces ya que el PUTOOOO VB6 no me refrezca el datagrid
        'se ejecuta la consulta que transfiere los registros de la tabla "turnos_tmp" a la tabla "turnos"
        dbdiet.Execute ("csl_tmp_a_IngredientesPlatos")
                       
        estadoAbm = 1
        
        Me.cmd_Guardar.Enabled = False
        Me.cmd_Cancelar.Enabled = False
               
    End If

Else

    strMsg = MsgBox("No se han realizado cambios", vbInformation)

End If

End Sub

Sub f_Valida_Datos()

If Adodc3.Recordset.Fields("tmp_idPlato").Value > 0 And Adodc3.Recordset.Fields("tmp_CodAlimento").Value > 0 And Adodc3.Recordset.Fields("tmp_Porcion").Value > 0 Then
        
    'actualizo el valor correspondiente del grupo alimentario segun el alimento seleccionado
        strQuery = " SELECT * FROM Alimentos, categoria WHERE CodAlimento = " & Datagrid1.Columns(4).Value & _
                    " AND Alimentos.idCategoria = Categoria.idCategoria "
        
        Set tb = dbdiet.OpenRecordset(strQuery)
        
        'Grupo Alimentario
        Datagrid1.Columns(0).Value = tb.Fields("Decripcion").Value
        'Ingrediente
        Datagrid1.Columns(1).Value = tb.Fields("DescripAlimento").Value
        
        tb.Close
    '--------------------------------------------------------------
    
End If

End Sub

Sub f_TotalKcal()
Dim KcalAux, PorcionAux, nKcal
Dim nRecordCount As Integer

nKcal = 0

nRecordCount = Adodc3.Recordset.RecordCount

If nRecordCount > 0 Then
    Adodc3.Recordset.MoveFirst
End If

For i = 1 To nRecordCount
           
    Set tb = dbdiet.OpenRecordset("alimentos", dbOpenDynaset)
    
    tb.FindFirst " codalimento = " & Adodc3.Recordset.Fields("tmp_CodAlimento").Value
    
    hc = tb.Fields("hc").Value
    prot = tb.Fields("prot").Value
    lip = tb.Fields("lip").Value
    
    tb.Close
    
    KcalAux = hc * 4 + prot * 4 + lip * 9
    
    PorcionAux = Adodc3.Recordset.Fields("tmp_Porcion").Value
    
    nKcal = nKcal + (PorcionAux * KcalAux / 100)
    
    If Adodc3.Recordset.EOF = False Then
        Adodc3.Recordset.MoveNext
    End If
    
Next

lbl_Kcal = nKcal

End Sub

Sub f_DatagridWidth()

Datagrid1.Columns(0).Width = 2145.26
Datagrid1.Columns(1).Width = 1530.142
Datagrid1.Columns(2).Width = 1530.142

End Sub

