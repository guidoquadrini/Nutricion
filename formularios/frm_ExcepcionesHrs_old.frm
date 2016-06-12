VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_ExcepcionesHrs_old 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excepciones de Horarios"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   Icon            =   "frm_ExcepcionesHrs_old.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9180
   Begin VB.Frame fme_body 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2655
      Left            =   960
      TabIndex        =   46
      Top             =   2520
      Width           =   6255
      Begin MSComCtl2.DTPicker DTP_Fecha 
         DataField       =   "ehr_fecha"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   360
         TabIndex        =   47
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   66191361
         CurrentDate     =   37867
      End
      Begin MSComCtl2.DTPicker DTP_hrHasta 
         Height          =   285
         Left            =   3120
         TabIndex        =   48
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Format          =   66191362
         CurrentDate     =   38495
      End
      Begin MSComCtl2.DTPicker DTP_hrDesde 
         Height          =   285
         Left            =   2040
         TabIndex        =   49
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Format          =   66191362
         CurrentDate     =   38495
      End
      Begin MSDataGridLib.DataGrid Datagrid1 
         Bindings        =   "frm_ExcepcionesHrs_old.frx":0ECA
         Height          =   1935
         Left            =   120
         TabIndex        =   50
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
         Caption         =   "Excepciones de horarios"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "ehr_idProf"
            Caption         =   "ehr_idProf"
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
            DataField       =   "ehr_fecha"
            Caption         =   "Fecha"
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
            DataField       =   "ehr_hrDesde"
            Caption         =   "hr. Desde"
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
            DataField       =   "ehr_hrHasta"
            Caption         =   "hr. Hasta"
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
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               Button          =   -1  'True
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               Button          =   -1  'True
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column03 
               Button          =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc ado_Datagrid 
         Height          =   375
         Left            =   0
         Top             =   2160
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
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
         Connect         =   "FILE NAME=D:\Omnia\OLEDB_Omnia.UDL"
         OLEDBString     =   ""
         OLEDBFile       =   "D:\Omnia\OLEDB_Omnia.UDL"
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from excepcionesHrs order by ehr_fecha, ehr_idprof, ehr_hrdesde"
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
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6840
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   4200
         TabIndex        =   42
         Top             =   240
         Width           =   615
         Begin VB.CommandButton cmd_Tipito 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_ExcepcionesHrs_old.frx":0EE5
            Height          =   315
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_ExcepcionesHrs_old.frx":15F5
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Agregar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.PictureBox Pic_Tipito 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            MouseIcon       =   "frm_ExcepcionesHrs_old.frx":1885
            Picture         =   "frm_ExcepcionesHrs_old.frx":19D7
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   43
            Top             =   120
            Width           =   315
         End
         Begin VB.PictureBox Pic_Tipito_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   120
            MouseIcon       =   "frm_ExcepcionesHrs_old.frx":1C67
            Picture         =   "frm_ExcepcionesHrs_old.frx":1DB9
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   45
            Top             =   120
            Width           =   315
         End
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_ExcepcionesHrs_old.frx":1EE9
         DataField       =   "ehr_idProf"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "nom"
         BoundColumn     =   "prf_codigo"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "ehr_hrDesde"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   3
         Top             =   1215
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "ehr_hrHasta"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   5040
         TabIndex        =   4
         Top             =   1215
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   3
         Height          =   15
         Left            =   120
         Top             =   840
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   7
         Top             =   1230
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1230
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   1230
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Profesional:"
         DragMode        =   1  'Automatic
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   390
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1440
      Top             =   0
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Adodc1 dataCbo"
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
   Begin VB.Frame fme_botones_abm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   1320
      TabIndex        =   8
      Top             =   5520
      Width           =   5535
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":1EFE
         Height          =   375
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":2057
         Picture         =   "frm_ExcepcionesHrs_old.frx":21A9
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Aceptar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":2465
         Height          =   375
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":25F9
         Picture         =   "frm_ExcepcionesHrs_old.frx":274B
         Style           =   1  'Graphical
         TabIndex        =   35
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":2BFE
         Picture         =   "frm_ExcepcionesHrs_old.frx":2D50
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   34
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":3051
         Picture         =   "frm_ExcepcionesHrs_old.frx":31A3
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   33
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":345F
         Height          =   375
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":3580
         Picture         =   "frm_ExcepcionesHrs_old.frx":36D2
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Modificar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":3945
         Height          =   375
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":3A5B
         Picture         =   "frm_ExcepcionesHrs_old.frx":3BAD
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Eliminar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":3D3C
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":3E89
         Picture         =   "frm_ExcepcionesHrs_old.frx":3FDB
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Agregar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUltimo 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":4415
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":45BD
         Picture         =   "frm_ExcepcionesHrs_old.frx":470F
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Ultimo"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSiguiente 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":4BDA
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":4D47
         Picture         =   "frm_ExcepcionesHrs_old.frx":4E99
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Siguiente"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":530E
         Height          =   375
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":5496
         Picture         =   "frm_ExcepcionesHrs_old.frx":55E8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Buscar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":58C5
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":5A2F
         Picture         =   "frm_ExcepcionesHrs_old.frx":5B81
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Anterior"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimero 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":5FEF
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":6194
         Picture         =   "frm_ExcepcionesHrs_old.frx":62E6
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":67A1
         Picture         =   "frm_ExcepcionesHrs_old.frx":68F3
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   16
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":6DAE
         Picture         =   "frm_ExcepcionesHrs_old.frx":6F00
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   15
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":736E
         Picture         =   "frm_ExcepcionesHrs_old.frx":74C0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   14
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":779D
         Picture         =   "frm_ExcepcionesHrs_old.frx":78EF
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   13
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":7D64
         Picture         =   "frm_ExcepcionesHrs_old.frx":7EB6
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":8381
         Picture         =   "frm_ExcepcionesHrs_old.frx":84D3
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   11
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":890D
         Picture         =   "frm_ExcepcionesHrs_old.frx":8A5F
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   10
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":8BEE
         Picture         =   "frm_ExcepcionesHrs_old.frx":8D40
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   9
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":8FB3
         Picture         =   "frm_ExcepcionesHrs_old.frx":9105
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   25
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":9226
         Picture         =   "frm_ExcepcionesHrs_old.frx":9378
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   26
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":948E
         Picture         =   "frm_ExcepcionesHrs_old.frx":95E0
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   27
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":972D
         Picture         =   "frm_ExcepcionesHrs_old.frx":987F
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   28
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":9A27
         Picture         =   "frm_ExcepcionesHrs_old.frx":9B79
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   29
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":9CE6
         Picture         =   "frm_ExcepcionesHrs_old.frx":9E38
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   30
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":9FC0
         Picture         =   "frm_ExcepcionesHrs_old.frx":A112
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   31
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":A27C
         Picture         =   "frm_ExcepcionesHrs_old.frx":A3CE
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   32
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":A573
         Picture         =   "frm_ExcepcionesHrs_old.frx":A6C5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   37
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":A81E
         Picture         =   "frm_ExcepcionesHrs_old.frx":A970
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   38
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_ExcepcionesHrs_old.frx":AB04
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_ExcepcionesHrs_old.frx":AC5C
         Style           =   1  'Graphical
         TabIndex        =   40
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":B0DC
         Picture         =   "frm_ExcepcionesHrs_old.frx":B22E
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   41
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
         MouseIcon       =   "frm_ExcepcionesHrs_old.frx":B386
         Picture         =   "frm_ExcepcionesHrs_old.frx":B4D8
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   39
         Top             =   120
         Width           =   375
      End
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
      RecordSource    =   "ExcepcionesHrs"
      Top             =   6330
      Visible         =   0   'False
      Width           =   9180
   End
End
Attribute VB_Name = "frm_ExcepcionesHrs_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
                            
Private Sub cmd_Tipito_Click()

Unload frm_abm_prof
frm_abm_prof.Show
frm_abm_prof.Data1.Recordset.FindFirst " prf_codigo = " & DataCombo1.BoundText

End Sub

Private Sub cmdAceptar_Click()

If estadoAbm = 2 Or estadoAbm = 3 Then 'si el estado es agregar o modificar
        
    If f_Valida_Update = True Then
    
        MDIForm1.ActiveForm.Data1.UpdateRecord
        MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
    
    
        'condiciones extras
            'If estadoAbm = 2 Then
            '    dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(MDIForm1.ActiveForm.Label1.Caption) & ", codalimento from alimentos where estado = true"
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
            
        Call f_Boton_Zorder
        
    Else
    
        MsgBox " No se puede actualizar el registro. " & vbCrLf & " Hay conflictos con registros ya cargados para el mismo profesional en la misma fecha. " & vbCrLf & " Verifique.", vbInformation
    
    End If
    
Else

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        Unload Me
    
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

MDIForm1.ActiveForm.DataCombo1.SetFocus

Me.DTPicker1.Value = Format(Now, "Short Date")

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

Call f_Boton_Zorder
End Sub

Private Sub cmdBuscar_Click()
Dim strQuery As String
'aclare campo por el cual buscar
    msg = InputBox("Ingrese apellido del profesional:", "Buscar por Apellido")
    
    If Len(msg) > 0 Then
        strQuery = " select * from ExcepcionesHrs, profesionales where ehr_idProf = prf_codigo and prf_apell like '" & msg & "*' order by prf_apell, prf_nombre"
    Else
        strQuery = "select * from excepcionesHrs order by ehr_fecha"
    End If
    
With MDIForm1.ActiveForm.Data1
    .RecordSource = strQuery
    .Refresh
End With

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
                 
        Unload Me
        
    End If

End If
End Sub



Private Sub cmdImprimir_Click()
Dim strQuery As String

CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_ExcepcionesHrs_one.rpt"

'aclare el filtro para imprimir
msg = MsgBox("¿Desea imprimir todos los registros?", vbYesNo, "Imprimir")
  
If msg = vbYes Then
    
    strQuery = " {ExcepcionesHrs.ehr_idprof} = " & DataCombo1.BoundText
    
Else
    
    strQuery = " {ExcepcionesHrs.ehr_idprof} = " & DataCombo1.BoundText & " and {ExcepcionesHrs.ehr_fecha} = Date (" & Year(DTPicker1.Value) & ", " & Month(DTPicker1.Value) & ", " & Day(DTPicker1.Value) & ")"
    
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
MDIForm1.ActiveForm.DataCombo1.SetFocus

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

Private Sub DataCombo1_Validate(Cancel As Boolean)

If DataCombo1.Text = "" Then
    MsgBox "Debe seleccionar un profesional", vbInformation, "Información"
    Cancel = True
Else
    Cancel = False
End If

End Sub

Private Sub Datagrid1_ButtonClick(ByVal ColIndex As Integer)

Select Case ColIndex
    Case Is = 1
    
            With Me.DTP_Fecha
                
                .Left = Datagrid1.Columns(ColIndex).Left '+ 100
                .Width = Datagrid1.Columns(ColIndex).Width '+ 50
                        
                .Top = Datagrid1.Top + Datagrid1.RowTop(Datagrid1.Row) + Datagrid1.RowHeight
                
                'en caso de error continuo con la siguiente intruccion
                'ya que cuando estoy agregando un registro la siguiente
                'intruccion provoca un error
                ''On Error Resume Next
                ''.BoundText = f_Devuelve_idCategoria(Datagrid1.Columns(4).Value) ' devuelve el idCategoria para el codAlimento dado
                
                .Visible = True
                .SetFocus
            End With
    
    Case Is = 2
            
            With Me.DTP_hrDesde
            
                .Left = Datagrid1.Columns(ColIndex).Left '+ 100
                .Width = Datagrid1.Columns(ColIndex).Width '+ 50
            
                .Top = Datagrid1.Top + Datagrid1.RowTop(Datagrid1.Row) + Datagrid1.RowHeight
                                                
                'en caso de error continuo con la siguiente intruccion
                'ya que cuando estoy agregando un registro la siguiente
                'intruccion provoca un error
                ''On Error Resume Next
                ''.BoundText = Datagrid1.Columns(4).Value 'CodAlimento
                                
                .Visible = True
                .SetFocus
            End With
    
    Case Is = 3
            
            With Me.DTP_hrHasta
            
                .Left = Datagrid1.Columns(ColIndex).Left '+ 100
                .Width = Datagrid1.Columns(ColIndex).Width '+ 50
            
                .Top = Datagrid1.Top + Datagrid1.RowTop(Datagrid1.Row) + Datagrid1.RowHeight
                                                
                'en caso de error continuo con la siguiente intruccion
                'ya que cuando estoy agregando un registro la siguiente
                'intruccion provoca un error
                ''On Error Resume Next
                ''.BoundText = Datagrid1.Columns(4).Value 'CodAlimento
                                
                .Visible = True
                .SetFocus
            End With

End Select
End Sub

Private Sub Form_Activate()
'''para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
''Me.Height = 2970
''Me.Width = 6150
''Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
''Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Call enabledDesplaz 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados

Call f_Boton_Zorder

End Sub

Private Sub Form_Load()
'Data1.DatabaseName = Lugar

estadoAbm = 1 ' el estado es sim cambios

Call f_CargarOrigenDatos

'-------------------------
'se refresca el data1 para que el metodo enabledDesplaz funcione correctamente con el recordset cargado
''strquery = " select * from excepcionesHrs order by ehr_fecha"
''
''With Data1
''    .RecordSource = strquery
''    .Refresh
''End With
'--------------------------------
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdCancelar_Click

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub MaskEdBox1_Change(Index As Integer)

MaskEdBox1(Index).SelStart = 0
MaskEdBox1(Index).SelLength = 5

End Sub

Private Sub MaskEdBox1_GotFocus(Index As Integer)

MaskEdBox1(Index).SelStart = 0
MaskEdBox1(Index).SelLength = 5

End Sub

Private Sub MaskEdBox1_Validate(Index As Integer, Cancel As Boolean)

If hora(Val(MaskEdBox1(Index).Text)) = True Then
    Cancel = False
Else
    Cancel = True
    MsgBox "Debe ingresar un hora valida", vbInformation, "Información"
    MaskEdBox1(Index).SetFocus
    MaskEdBox1(Index).SelStart = 0
    MaskEdBox1(Index).SelLength = 5
End If

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Data1.Recordset = Nothing
Set Me.Adodc1.Recordset = Nothing

strQuery = "select * from excepcionesHrs order by ehr_fecha, ehr_idprof, ehr_hrdesde"
Call f_Data_DatabaseName(Data1, strQuery)

strQuery = "select *, (prf_apell & ', ' & prf_nombre) as nom from profesionales order by prf_apell, prf_nombre"
Call f_Adodc_ConnectionString(Adodc1, strQuery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo1, Data1, Adodc1, "ehr_idProf", "prf_codigo", "nom")

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


Private Sub Pic_Imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Imprimir

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

Private Sub Pic_Tipito_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Tipito

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
Function f_Valida_Update() As Boolean
'valida que el horario que se ingreso sea valido, es decir, que en el caso de que ya se haya ingresado
'un registro para el profesional y fecha seleccionada sea un horario valido
'devuelve faso en el caso de que no se deba permitir agregar el registro, verdadero en caso contrario
Dim strQuery, ehr_hrDesde, ehr_hrHasta As String

f_Valida_Update = False

strQuery = " select * from excepcionesHrs where ehr_idProf = " & DataCombo1.BoundText & " and ehr_fecha = #" & DTPicker1.Value & "#"

Set tb = dbdiet.OpenRecordset(strQuery)

'ya que solo se va a permitir tener dos turnos solo se pueden encontrar a lo sumo dos registros
'por lo que solo la validacion sera para el caso de que haya un solo registro
'si no hay ninguno se permite el update
'si ya hay dos se niega el update
If f_Cant_Registros(tb) = 1 Then
       
    tb.MoveFirst
    
    ehr_hrDesde = tb.Fields("ehr_hrDesde").Value
    ehr_hrHasta = tb.Fields("ehr_hrHasta").Value
    
    If Me.MaskEdBox1(0).Text > ehr_hrDesde And Me.MaskEdBox1(0).Text < ehr_hrHasta Then
    
    End If
    
    If Me.MaskEdBox1(0).Text < ehr_hrDesde Then
        
        If Me.MaskEdBox1(1).Text < ehr_hrDesde And Me.MaskEdBox1(1).Text > Me.MaskEdBox1(0).Text Then
        
            f_Valida_Update = True
        
        End If
        
    Else
        
        If Me.MaskEdBox1(0).Text > ehr_hrHasta Then
        
            If Me.MaskEdBox1(1).Text > Me.MaskEdBox1(0).Text Then
            
                f_Valida_Update = True
                
            End If
        
        End If
        
    End If
    
    
Else
    
    If tb.RecordCount = 0 Then
        
        f_Valida_Update = True
        
    End If
    
End If

tb.Close

End Function


