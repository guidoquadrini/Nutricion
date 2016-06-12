VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frm_agendaTurnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda de Turnos"
   ClientHeight    =   6765
   ClientLeft      =   960
   ClientTop       =   1185
   ClientWidth     =   6390
   Icon            =   "frm_agendaTurnos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6390
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frm_agendaTurnos.frx":0ECA
      DataField       =   "tmp_idPaciente"
      Height          =   1815
      Left            =   1680
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      _Version        =   393216
      ListField       =   "nom"
      BoundColumn     =   "Legajo"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   5040
      Top             =   0
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
      Caption         =   "Adodc4 List"
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
      Height          =   6735
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6375
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   495
         Left            =   2048
         TabIndex        =   26
         Top             =   6135
         Width           =   2295
         Begin VB.CommandButton cmdAceptar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_agendaTurnos.frx":0EDF
            Enabled         =   0   'False
            Height          =   375
            Left            =   720
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_agendaTurnos.frx":1038
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Guardar cambios"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmd_salir 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_agendaTurnos.frx":12F4
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_agendaTurnos.frx":1488
            Style           =   1  'Graphical
            TabIndex        =   38
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
            MouseIcon       =   "frm_agendaTurnos.frx":1789
            Picture         =   "frm_agendaTurnos.frx":18DB
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   37
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox PicAceptar 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   720
            MouseIcon       =   "frm_agendaTurnos.frx":1BDC
            Picture         =   "frm_agendaTurnos.frx":1D2E
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   36
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
            MouseIcon       =   "frm_agendaTurnos.frx":1FEA
            Picture         =   "frm_agendaTurnos.frx":213C
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   35
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox PicAceptar_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   720
            MouseIcon       =   "frm_agendaTurnos.frx":22D0
            Picture         =   "frm_agendaTurnos.frx":2422
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   34
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmd_deshacer 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_agendaTurnos.frx":257B
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_agendaTurnos.frx":26F3
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Deshacer cambios"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.PictureBox Pic_Deshacer 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1200
            MouseIcon       =   "frm_agendaTurnos.frx":2B73
            Picture         =   "frm_agendaTurnos.frx":2CC5
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   32
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Pic_Deshacer_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1200
            MouseIcon       =   "frm_agendaTurnos.frx":3145
            Picture         =   "frm_agendaTurnos.frx":3297
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   31
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmd_cancel 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_agendaTurnos.frx":340F
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_agendaTurnos.frx":35A3
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Salir"
            Top             =   120
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Pic_cancel 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1200
            MouseIcon       =   "frm_agendaTurnos.frx":38A4
            Picture         =   "frm_agendaTurnos.frx":39F6
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   29
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Pic_cancel_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1200
            MouseIcon       =   "frm_agendaTurnos.frx":3CF7
            Picture         =   "frm_agendaTurnos.frx":3E49
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   28
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImprimir 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_agendaTurnos.frx":3FDD
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_agendaTurnos.frx":4135
            Style           =   1  'Graphical
            TabIndex        =   41
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
            MouseIcon       =   "frm_agendaTurnos.frx":45B5
            Picture         =   "frm_agendaTurnos.frx":4707
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   40
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
            MouseIcon       =   "frm_agendaTurnos.frx":4B87
            Picture         =   "frm_agendaTurnos.frx":4CD9
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   27
            Top             =   120
            Width           =   375
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_agendaTurnos.frx":4E31
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "tmp_idProf"
            Caption         =   "Profesional"
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
            DataField       =   "tmp_fecha"
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
            DataField       =   "tmp_hrDesde"
            Caption         =   "Horario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0#:##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "tmp_duracion"
            Caption         =   "Duracion Turno"
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
            DataField       =   "tmp_idPaciente"
            Caption         =   "Hist. Clinica"
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
            DataField       =   "tmp_Nombre"
            Caption         =   "Nombre"
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
            DataField       =   "tmp_observacion"
            Caption         =   "Observacion"
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
            DataField       =   "tmp_hrHasta"
            Caption         =   "tmp_hrHasta"
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
            DataField       =   "tmp_estado"
            Caption         =   "Estado"
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
               Object.Visible         =   0   'False
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   5280
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.TextBox txt_legajo 
         Height          =   285
         Left            =   5520
         TabIndex        =   16
         Top             =   4440
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   7080
         Top             =   5400
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Adodc1 grilla"
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
         DatabaseName    =   "D:\Dietetica\Database\db1nueva prueba anterior sin replica.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   7320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Profesionales"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos del Paciente:"
         Enabled         =   0   'False
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   4560
         Width           =   5895
         Begin VB.ComboBox cbo_estado 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frm_agendaTurnos.frx":4E46
            Left            =   1560
            List            =   "frm_agendaTurnos.frx":4E53
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox txt_nombrePaciente 
            Height          =   285
            Left            =   1560
            TabIndex        =   2
            Top             =   240
            Width           =   3735
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_agendaTurnos.frx":4E79
            Enabled         =   0   'False
            Height          =   315
            Left            =   5400
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_agendaTurnos.frx":5589
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Agregar"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox txt_observacion 
            Height          =   495
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   960
            Width           =   4215
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   7320
            Top             =   840
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
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   330
            Left            =   7320
            Top             =   480
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
         Begin VB.Label lblLabels 
            Caption         =   "Estado:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   630
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Apellido y Nombre:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Observaciones:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1335
         End
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   120
         Top             =   6240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "D:\Dietetica\rpts\rep_pacientes.rpt"
         PrintFileLinesPerPage=   60
      End
      Begin MSDataGridLib.DataGrid DataGrid1_back 
         Bindings        =   "frm_agendaTurnos.frx":5C8D
         Height          =   2895
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "tmp_idProf"
            Caption         =   "Profesional"
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
            DataField       =   "tmp_fecha"
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
            DataField       =   "tmp_hrDesde"
            Caption         =   "Horario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0#:##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "tmp_duracion"
            Caption         =   "Duracion Turno"
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
            DataField       =   "tmp_idPaciente"
            Caption         =   "Hist. Clinica"
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
            DataField       =   "tmp_Nombre"
            Caption         =   "Nombre"
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
            DataField       =   "tmp_observacion"
            Caption         =   "Observacion"
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
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame8 
         Caption         =   "Seleccionar profesional:"
         Height          =   1095
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   6135
         Begin MSComCtl2.DTPicker DTPicker1 
            Bindings        =   "frm_agendaTurnos.frx":5CA2
            Height          =   315
            Left            =   1560
            TabIndex        =   0
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Format          =   65798145
            CurrentDate     =   38173
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "nom"
            BoundColumn     =   ""
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
            Left            =   3720
            TabIndex        =   19
            Top             =   480
            Width           =   1695
            Begin VB.CommandButton cmd_Cancelar 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_agendaTurnos.frx":5CAD
               Height          =   375
               Left            =   1200
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frm_agendaTurnos.frx":5E41
               Picture         =   "frm_agendaTurnos.frx":5F93
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Cancelar"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.CommandButton cmd_Aceptar 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_agendaTurnos.frx":6446
               Height          =   375
               Left            =   720
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frm_agendaTurnos.frx":659F
               Picture         =   "frm_agendaTurnos.frx":66F1
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Aceptar"
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
               Left            =   1200
               MouseIcon       =   "frm_agendaTurnos.frx":69AD
               Picture         =   "frm_agendaTurnos.frx":6AFF
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   20
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Aceptar 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   720
               MouseIcon       =   "frm_agendaTurnos.frx":6E00
               Picture         =   "frm_agendaTurnos.frx":6F52
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
               MouseIcon       =   "frm_agendaTurnos.frx":720E
               Picture         =   "frm_agendaTurnos.frx":7360
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   25
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Aceptar_Gris 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   720
               MouseIcon       =   "frm_agendaTurnos.frx":74F4
               Picture         =   "frm_agendaTurnos.frx":7646
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   24
               Top             =   120
               Width           =   375
            End
            Begin VB.CommandButton cmd_Calendario 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_agendaTurnos.frx":779F
               Height          =   375
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frm_agendaTurnos.frx":7C1F
               Picture         =   "frm_agendaTurnos.frx":7D71
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Calendario"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.PictureBox Pic_Calendario_Gris 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               MouseIcon       =   "frm_agendaTurnos.frx":8232
               Picture         =   "frm_agendaTurnos.frx":8384
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   43
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Calendario 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               MouseIcon       =   "frm_agendaTurnos.frx":8514
               Picture         =   "frm_agendaTurnos.frx":8666
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   42
               Top             =   120
               Width           =   375
            End
         End
         Begin VB.Label lbl_matricula 
            AutoSize        =   -1  'True
            Caption         =   "Label1"
            DataField       =   "prf_matricula"
            DataSource      =   "Data1"
            Height          =   195
            Left            =   4440
            TabIndex        =   46
            Top             =   285
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Matricula:"
            Height          =   195
            Left            =   3705
            TabIndex        =   45
            Top             =   285
            Width           =   690
         End
         Begin VB.Label lblLabels 
            Caption         =   "Profesional:"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   15
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Fecha:"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   14
            Top             =   630
            Width           =   1815
         End
      End
      Begin VB.Shape Shape2 
         Height          =   45
         Left            =   120
         Top             =   1320
         Width           =   6135
      End
      Begin VB.Shape Shape1 
         Height          =   45
         Left            =   120
         Top             =   4440
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frm_agendaTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
Dim tb1 As Recordset

Dim tb_ExcepcionHrs As Recordset

Dim HrDesde, HrHasta, HrDesde2, HrHasta2 As String
Dim TiempoTur As Integer
' columnas del datagrid
' tmp_idProf ===> Profesional
' tmp_fecha ===> Fecha
' tmp_hsDesde ===> Horario
' tmp_duracion ===> Duracion Turno
' tmp_idPaciente ===> ID Paciente
' tmp_nombre ===> Nombre
' tmp_observacion ===> Obserbacion
'========================
'Cbo_estado:
'       - Ninguno -------> listindex = -1
'       - Pendiente -------> listindex = 0
'       - Concretado ------> listindex = 1
'       - Cancelado -------> listindex = 2
'========================
Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar


Private Sub cmd_aceptar_Click()
Dim sMsg As String

sMsg = vbNo

If f_ValidaCarga = True Then
    
    If f_isFeriado(DTPicker1.Value) = True Then
        
        sMsg = MsgBox("La fecha seleccionada es feriado..." & vbCrLf & vbTab & " - ¿esta seguro que desea asignar turnos?", vbYesNo)
           
    End If
    
    If f_isFeriado(DTPicker1.Value) = False Or sMsg = vbYes Then
        
        Call newTurnos

        Frame8.Enabled = False
        cmd_aceptar.Enabled = False
        
        cmdImprimir.Enabled = True
        cmd_deshacer.Enabled = True
        cmdAceptar.Enabled = True
        cmd_Cancelar.Enabled = False
        
        cmd_Calendario.Enabled = False
        
        DataGrid1.Enabled = True
        DataGrid1.BorderStyle = dbgFixedSingle
        DataGrid1.Appearance = dbg3D
        DataGrid1.SetFocus
        DataGrid1.BackColor = &H80000005
        
        cmd_salir.Enabled = True
        
        'Adodc1.Refresh
        Call DatagridRefresh(Adodc1, DataGrid1)
    
        Call f_Boton_Zorder
        
    End If
End If

End Sub

Private Sub cmd_cancel_Click()

DataGrid1.Enabled = True
DataGrid1.BorderStyle = dbgFixedSingle
DataGrid1.Appearance = dbg3D
DataGrid1.SetFocus
DataGrid1.BackColor = &H80000005

'cmd_aceptar.Enabled = True
'cmd_cancelar.Enabled = True

cmd_cancel.Visible = False
Me.Pic_cancel.Visible = False
Me.Pic_cancel_Gris.Visible = False
cmd_cancel.Enabled = False
'Me.Pic_cancel_Gris.ZOrder 0

cmd_deshacer.Visible = True
Me.Pic_Deshacer.Visible = True
Me.Pic_Deshacer_Gris.Visible = True
cmd_deshacer.Enabled = True
Me.Pic_Deshacer.ZOrder 0

cmd_salir.Enabled = True

Frame3.Enabled = False
'Frame3.Visible = False

'cmd_Calendario.Enabled = True
Command3.Enabled = False
cbo_estado.Enabled = False

txt_legajo.Text = ""
txt_nombrePaciente.Text = ""
txt_observacion.Text = ""
cbo_estado.ListIndex = -1

Call f_Boton_Zorder

End Sub

Private Sub cmd_Cancelar_Click()

Unload Me
End Sub

Private Sub cmd_deshacer_Click()
Dim sMsg As String

If estadoAbm = 3 Then
    
    sMsg = MsgBox("¿Esta seguro que desea deshacer los cambios?", vbYesNo)
    
    If sMsg = vbYes Then
    
        Call DeleteContTable("turnos_tmp")
        newTurnos
        Call DatagridRefresh(Adodc1, DataGrid1)
        
        estadoAbm = 1
        cmdAceptar.ToolTipText = "Aceptar"
                
        Call f_Boton_Zorder
        
    Else
    End If

Else
    MsgBox "No se han realizado cambios", vbInformation

End If

  
End Sub

Private Sub cmd_salir_Click()
Dim strMsg As String
Dim strQuery As String

If estadoAbm = 3 Then
    strMsg = MsgBox("¿Esta seguro que desea finalizar la operacion?" & vbCrLf & vbTab & "- Se perderan los cambios realizados", vbYesNo)
Else
    strMsg = MsgBox("¿Esta seguro que desea finalizar la operacion?", vbYesNo)
End If

If strMsg = vbYes Then

    Call f_Cancela
        
    Call f_Boton_Zorder
    
End If

End Sub

Private Sub cmdAceptar_Click()
Dim nIdPaciente 'As Long
Dim sNombre 'As String
Dim sObservacion 'As String
Dim sEstado

'sNombre = ""

'nIdPaciente = 0

'sObservacion = ""

MousePointer = vbHourglass

If DataGrid1.Enabled = False Then 'significa que se esta agregando un registro
    
    If txt_nombrePaciente.Text <> "" Then
    
        DataGrid1.Enabled = True
        Frame3.Enabled = True
        
        DataGrid1.BorderStyle = dbgNoBorder
        DataGrid1.Appearance = dbgFlat
        DataGrid1.BackColor = &H80000005
              
        sNombre = txt_nombrePaciente.Text
        nIdPaciente = Val(txt_legajo.Text)
        sObservacion = txt_observacion.Text
        
        If cbo_estado.Text = "" Then
            cbo_estado.ListIndex = 0
        End If
        
        sEstado = cbo_estado.Text
        
    '    If Len(sNombre) = 0 Then
    '        sNombre = " "
    '    End If
    '
    '    If Len(nIdPaciente) = 0 Then
    '        nIdPaciente = " "
    '    End If
        
        If Len(sObservacion) = 0 Then
            sObservacion = " "
        End If
        
        'solo actualizo el campo nombre ya que el resto estan
        'previamente cargados cuando se ejecuto el procedimiento "calculaTurnos"
        Adodc1.Recordset.Update "tmp_Nombre", sNombre
        Adodc1.Recordset.Update "tmp_idPaciente", nIdPaciente
        Adodc1.Recordset.Update "tmp_observacion", sObservacion
        Adodc1.Recordset.Update "tmp_estado", sEstado
               
        estadoAbm = 3
        cmdAceptar.ToolTipText = "Guardar cambios"
        
        'Frame3.Visible = False
        Frame3.Enabled = False
        Command3.Enabled = False
        
        cmd_cancel.Visible = False
        Me.Pic_cancel.Visible = False
        Me.Pic_cancel_Gris.Visible = False
        cmd_cancel.Enabled = False
        'Me.Pic_cancel_Gris.ZOrder 0
        
        cmd_deshacer.Visible = True
        Me.Pic_Deshacer.Visible = True
        Me.Pic_Deshacer_Gris.Visible = True
        cmd_deshacer.Enabled = True
        Me.Pic_Deshacer.ZOrder 0
        
        cmd_salir.Enabled = True
        
        'cmdAceptar.Enabled = False
        DataGrid1.SetFocus
        
        MousePointer = vbDefault
        
    Else
        MousePointer = vbDefault
        
        MsgBox "El campo apellido y nombre no puede quedar en blanco", vbInformation
                
    End If
    
Else

    If estadoAbm = 3 Then
    
        BeginTrans
        
        'elimino los datos del profesional en la fecha actual de la tabla turnos para luego actualizarlos con los nuevos
        strQuery = "Delete * from turnos where tur_idProf = " & DataCombo1.BoundText & " and tur_fecha = #" & Format(DTPicker1.Value, "mm/dd/yy") & "#"
        dbdiet.Execute strQuery
        
        'se ejecuta la consulta que transfiere los registros de la tabla "turnos_tmp" a la tabla "turnos"
        dbdiet.Execute ("csl_tmp_a_tur")
        
        msg = MsgBox("¿Desea guardar los cambios?", vbYesNo, "Guardar")
    
        If msg = vbYes Then
    
            CommitTrans
              
            estadoAbm = 1
            cmdAceptar.ToolTipText = "Aceptar"
            
        Else
    
            Rollback
    
        End If
            
        MousePointer = vbDefault
    
    Else
        
        MousePointer = vbDefault
        
        MsgBox "No se han realizado cambios", vbInformation
    
    End If

End If

Call f_Boton_Zorder

End Sub



Private Sub cmdImprimir_Click()
Dim strQuery As String

CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_turnos_one.rpt"

strQuery = " {turnos.tur_idprof} = " & DataCombo1.BoundText & " and {turnos.tur_fecha} = Date (" & Year(DTPicker1.Value) & ", " & Month(DTPicker1.Value) & ", " & Day(DTPicker1.Value) & ")"

Call f_print(CrystalReport1, strQuery, crptToWindow)


End Sub

Private Sub Command1_Click()

strQuery = "select * from turnos_tmp order by tmp_hrdesde"
With Adodc1
    .RecordSource = strQuery
    .Refresh
End With


Call DatagridRefresh(Adodc1, DataGrid1)


End Sub

Private Sub cmd_Calendario_Click()
Dim strQuery As String

strQuery = "DblClick para recuperar los registros de la fecha"

frm_calendario.cargarParametros 2, DataCombo1.BoundText, strQuery

frm_calendario.Show

'Me.Hide

End Sub

Private Sub Command3_Click()
'If IsNumeric(DataCombo2.BoundText) Then
'    MsgBox "Number"
'    MsgBox DataCombo2.BoundText
'    MsgBox DataCombo2.Text
'
'Else
'    MsgBox "character"
'    MsgBox DataCombo2.BoundText
'    MsgBox DataCombo2.Text
'
'End If

DataList1.Visible = True
DataList1.SetFocus

End Sub


Private Sub DataCombo1_Click(Area As Integer)
Dim strQuery As String

If DataCombo1.Text <> "" Then
    strQuery = "select * from profesionales where prf_codigo = " & DataCombo1.BoundText
    
    With Data1
        .RecordSource = strQuery
        .Refresh
    End With
    
    lbl_matricula.Refresh

End If

End Sub


Private Sub DataGrid1_DblClick()

aux = DataGrid1.Bookmark
DataGrid1.FirstRow = aux
DataGrid1.Enabled = False
'DataGrid1.BorderStyle = dbgNoBorder
'DataGrid1.Appearance = dbgFlat
'DataGrid1.BackColor = &H8000000F

'Frame3.Visible = True
Frame3.Enabled = True
Command3.Enabled = True

cmd_cancel.Visible = True
Me.Pic_cancel.Visible = True
Me.Pic_cancel_Gris.Visible = True
cmd_cancel.Enabled = True
Me.Pic_cancel.ZOrder 0

cmd_deshacer.Visible = False
Me.Pic_Deshacer.Visible = False
Me.Pic_Deshacer_Gris.Visible = False
cmd_deshacer.Enabled = False
'Me.Pic_Deshacer_Gris.ZOrder 0

cmd_salir.Enabled = False

cbo_estado.Enabled = True

If cbo_estado.Text = "" Then
    cbo_estado.ListIndex = 0
End If

Call f_Boton_Zorder

End Sub

Private Sub DataGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    
    PopupMenu MDIForm1.popupTurnos, vbPopupMenuLeftAlign
    
End If

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim nIdPaciente ' As Integer
Dim sNombre 'As String
Dim sObservacion
Dim sEstado

nIdPaciente = Adodc1.Recordset.Fields("tmp_idpaciente").Value
sNombre = Adodc1.Recordset.Fields("tmp_nombre").Value
sObservacion = Adodc1.Recordset.Fields("tmp_observacion").Value
sEstado = Adodc1.Recordset.Fields("tmp_estado").Value

If Not IsNull(nIdPaciente) Then

    txt_legajo.Text = nIdPaciente
    
    If Not IsNull(sNombre) Then
        txt_nombrePaciente.Text = sNombre
    Else
        txt_nombrePaciente.Text = ""
    End If
    
    If Not IsNull(sObservacion) Then
        txt_observacion.Text = sObservacion
    Else
        txt_observacion.Text = ""
    End If
    
    Select Case sEstado
        Case Is = "Pendiente"
            cbo_estado.ListIndex = 0
            
        Case Is = "Concretado"
            cbo_estado.ListIndex = 1
            
        Case Is = "Cancelado"
            cbo_estado.ListIndex = 2
    
    End Select
    'Debug.Print Adodc1.Recordset.Fields("tmp_idpaciente").Value
Else

    txt_legajo.Text = ""
    txt_nombrePaciente.Text = ""
    txt_observacion.Text = ""
    cbo_estado.ListIndex = -1
    
    'If Not IsNull(sNombre) Then

        'txt_nombrePaciente.Text = sNombre

'    Else
'
'        DataCombo2.BoundText = ""
'        DataCombo2.Text = ""

    'End If

End If

End Sub

Private Sub DataList1_DblClick()

txt_nombrePaciente.Text = DataList1.Text
txt_legajo.Text = DataList1.BoundText

DataList1.Visible = False

End Sub

Private Sub DataList1_LostFocus()

DataList1.Visible = False

End Sub

Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 7140 '6765
Me.Width = 6480
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

End Sub

Private Sub Form_Load()
'Dim Lugar As String
'Lugar = App.Path & "\database\db1nueva prueba anterior sin replica.mdb"
'Set dbdiet = OpenDatabase(Lugar, False, False, ";pwd=208018")

estadoAbm = 1

Call f_CargarOrigenDatos
'Data1.DatabaseName = Lugar
'
'Adodc4.RecordSource = "select *, (prf_apell & ', ' & prf_nombre) as nom from profesionales order by prf_apell, prf_nombre"
'
'Adodc4.Refresh
Call DatagridWidth

DataGrid1.MarqueeStyle = dbgHighlightRow

Call f_Boton_Zorder

End Sub

Private Sub turnos()
Dim srtQuery As String
Dim hrs_hrDesde, hrs_hrHasta, hrs_hrTpoTur As String

strQuery = " select * from horarios where hrs_idprof = " & DataCombo1.BoundText & " and hrs_dia = " & Weekday(DTPicker1.Value, vbMonday) - 1
Set tb = dbdiet.OpenRecordset(strQuery)

hrs_hrDesde = tb.Fields("hrs_hrdesde").Value
hrs_hrHasta = tb.Fields("hrs_hrhasta").Value
hrs_TpoTur = tb.Fields("hrs_tpotur").Value

tb.Close

Call calculaTurnos(Val(hrs_hrDesde), Val(hrs_hrHasta), Val(hrs_TpoTur))

End Sub


Private Sub calculaTurnos(ByVal hrs_hrDesde As String, ByVal hrs_hrHasta As String, hrs_TpoTur As Integer)
Dim Mm As Integer
Dim hr_Desde_aux, hr_Hasta_aux As String

hr_Desde_aux = Val(hrs_hrDesde)
hr_Hasta_aux = Val(hrs_hrDesde) + hrs_TpoTur

While Val(hr_Desde_aux) < Val(hrs_hrHasta)
        
    'en el caso que hr_Desde_aux(que es un string)
    'no se corresponda con una hora valida se la fuerza atraves de calculos
    If hora(Val(hr_Hasta_aux)) = False Then
        hr_Hasta_aux = (Val(hr_Hasta_aux) - 60) + 100
    End If
          
    'se le concatena un cero delante en el string de hora en el caso
    'de que que se encuentre entre las 00:00 y las 10:00 ya que el valor
    'numerico elimina el primer cero y no lo guarda en la DB por lo que al ordenarlo por hrDesde quedaba mal
    If Val(hr_Hasta_aux) > 0 And Val(hr_Hasta_aux) < 1000 Then
        hr_Hasta_aux = "0" & hr_Hasta_aux
    End If
    
    If Val(hr_Desde_aux) > 0 And Val(hr_Desde_aux) < 1000 Then
        hr_Desde_aux = "0" & hr_Desde_aux
    End If
    ''''''''''''''''''''''''''''''''''''''''''
    
    dbdiet.Execute "insert into turnos_tmp (tmp_idprof, tmp_fecha, tmp_hrdesde, tmp_hrhasta, tmp_duracion) " & _
                            " values (" & DataCombo1.BoundText & ", '" & Format(DTPicker1.Value, "dd/mm/yy") & "', '" & hr_Desde_aux & "', '" & hr_Hasta_aux & "', " & hrs_TpoTur & ")"
    
    hr_Desde_aux = Val(hr_Desde_aux) + hrs_TpoTur
    hr_Hasta_aux = Val(hr_Hasta_aux) + hrs_TpoTur
       
    'en el caso que hr_Desde_aux(que es un string)
    'no se corresponda con una hora valida se la fuerza atraves de calculos
    If hora(Val(hr_Desde_aux)) = False Then
        hr_Desde_aux = (Val(hr_Desde_aux) - 60) + 100
    End If
    
Wend

End Sub

Sub DatagridWidth()

DataGrid1.Columns("Profesional").Width = 1005.165
DataGrid1.Columns("Fecha").Width = 1005.165
DataGrid1.Columns("Horario").Width = 615.1182
DataGrid1.Columns("Duracion Turno").Width = 1005.165
DataGrid1.Columns("Hist. Clinica").Width = 1005.165
DataGrid1.Columns("Nombre").Width = 2294.929
DataGrid1.Columns("Observacion").Width = 4995.213
DataGrid1.Columns("estado").Width = 1514.835

End Sub

Sub newTurnos()
Dim srtQuery As String
Dim hrs_hrDesde, hrs_hrHasta, hrs_hrDesde2, hrs_hrHasta2 As String
Dim hrs_TpoTur As Integer
Dim tur_hrDesde As String
Dim tur_Duracion As Integer
Dim tur_hrHasta As String
Dim tur_idPaciente As Integer
Dim tur_Nombre, tur_Observacion 'As String

hrs_hrDesde = ""
hrs_hrHasta = ""
hrs_hrDesde2 = ""
hrs_hrHasta2 = ""

'obtengo la hora de entrada, salida y duracion del turno de un determinado
'profesional en el dia seleccionado
    
    strQuery = " select * from horarios where hrs_idprof = " & DataCombo1.BoundText & " and hrs_dia = " & Weekday(DTPicker1.Value, vbMonday) - 1
    Set tb1 = dbdiet.OpenRecordset(strQuery)
    
    'aqui se valida si tiene una excepcion de horario para los turnos
    'se obtienen los valores de hora correctos segun se trate de un horario normal o de una excepcion
    If f_Tiene_ExcepcionHrs(DataCombo1.BoundText, DTPicker1.Value) = False Then
    
        'primer turno
        HrDesde = tb1.Fields("hrs_hrdesde").Value
        HrHasta = tb1.Fields("hrs_hrhasta").Value
        
        'segundo turno
        If f_getTurnos = 2 Then
            
            HrDesde2 = tb1.Fields("hrs_hrdesde2").Value
            HrHasta2 = tb1.Fields("hrs_hrhasta2").Value
        
            hrs_hrDesde2 = HrDesde2
            hrs_hrHasta2 = HrHasta2
            
        End If
    
    Else

        strQuery = " select * from excepcionesHrs where ehr_idProf = " & DataCombo1.BoundText & " and ehr_fecha = #" & DTPicker1.Value & "#"
        
        Set tb_ExcepcionHrs = dbdiet.OpenRecordset(strQuery)
        
        If f_Cant_Registros(tb_ExcepcionHrs) = 1 Then
            
            tb_ExcepcionHrs.MoveFirst
            
            HrDesde = tb_ExcepcionHrs.Fields("ehr_hrdesde").Value
            HrHasta = tb_ExcepcionHrs.Fields("ehr_hrhasta").Value
        
        Else
        
            If f_Cant_Registros(tb_ExcepcionHrs) = 2 Then
                
                tb_ExcepcionHrs.MoveFirst
                
                HrDesde = tb_ExcepcionHrs.Fields("ehr_hrdesde").Value
                HrHasta = tb_ExcepcionHrs.Fields("ehr_hrhasta").Value
                
                tb_ExcepcionHrs.MoveNext
                
                HrDesde2 = tb_ExcepcionHrs.Fields("ehr_hrdesde").Value
                HrHasta2 = tb_ExcepcionHrs.Fields("ehr_hrhasta").Value
                
                hrs_hrDesde2 = HrDesde2
                hrs_hrHasta2 = HrHasta2
                
            End If
            
        End If
           
        tb_ExcepcionHrs.Close
        
    End If
      
    TiempoTur = tb1.Fields("hrs_tpotur").Value
    
    hrs_hrDesde = HrDesde
    hrs_hrHasta = HrHasta
    
    hrs_TpoTur = TiempoTur

'obtengo los turnos asignados para un profesional en la fecha dada
    strQuery = " select * from turnos where tur_idprof = " & DataCombo1.BoundText & " and tur_fecha = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#" & " order by tur_hrdesde"
    Set tb = dbdiet.OpenRecordset(strQuery)
    
    If tb.RecordCount <> 0 Then
        
        dbdiet.Execute "insert into turnos_tmp (tmp_idprof, tmp_fecha, tmp_hrdesde, tmp_duracion, tmp_idpaciente, tmp_nombre, tmp_observacion, tmp_hrhasta, tmp_estado) select tur_idprof, tur_fecha, tur_hrdesde, tur_duracion, tur_idpaciente, tur_nombre, tur_observacion, tur_hrhasta, tur_estado from turnos where tur_idprof = " & DataCombo1.BoundText & " and tur_fecha = #" & Format(DTPicker1.Value, "mm/dd/yy") & "#"
        
        tb.MoveFirst
        
        'por cada turno obtengo horario de inicio, fin y duracion
        For i = 1 To tb.RecordCount
            
            tur_hrDesde = tb.Fields("tur_hrdesde").Value
            tur_Duracion = tb.Fields("tur_duracion").Value
            tur_hrHasta = Val(tur_hrDesde) + tur_Duracion
            tur_idPaciente = tb.Fields("tur_idpaciente").Value
            tur_Nombre = tb.Fields("tur_nombre").Value
            tur_Observacion = tb.Fields("tur_observacion").Value
            
            Select Case Val(tur_hrDesde)
                Case Val(hrs_hrDesde) To Val(hrs_hrHasta) 'primer turno
                
                                        'Call calculaTurnos(Val(hrs_hrDesde), Val(tur_hrDesde), Val(hrs_TpoTur))
                    Call calculaTurnos(hrs_hrDesde, tur_hrDesde, hrs_TpoTur)
                    hrs_hrDesde = tur_hrHasta
                                   
                Case Val(hrs_hrDesde2) To Val(hrs_hrHasta2) 'segundo turno
                    
                    'Call calculaTurnos(Val(hrs_hrDesde2), Val(tur_hrDesde), Val(hrs_TpoTur))
                    Call calculaTurnos(hrs_hrDesde2, tur_hrDesde, hrs_TpoTur)
                    hrs_hrDesde2 = tur_hrHasta
                    
            End Select
           
            tb.MoveNext
        Next
    End If
    
    tb.Close
    tb1.Close

'calculo turnos disponibles desde el ultimo turno asignado hasta el final del
'dia de trabajo del profesional
Call calculaTurnos(hrs_hrDesde, hrs_hrHasta, hrs_TpoTur)

If f_getTurnos = 2 Then
    Call calculaTurnos(hrs_hrDesde2, hrs_hrHasta2, hrs_TpoTur)
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Public Sub f_eliminarTurno()
Dim sMsg As String
sMsg = MsgBox("¿Esta seguro que desea eliminar el registro actual?", vbYesNo, "Eliminar")

If sMsg = vbYes Then

    Adodc1.Recordset.Update "tmp_idpaciente", Null
    Adodc1.Recordset.Update "tmp_nombre", Null
    Adodc1.Recordset.Update "tmp_observacion", Null
    
    estadoAbm = 3

End If

End Sub

Function f_ValidaCarga() As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''
'Valida que el profesional seleccionado para ver los turnos posea los datos necesarios previamente cargados
'''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strQuery, sText As String

f_ValidaCarga = True

sText = "El profesional seleccionado no tiene cargado: " & vbCrLf & vbTab

strQuery = " select * from horarios where hrs_idprof = " & DataCombo1.BoundText & " and hrs_dia = " & Weekday(DTPicker1.Value, vbMonday) - 1

Set tb1 = dbdiet.OpenRecordset(strQuery)

If tb1.RecordCount <> 0 Then

    If IsNull(tb1.Fields("hrs_hrdesde").Value) Then
           sText = sText & " - Horario de inicio laboral" & vbCrLf & vbTab
           f_ValidaCarga = False
    End If

    If IsNull(tb1.Fields("hrs_hrhasta").Value) Then
           sText = sText & " - Horario de finalizacion laboral" & vbCrLf & vbTab
           f_ValidaCarga = False
    End If

    If IsNull(tb1.Fields("hrs_tpotur").Value) Then
           sText = sText & " - Duracion del turno" & vbCrLf '& vbTab
           f_ValidaCarga = False
    End If

    If f_ValidaCarga = False Then
        sText = sText & "Verifique los datos. Proceso abortado"
        MsgBox sText, vbInformation
    End If
        
Else
    
    f_ValidaCarga = False

End If

tb1.Close
End Function

Function f_getTurnos() As Integer 'devuelve 1 o 2 segun la cantidad de turnos
Dim strQuery As String

f_getTurnos = 1

strQuery = " select * from horarios where hrs_idprof = " & DataCombo1.BoundText & " and hrs_dia = " & Weekday(DTPicker1.Value, vbMonday) - 1

Set tb1 = dbdiet.OpenRecordset(strQuery)

If tb1.RecordCount <> 0 Then

    If IsNull(tb1.Fields("hrs_hrdesde2").Value) And IsNull(tb1.Fields("hrs_hrhasta2").Value) Then
        f_getTurnos = 1
    Else
        f_getTurnos = 2
    End If

End If
    
End Function

Public Sub f_imprimirTurno()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmd_salir.Enabled = True Then

    Call f_Cancela

End If

End Sub

Sub f_Cancela()

    BeginTrans 'para forzar el borrado del buffer y poder usar el
               'procedimiento "DatagridRefresh"
        'se elimina el contenido de la tabla
        Call DeleteContTable("turnos_tmp")
    
    CommitTrans
    
    Call DatagridRefresh(Adodc1, DataGrid1)
    '
    DataGrid1.Enabled = False
    'DataGrid1.BorderStyle = dbgNoBorder
    'DataGrid1.Appearance = dbgFlat
    'DataGrid1.BackColor = &H8000000F
    
    'Frame3.Visible = False
    Frame8.Enabled = True
    DataCombo1.SetFocus
    
    cmd_aceptar.Enabled = True
    cmd_Cancelar.Enabled = True
    
    cmdImprimir.Enabled = False
    cmdAceptar.Enabled = False
    cmd_deshacer.Enabled = False
    
    cmd_Calendario.Enabled = True
    Command3.Enabled = False
    cmd_salir.Enabled = False
    
    cbo_estado.Enabled = False

    txt_legajo.Text = ""
    txt_nombrePaciente.Text = ""
    txt_observacion.Text = ""
    
    estadoAbm = 1
    cmdAceptar.ToolTipText = "Aceptar"

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Aceptar

End Sub

Private Sub Pic_Calendario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Calendario

End Sub

Private Sub Pic_cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cancel

End Sub

Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cancelar

End Sub

Private Sub Pic_Deshacer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Deshacer

End Sub

Private Sub Pic_Imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Imprimir

End Sub

Private Sub Pic_Salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Salir

End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Deshacer

End Sub

Private Sub PicAceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call fAceptar

End Sub

Private Sub txt_nombrePaciente_GotFocus()

txt_nombrePaciente.SelStart = 0
txt_nombrePaciente.SelLength = 50

End Sub

Private Sub txt_nombrePaciente_LostFocus()
Dim strQuery As String
'en el caso de que el nombre del paciente se modifique manualmente el legajo se blanquea
'ya que el paciente no esta cargado
'solo verifica que el nombre ingresado sea un paciente cargado de lo contrario borra el legajo
strQuery = "select (apell & ', ' & nombre) as nom from pacientes order by apell, nombre"

Set tb = dbdiet.OpenRecordset(strQuery)
tb.FindFirst "nom = '" & txt_nombrePaciente.Text & "'"

If tb.NoMatch = True Then
    txt_legajo.Text = ""
End If

tb.Close

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Data1.Recordset = Nothing
Set Me.Adodc1.Recordset = Nothing
Set Me.Adodc2.Recordset = Nothing
Set Me.Adodc3.Recordset = Nothing
Set Me.Adodc4.Recordset = Nothing
Set Me.Adodc5.Recordset = Nothing

strQuery = "Select * From Profesionales"
Call f_Data_DatabaseName(Data1, strQuery)

strQuery = "select * from turnos_tmp order by tmp_hrdesde"
Call f_Adodc_ConnectionString(Adodc1, strQuery)

strQuery = "Pacientes"
Call f_Adodc_ConnectionString(Adodc2, strQuery)

strQuery = "select *, (apell & ', ' & nombre) as nom from pacientes order by apell, nombre"
Call f_Adodc_ConnectionString(Adodc3, strQuery)
           
strQuery = "select *, (prf_apell & ', ' & prf_nombre) as nom from profesionales order by prf_apell, prf_nombre"
Call f_Adodc_ConnectionString(Adodc4, strQuery)

strQuery = "select * from turnos_tmp order by tmp_hrdesde"
Call f_Adodc_ConnectionString(Adodc5, strQuery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo1, Data1, Adodc4, "prf_codigo", "prf_codigo", "nom")

Call f_Enlaza_ControlData(DataList1, Adodc5, Adodc3, "tmp_idpaciente", "legajo", "nom")
'==============================================

End Sub

Sub f_Aceptar()

Me.cmd_aceptar.ZOrder 0

Me.cmd_Cancelar.ZOrder 1

End Sub

Sub f_Cancelar()

Me.cmd_aceptar.ZOrder 1
Me.cmd_Cancelar.ZOrder 0

End Sub

Sub fAceptar()

Me.cmdImprimir.ZOrder 1
Me.cmd_cancel.ZOrder 1
Me.cmd_deshacer.ZOrder 1
Me.cmd_salir.ZOrder 1
Me.cmdAceptar.ZOrder 0

End Sub

Sub f_Deshacer()

Me.cmdImprimir.ZOrder 1
Me.cmd_cancel.ZOrder 1
Me.cmd_deshacer.ZOrder 0
Me.cmd_salir.ZOrder 1
Me.cmdAceptar.ZOrder 1

End Sub

Sub f_Cancel()

Me.cmdImprimir.ZOrder 1
Me.cmd_cancel.ZOrder 0
Me.cmd_deshacer.ZOrder 1
Me.cmd_salir.ZOrder 1
Me.cmdAceptar.ZOrder 1

End Sub

Sub f_Salir()

Me.cmdImprimir.ZOrder 1
Me.cmd_cancel.ZOrder 1
Me.cmd_deshacer.ZOrder 1
Me.cmd_salir.ZOrder 0
Me.cmdAceptar.ZOrder 1

End Sub

Sub f_Imprimir()

Me.cmdImprimir.ZOrder 0
Me.cmd_cancel.ZOrder 1
Me.cmd_deshacer.ZOrder 1
Me.cmd_salir.ZOrder 1
Me.cmdAceptar.ZOrder 1

End Sub

Sub f_Calendario()

Me.cmd_Calendario.ZOrder 0

End Sub

Sub f_Boton_Zorder()

If Me.cmd_Calendario.Enabled = True Then
    Me.Pic_Calendario.ZOrder 0
Else
    Me.Pic_Calendario_Gris.ZOrder 0
End If

If Me.cmd_aceptar.Enabled = True Then
    Me.Pic_Aceptar.ZOrder 0
Else
    Me.Pic_Aceptar_Gris.ZOrder 0
End If

If Me.cmd_Cancelar.Enabled = True Then
    Me.Pic_Cancelar.ZOrder 0
Else
    Me.Pic_Cancelar_Gris.ZOrder 0
End If

If Me.cmdImprimir.Enabled = True Then
    Me.Pic_Imprimir.ZOrder 0
Else
    Me.Pic_Imprimir_Gris.ZOrder 0
End If

If Me.cmdAceptar.Enabled = True Then
    Me.PicAceptar.ZOrder 0
Else
    Me.PicAceptar_Gris.ZOrder 0
End If

If Me.cmd_salir.Enabled = True Then
    Me.Pic_Salir.ZOrder 0
Else
    Me.Pic_Salir_Gris.ZOrder 0
End If

If Me.cmd_deshacer.Enabled = True Then
    Me.Pic_Deshacer.ZOrder 0
Else
    Me.Pic_Deshacer_Gris.ZOrder 0
End If

If Me.cmd_cancel.Enabled = True Then
    Me.Pic_cancel.ZOrder 0
Else
    Me.Pic_cancel_Gris.ZOrder 0
End If


Me.cmd_aceptar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmdImprimir.ZOrder 1
Me.cmd_cancel.ZOrder 1
Me.cmd_deshacer.ZOrder 1
Me.cmd_salir.ZOrder 1
Me.cmdAceptar.ZOrder 1

End Sub


