VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_abm_horarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horarios de Profesionales"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "frm_abm_horarios.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   9990
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   1680
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   123
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Horarios:"
      Height          =   3495
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   9975
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   6
         ItemData        =   "frm_abm_horarios.frx":0ECA
         Left            =   3480
         List            =   "frm_abm_horarios.frx":0ED4
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   5
         ItemData        =   "frm_abm_horarios.frx":0EEB
         Left            =   3480
         List            =   "frm_abm_horarios.frx":0EF5
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         ItemData        =   "frm_abm_horarios.frx":0F0C
         Left            =   3480
         List            =   "frm_abm_horarios.frx":0F16
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         ItemData        =   "frm_abm_horarios.frx":0F2D
         Left            =   3480
         List            =   "frm_abm_horarios.frx":0F37
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         ItemData        =   "frm_abm_horarios.frx":0F4E
         Left            =   3480
         List            =   "frm_abm_horarios.frx":0F58
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "frm_abm_horarios.frx":0F6F
         Left            =   3480
         List            =   "frm_abm_horarios.frx":0F79
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         ItemData        =   "frm_abm_horarios.frx":0F90
         Left            =   3480
         List            =   "frm_abm_horarios.frx":0F9A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   10560
         TabIndex        =   80
         Text            =   "Text2"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   10560
         TabIndex        =   79
         Text            =   "Text2"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   10560
         TabIndex        =   78
         Text            =   "Text2"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   10560
         TabIndex        =   77
         Text            =   "Text2"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   10560
         TabIndex        =   76
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   10560
         TabIndex        =   75
         Text            =   "Text2"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   10560
         TabIndex        =   74
         Text            =   "Text2"
         Top             =   600
         Width           =   615
      End
      Begin MSAdodcLib.Adodc ado_Consultorio 
         Height          =   330
         Index           =   0
         Left            =   5760
         Top             =   0
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
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
         Caption         =   "ado_Consultorio"
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
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_abm_horarios.frx":0FB1
         DataField       =   "hrs_consul"
         DataSource      =   "DataDia(0)"
         Height          =   315
         Index           =   0
         Left            =   8400
         TabIndex        =   7
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         ListField       =   "con_descrip"
         BoundColumn     =   "con_codigo"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataSource      =   "DataDia(6)"
         ForeColor       =   &H8000000C&
         Height          =   285
         Index           =   34
         Left            =   10080
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataSource      =   "DataDia(5)"
         ForeColor       =   &H8000000C&
         Height          =   285
         Index           =   33
         Left            =   10080
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataSource      =   "DataDia(4)"
         Height          =   285
         Index           =   32
         Left            =   10080
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataSource      =   "DataDia(3)"
         Height          =   285
         Index           =   31
         Left            =   10080
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataSource      =   "DataDia(2)"
         Height          =   285
         Index           =   30
         Left            =   10080
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataSource      =   "DataDia(1)"
         Height          =   285
         Index           =   29
         Left            =   10080
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataSource      =   "DataDia(0)"
         Height          =   285
         Index           =   28
         Left            =   10080
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   600
         Width           =   975
      End
      Begin VB.Data DataDia 
         Caption         =   "DataLunes"
         Connect         =   "Access"
         DatabaseName    =   "db1nueva prueba anterior sin replica"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   0
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Horarios where hrs_dia = 0"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Data DataDia 
         Caption         =   "DataDomingo"
         Connect         =   "Access"
         DatabaseName    =   "db1nueva prueba anterior sin replica"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   6
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Horarios where hrs_dia = 6"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Data DataDia 
         Caption         =   "DataSabado"
         Connect         =   "Access"
         DatabaseName    =   "db1nueva prueba anterior sin replica"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   5
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Horarios where hrs_dia = 5"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Data DataDia 
         Caption         =   "DataViernes"
         Connect         =   "Access"
         DatabaseName    =   "db1nueva prueba anterior sin replica"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   4
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Horarios where hrs_dia = 4"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Data DataDia 
         Caption         =   "DataJueves"
         Connect         =   "Access"
         DatabaseName    =   "db1nueva prueba anterior sin replica"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   3
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Horarios where hrs_dia = 3"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Data DataDia 
         Caption         =   "DataMiercoles"
         Connect         =   "Access"
         DatabaseName    =   "db1nueva prueba anterior sin replica"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   2
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Horarios where hrs_dia = 2"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Data DataDia 
         Caption         =   "DataMartes"
         Connect         =   "Access"
         DatabaseName    =   "db1nueva prueba anterior sin replica"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Index           =   1
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Horarios where hrs_dia = 1"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "hrs_tpoTur"
         DataSource      =   "DataDia(6)"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   20
         Left            =   7680
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   2880
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "hrs_tpoTur"
         DataSource      =   "DataDia(5)"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   19
         Left            =   7680
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   2520
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "hrs_tpoTur"
         DataSource      =   "DataDia(4)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   18
         Left            =   7680
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   2160
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "hrs_tpoTur"
         DataSource      =   "DataDia(3)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   7680
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1800
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "hrs_tpoTur"
         DataSource      =   "DataDia(2)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   7680
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1440
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "hrs_tpoTur"
         DataSource      =   "DataDia(1)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   7680
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1080
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "hrs_tpoTur"
         DataSource      =   "DataDia(0)"
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   7680
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   720
         Width           =   585
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   240
         Top             =   3120
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
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   "Alimentos anterior sin replica.UDL"
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
      Begin MSDataListLib.DataList DataList1 
         Height          =   2790
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   4921
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483634
         ListField       =   ""
         BoundColumn     =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_abm_horarios.frx":0FD2
         DataField       =   "hrs_consul"
         DataSource      =   "DataDia(1)"
         Height          =   315
         Index           =   1
         Left            =   8400
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ListField       =   "con_descrip"
         BoundColumn     =   "con_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_abm_horarios.frx":0FF3
         DataField       =   "hrs_consul"
         DataSource      =   "DataDia(2)"
         Height          =   315
         Index           =   2
         Left            =   8400
         TabIndex        =   21
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ListField       =   "con_descrip"
         BoundColumn     =   "con_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_abm_horarios.frx":1014
         DataField       =   "hrs_consul"
         DataSource      =   "DataDia(3)"
         Height          =   315
         Index           =   3
         Left            =   8400
         TabIndex        =   28
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ListField       =   "con_descrip"
         BoundColumn     =   "con_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_abm_horarios.frx":1035
         DataField       =   "hrs_consul"
         DataSource      =   "DataDia(4)"
         Height          =   315
         Index           =   4
         Left            =   8400
         TabIndex        =   35
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ListField       =   "con_descrip"
         BoundColumn     =   "con_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_abm_horarios.frx":1056
         DataField       =   "hrs_consul"
         DataSource      =   "DataDia(5)"
         Height          =   315
         Index           =   5
         Left            =   8400
         TabIndex        =   42
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ForeColor       =   255
         ListField       =   "con_descrip"
         BoundColumn     =   "con_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_abm_horarios.frx":1077
         DataField       =   "hrs_consul"
         DataSource      =   "DataDia(6)"
         Height          =   315
         Index           =   6
         Left            =   8400
         TabIndex        =   49
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ForeColor       =   255
         ListField       =   "con_descrip"
         BoundColumn     =   "con_codigo"
         Text            =   "DataCombo1"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrDesde"
         DataSource      =   "DataDia(1)"
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   9
         Top             =   1080
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
         DataField       =   "hrs_hrDesde"
         DataSource      =   "DataDia(2)"
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   16
         Top             =   1440
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
         DataField       =   "hrs_hrDesde"
         DataSource      =   "DataDia(3)"
         Height          =   285
         Index           =   3
         Left            =   4800
         TabIndex        =   23
         Top             =   1800
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
         DataField       =   "hrs_hrDesde"
         DataSource      =   "DataDia(4)"
         Height          =   285
         Index           =   4
         Left            =   4800
         TabIndex        =   30
         Top             =   2160
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
         DataField       =   "hrs_hrDesde"
         DataSource      =   "DataDia(5)"
         Height          =   285
         Index           =   5
         Left            =   4800
         TabIndex        =   37
         Top             =   2520
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrDesde"
         DataSource      =   "DataDia(6)"
         Height          =   285
         Index           =   6
         Left            =   4800
         TabIndex        =   44
         Top             =   2880
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrHasta"
         DataSource      =   "DataDia(0)"
         Height          =   285
         Index           =   7
         Left            =   5520
         TabIndex        =   3
         Top             =   720
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
         DataField       =   "hrs_hrHasta"
         DataSource      =   "DataDia(1)"
         Height          =   285
         Index           =   8
         Left            =   5520
         TabIndex        =   10
         Top             =   1080
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
         DataField       =   "hrs_hrHasta"
         DataSource      =   "DataDia(2)"
         Height          =   285
         Index           =   9
         Left            =   5520
         TabIndex        =   17
         Top             =   1440
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
         DataField       =   "hrs_hrHasta"
         DataSource      =   "DataDia(3)"
         Height          =   285
         Index           =   10
         Left            =   5520
         TabIndex        =   24
         Top             =   1800
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
         DataField       =   "hrs_hrHasta"
         DataSource      =   "DataDia(4)"
         Height          =   285
         Index           =   11
         Left            =   5520
         TabIndex        =   31
         Top             =   2160
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
         DataField       =   "hrs_hrHasta"
         DataSource      =   "DataDia(5)"
         Height          =   285
         Index           =   12
         Left            =   5520
         TabIndex        =   38
         Top             =   2520
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrHasta"
         DataSource      =   "DataDia(6)"
         Height          =   285
         Index           =   13
         Left            =   5520
         TabIndex        =   45
         Top             =   2880
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrDesde"
         DataSource      =   "DataDia(0)"
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   2
         Top             =   720
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
         DataField       =   "hrs_hrDesde2"
         DataSource      =   "DataDia(1)"
         Height          =   285
         Index           =   15
         Left            =   6240
         TabIndex        =   11
         Top             =   1080
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
         DataField       =   "hrs_hrDesde2"
         DataSource      =   "DataDia(2)"
         Height          =   285
         Index           =   16
         Left            =   6240
         TabIndex        =   18
         Top             =   1440
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
         DataField       =   "hrs_hrDesde2"
         DataSource      =   "DataDia(3)"
         Height          =   285
         Index           =   17
         Left            =   6240
         TabIndex        =   25
         Top             =   1800
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
         DataField       =   "hrs_hrDesde2"
         DataSource      =   "DataDia(4)"
         Height          =   285
         Index           =   18
         Left            =   6240
         TabIndex        =   32
         Top             =   2160
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
         DataField       =   "hrs_hrDesde2"
         DataSource      =   "DataDia(5)"
         Height          =   285
         Index           =   19
         Left            =   6240
         TabIndex        =   39
         Top             =   2520
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrDesde2"
         DataSource      =   "DataDia(6)"
         Height          =   285
         Index           =   20
         Left            =   6240
         TabIndex        =   46
         Top             =   2880
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrHasta2"
         DataSource      =   "DataDia(0)"
         Height          =   285
         Index           =   21
         Left            =   6960
         TabIndex        =   5
         Top             =   720
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
         DataField       =   "hrs_hrHasta2"
         DataSource      =   "DataDia(1)"
         Height          =   285
         Index           =   22
         Left            =   6960
         TabIndex        =   12
         Top             =   1080
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
         DataField       =   "hrs_hrHasta2"
         DataSource      =   "DataDia(2)"
         Height          =   285
         Index           =   23
         Left            =   6960
         TabIndex        =   19
         Top             =   1440
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
         DataField       =   "hrs_hrHasta2"
         DataSource      =   "DataDia(3)"
         Height          =   285
         Index           =   24
         Left            =   6960
         TabIndex        =   26
         Top             =   1800
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
         DataField       =   "hrs_hrHasta2"
         DataSource      =   "DataDia(4)"
         Height          =   285
         Index           =   25
         Left            =   6960
         TabIndex        =   33
         Top             =   2160
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
         DataField       =   "hrs_hrHasta2"
         DataSource      =   "DataDia(5)"
         Height          =   285
         Index           =   26
         Left            =   6960
         TabIndex        =   40
         Top             =   2520
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrHasta2"
         DataSource      =   "DataDia(6)"
         Height          =   285
         Index           =   27
         Left            =   6960
         TabIndex        =   47
         Top             =   2880
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrDesde2"
         DataSource      =   "DataDia(0)"
         Height          =   285
         Index           =   14
         Left            =   6240
         TabIndex        =   4
         Top             =   720
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
      Begin VB.Label Label2 
         Caption         =   "Turnos"
         Height          =   375
         Index           =   15
         Left            =   3600
         TabIndex        =   88
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Segundo Turno"
         Height          =   195
         Left            =   6275
         TabIndex        =   87
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Primer Turno"
         Height          =   195
         Left            =   4930
         TabIndex        =   86
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "Consultorio"
         Height          =   255
         Index           =   11
         Left            =   8520
         TabIndex        =   73
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Duracion Turno (min.)"
         Height          =   375
         Index           =   10
         Left            =   7560
         TabIndex        =   63
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   13
         Left            =   6960
         TabIndex        =   84
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Index           =   14
         Left            =   6240
         TabIndex        =   85
         Top             =   480
         Width           =   855
      End
      Begin VB.Shape Shape1 
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   2775
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label2 
         Caption         =   "hrs_dia"
         Height          =   255
         Index           =   20
         Left            =   10560
         TabIndex        =   81
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "hrs_idProf"
         Height          =   255
         Index           =   12
         Left            =   10080
         TabIndex        =   71
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   9
         Left            =   5520
         TabIndex        =   62
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   255
         Index           =   8
         Left            =   4800
         TabIndex        =   61
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Domingo"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   60
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Sabado"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   59
         Top             =   2550
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Viernes"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   58
         Top             =   2190
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Jueves"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   57
         Top             =   1830
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Miercoles"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   56
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Martes"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   55
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Lunes"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   54
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Dia"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5640
      TabIndex        =   83
      Text            =   "Text3"
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4320
      TabIndex        =   82
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ac&tualizar"
      Height          =   255
      Left            =   9960
      TabIndex        =   72
      ToolTipText     =   "Mostrar Todos"
      Top             =   2280
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Frame Frame2 
      Caption         =   "consultorios"
      Height          =   495
      Left            =   1680
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Label Label1 
         Caption         =   "Label1"
         DataField       =   "hrs_codigo"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   840
         TabIndex        =   52
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame fme_botones_abm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   2228
      TabIndex        =   89
      Top             =   3480
      Width           =   5535
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":1098
         Height          =   375
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":11F1
         Picture         =   "frm_abm_horarios.frx":1343
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Aceptar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":15FF
         Height          =   375
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":1793
         Picture         =   "frm_abm_horarios.frx":18E5
         Style           =   1  'Graphical
         TabIndex        =   116
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
         MouseIcon       =   "frm_abm_horarios.frx":1D98
         Picture         =   "frm_abm_horarios.frx":1EEA
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   115
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
         MouseIcon       =   "frm_abm_horarios.frx":21EB
         Picture         =   "frm_abm_horarios.frx":233D
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   114
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":25F9
         Height          =   375
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":271A
         Picture         =   "frm_abm_horarios.frx":286C
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Modificar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":2ADF
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":2BF5
         Picture         =   "frm_abm_horarios.frx":2D47
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Eliminar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":2ED6
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":3023
         Picture         =   "frm_abm_horarios.frx":3175
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Agregar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUltimo 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":35AF
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":3757
         Picture         =   "frm_abm_horarios.frx":38A9
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Ultimo"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSiguiente 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":3D74
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":3EE1
         Picture         =   "frm_abm_horarios.frx":4033
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Siguiente"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":44A8
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":4630
         Picture         =   "frm_abm_horarios.frx":4782
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Buscar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":4A5F
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":4BC9
         Picture         =   "frm_abm_horarios.frx":4D1B
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Anterior"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimero 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":5189
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frm_abm_horarios.frx":532E
         Picture         =   "frm_abm_horarios.frx":5480
         Style           =   1  'Graphical
         TabIndex        =   98
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
         MouseIcon       =   "frm_abm_horarios.frx":593B
         Picture         =   "frm_abm_horarios.frx":5A8D
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   97
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
         MouseIcon       =   "frm_abm_horarios.frx":5F48
         Picture         =   "frm_abm_horarios.frx":609A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   96
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
         MouseIcon       =   "frm_abm_horarios.frx":6508
         Picture         =   "frm_abm_horarios.frx":665A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   95
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
         MouseIcon       =   "frm_abm_horarios.frx":6937
         Picture         =   "frm_abm_horarios.frx":6A89
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   94
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
         MouseIcon       =   "frm_abm_horarios.frx":6EFE
         Picture         =   "frm_abm_horarios.frx":7050
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   93
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
         MouseIcon       =   "frm_abm_horarios.frx":751B
         Picture         =   "frm_abm_horarios.frx":766D
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   92
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
         MouseIcon       =   "frm_abm_horarios.frx":7AA7
         Picture         =   "frm_abm_horarios.frx":7BF9
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   91
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
         MouseIcon       =   "frm_abm_horarios.frx":7D88
         Picture         =   "frm_abm_horarios.frx":7EDA
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   90
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
         MouseIcon       =   "frm_abm_horarios.frx":814D
         Picture         =   "frm_abm_horarios.frx":829F
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   106
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
         MouseIcon       =   "frm_abm_horarios.frx":83C0
         Picture         =   "frm_abm_horarios.frx":8512
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   107
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
         MouseIcon       =   "frm_abm_horarios.frx":8628
         Picture         =   "frm_abm_horarios.frx":877A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   108
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
         MouseIcon       =   "frm_abm_horarios.frx":88C7
         Picture         =   "frm_abm_horarios.frx":8A19
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   109
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
         MouseIcon       =   "frm_abm_horarios.frx":8BC1
         Picture         =   "frm_abm_horarios.frx":8D13
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   110
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
         MouseIcon       =   "frm_abm_horarios.frx":8E80
         Picture         =   "frm_abm_horarios.frx":8FD2
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   111
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
         MouseIcon       =   "frm_abm_horarios.frx":915A
         Picture         =   "frm_abm_horarios.frx":92AC
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   112
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
         MouseIcon       =   "frm_abm_horarios.frx":9416
         Picture         =   "frm_abm_horarios.frx":9568
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   113
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
         MouseIcon       =   "frm_abm_horarios.frx":970D
         Picture         =   "frm_abm_horarios.frx":985F
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   118
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
         MouseIcon       =   "frm_abm_horarios.frx":99B8
         Picture         =   "frm_abm_horarios.frx":9B0A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   119
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frm_abm_horarios.frx":9C9E
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frm_abm_horarios.frx":9DF6
         Style           =   1  'Graphical
         TabIndex        =   121
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
         MouseIcon       =   "frm_abm_horarios.frx":A276
         Picture         =   "frm_abm_horarios.frx":A3C8
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   120
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
         MouseIcon       =   "frm_abm_horarios.frx":A848
         Picture         =   "frm_abm_horarios.frx":A99A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   122
         Top             =   120
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_abm_horarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Titulo As String 'titulo del form
Dim Profesional As Long 'id del profesional

Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
                            
Dim MaskEdBox_14, MaskEdBox_21 As String
Dim MaskEdBox_15, MaskEdBox_22 As String
Dim MaskEdBox_16, MaskEdBox_23 As String
Dim MaskEdBox_17, MaskEdBox_24 As String
Dim MaskEdBox_18, MaskEdBox_25 As String
Dim MaskEdBox_19, MaskEdBox_26 As String
Dim MaskEdBox_20, MaskEdBox_27 As String


Private Sub Combo1_Click(Index As Integer)
If estadoAbm = 3 Then

    Select Case Index
        Case Is = 0
                If MaskEdBox1(14).Text <> "" Then
                    MaskEdBox_14 = MaskEdBox1(14).Text
                End If
                
                If MaskEdBox1(21).Text <> "" Then
                     MaskEdBox_21 = MaskEdBox1(21).Text
                End If
                
        Case Is = 1
                If MaskEdBox1(15).Text <> "" Then
                    MaskEdBox_15 = MaskEdBox1(15).Text
                End If
                
                If MaskEdBox1(22).Text <> "" Then
                    MaskEdBox_22 = MaskEdBox1(22).Text
                End If
        
        Case Is = 2
                If MaskEdBox1(16).Text <> "" Then
                    MaskEdBox_16 = MaskEdBox1(16).Text
                End If
                    
                If MaskEdBox1(23).Text <> "" Then
                    MaskEdBox_23 = MaskEdBox1(23).Text
                End If
        
        Case Is = 3
                If MaskEdBox1(17).Text <> "" Then
                    MaskEdBox_17 = MaskEdBox1(17).Text
                End If
                    
                If MaskEdBox1(24).Text <> "" Then
                    MaskEdBox_24 = MaskEdBox1(24).Text
                End If
        
        Case Is = 4
                If MaskEdBox1(18).Text <> "" Then
                    MaskEdBox_18 = MaskEdBox1(18).Text
                End If
                
                If MaskEdBox1(25).Text <> "" Then
                    MaskEdBox_25 = MaskEdBox1(25).Text
                End If
        
        Case Is = 5
                If MaskEdBox1(19).Text <> "" Then
                    MaskEdBox_19 = MaskEdBox1(19).Text
                End If
                
                If MaskEdBox1(26).Text <> "" Then
                    MaskEdBox_26 = MaskEdBox1(26).Text
                End If
                
        Case Is = 6
                If MaskEdBox1(20).Text <> "" Then
                    MaskEdBox_20 = MaskEdBox1(20).Text
                End If
                    
                If MaskEdBox1(27).Text <> "" Then
                    MaskEdBox_27 = MaskEdBox1(27).Text
                End If
    
    End Select
                
    Select Case Combo1(Index).ListIndex
        Case Is = 0 '"1 turno"
            Select Case Index
                Case Is = 0
                    MaskEdBox1(14).Enabled = False
                    MaskEdBox1(21).Enabled = False
                   
                    MaskEdBox1(14).Text = ""
                    MaskEdBox1(21).Text = ""
                Case Is = 1
                    MaskEdBox1(15).Enabled = False
                    MaskEdBox1(22).Enabled = False
                    
                    MaskEdBox1(15).Text = ""
                    MaskEdBox1(22).Text = ""
                Case Is = 2
                    MaskEdBox1(16).Enabled = False
                    MaskEdBox1(23).Enabled = False
                                                    
                    MaskEdBox1(16).Text = ""
                    MaskEdBox1(23).Text = ""
                Case Is = 3
                    MaskEdBox1(17).Enabled = False
                    MaskEdBox1(24).Enabled = False
                                    
                    MaskEdBox1(17).Text = ""
                    MaskEdBox1(24).Text = ""
                Case Is = 4
                    MaskEdBox1(18).Enabled = False
                    MaskEdBox1(25).Enabled = False
                                    
                    MaskEdBox1(18).Text = ""
                    MaskEdBox1(25).Text = ""
                Case Is = 5
                    MaskEdBox1(19).Enabled = False
                    MaskEdBox1(26).Enabled = False
                    
                    MaskEdBox1(19).Text = ""
                    MaskEdBox1(26).Text = ""
                Case Is = 6
                    MaskEdBox1(20).Enabled = False
                    MaskEdBox1(27).Enabled = False
                                    
                    MaskEdBox1(20).Text = ""
                    MaskEdBox1(27).Text = ""
            End Select
        Case Is = 1 '"2 turnos"
            Select Case Index
                Case Is = 0
                    MaskEdBox1(14).Enabled = True
                    MaskEdBox1(21).Enabled = True
                    
                    MaskEdBox1(14).Text = MaskEdBox_14
                    MaskEdBox1(21).Text = MaskEdBox_21
                    
                Case Is = 1
                    MaskEdBox1(15).Enabled = True
                    MaskEdBox1(22).Enabled = True
                    
                    MaskEdBox1(15).Text = MaskEdBox_15
                    MaskEdBox1(22).Text = MaskEdBox_22
                Case Is = 2
                    MaskEdBox1(16).Enabled = True
                    MaskEdBox1(23).Enabled = True
                    
                    MaskEdBox1(16).Text = MaskEdBox_16
                    MaskEdBox1(23).Text = MaskEdBox_23
                Case Is = 3
                    MaskEdBox1(17).Enabled = True
                    MaskEdBox1(24).Enabled = True
                    
                    MaskEdBox1(17).Text = MaskEdBox_17
                    MaskEdBox1(24).Text = MaskEdBox_24
                Case Is = 4
                    MaskEdBox1(18).Enabled = True
                    MaskEdBox1(25).Enabled = True
                    
                    MaskEdBox1(18).Text = MaskEdBox_18
                    MaskEdBox1(25).Text = MaskEdBox_25
                Case Is = 5
                    MaskEdBox1(19).Enabled = True
                    MaskEdBox1(26).Enabled = True
                    
                    MaskEdBox1(19).Text = MaskEdBox_19
                    MaskEdBox1(26).Text = MaskEdBox_26
                Case Is = 6
                    MaskEdBox1(20).Enabled = True
                    MaskEdBox1(27).Enabled = True
                    
                    MaskEdBox1(20).Text = MaskEdBox_20
                    MaskEdBox1(27).Text = MaskEdBox_27
            End Select
    End Select

End If
End Sub

Private Sub Command1_Click()
Dim strQuery As String
strQuery = " select * from consultorios order by con_descrip, con_dir"

With Data1
    .RecordSource = strQuery
    .Refresh
End With

Call enabledDesplaz
End Sub

Private Sub Data1_Validate(Index As Integer, Action As Integer, Save As Integer)

End Sub

Private Sub Command2_Click()
'For i = 0 To 6
'    Text1(i).Text = CDate(Text1(i).Text)
    'Text1(i).Text = Format(Text1(i).Text, "h:mm")
'    Text1(i).DataFormat.Format ("short time")
' Text1(0).Text = Format(Text1(0).Text, "standar")
'Next

MsgBox Text3.Text Mod 100

End Sub


Private Sub DataCombo1_LostFocus(Index As Integer)

If DataCombo1(Index).Text = "" Then
    
    DataCombo1(Index).BoundText = DataDia(Index).Recordset.Fields("hrs_consul").Value
    
End If

End Sub

Private Sub DataList1_Click()
'For i = 28 To 34
'    Text1(i).Text = DataList1.BoundText
'Next
'
'For i = 0 To 6
'    DataDia(i).Refresh
'Next

Call ValorDefault

End Sub

Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 4455
Me.Width = 10110
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

'Call enabledDesplaz 'verifica y establece que botones de desplazamiento permanecen habilitados y culaes deshabilitados

Call f_Boton_Zorder

End Sub

Private Sub Form_Load()

Call f_CargarOrigenDatos

estadoAbm = 1 ' el estado es sim cambios

Titulo = Me.Caption
Profesional = 0

Call ValorDefault

For i = 0 To 6
    Combo1(i).ListIndex = 1 '2 turnos
Next

'-------------------------
'se refresca el data1 para que el metodo enabledDesplaz funcione correctamente con el recordset cargado
'strquery = " select * from horarios "
'
'With Data1
'    .RecordSource = strquery
'    .Refresh
'End With
'--------------------------------

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdCancelar_Click

End Sub

Private Sub Label1_Change()

Me.Caption = Titulo & " - Nro. " & Val(Label1.Caption)

End Sub

Private Sub cmdAceptar_Click()

If estadoAbm = 2 Or estadoAbm = 3 Then 'si el estado es agregar o modificar

    For i = 0 To 6
        DataDia(i).UpdateRecord
        DataDia(i).Recordset.Bookmark = DataDia(i).Recordset.LastModified
    Next
    
    'Call abm("UpdateRecord")
     
'    For i = 0 To 6
'        Call abm("Bookmark = DataDia(" & i & ").Recordset.LastModified")
'    Next
          
'    'condiciones extras
'    If estadoAbm = 2 Then
'        dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(MDIForm1.ActiveForm.Label1.Caption) & ", codalimento from alimentos where estado = true"
'        dbdiet.Execute "insert into histclinicas (legajo) select " & Val(MDIForm1.ActiveForm.Label1.Caption) '& ", codalimento from alimentos where estado = true"
'    End If
        
    'cmdBuscar.Enabled = True
    'cmdAgregar.Enabled = True
    'cmdBorrar.Enabled = True
    'cmdClose.Enabled = True
    cmdModificar.Enabled = True
    
    'cmdAgregar.SetFocus
    'cmdAgregar.default = True
    cmdCancelar.Cancel = True
    
    'cmdPrimero.Enabled = True
    'cmdAnterior.Enabled = True
    'cmdSiguiente.Enabled = True
    'cmdUltimo.Enabled = True
   
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)

    estadoAbm = 1 ' el estado del form es "sin cambios"
            
    Call RefrescaDataList(False, True)
    'Call enabledDesplaz
    
    Call f_Boton_Zorder
    
Else

    'MDIForm1.ActiveForm.Hide
    Unload Me
    
End If

End Sub

Private Sub cmdAgregar_Click()

Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

estadoAbm = 2 ' el estado es agregar

Call abm("AddNew")

cmdAgregar.Enabled = False
'cmdBorrar.Enabled = False
'cmdclose.Enabled = False
cmdModificar.Enabled = False
'cmdBuscar.Enabled = False
cmdAceptar.Visible = True
cmdCancelar.Visible = True
'cmdPrimero.Enabled = False
'cmdAnterior.Enabled = False
'cmdSiguiente.Enabled = False
'cmdUltimo.Enabled = False

Text1(0).SetFocus

cmdAceptar.Default = True
cmdCancelar.Cancel = True


Call f_Enabled2Turno

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

strQuery = " select * from consultorios order by con_descrip, con_dir"

With Data1
    .RecordSource = strQuery
    .Refresh
End With

'aclare campo por el cual buscar
msg = InputBox("Ingrese la descripcion del consultorio:", "Buscar por descripcion del consultorio")
    
If msg <> "" Then
    
    strQuery = " select * from consultorios where con_descrip like '" & msg & "*' order by con_descrip, con_dir"
    
    With MDIForm1.ActiveForm.Data1
        .RecordSource = strQuery
        .Refresh
    End With

End If

Call enabledDesplaz

End Sub

Private Sub cmdCancelar_Click()
If estadoAbm = 2 Or estadoAbm = 3 Then ' el estado del form es agregar o modificar

    Call abm("CancelUpdate")
    
    'cmdBuscar.Enabled = True
    'cmdAgregar.Enabled = True
    'cmdBorrar.Enabled = True
    'cmdClose.Enabled = True
    cmdModificar.Enabled = True
    
    'cmdAgregar.SetFocus
    'cmdAgregar.default = True
    'cmdClose.Cancel = True
    'cmdPrimero.Enabled = True
    'cmdAnterior.Enabled = True
    'cmdSiguiente.Enabled = True
    'cmdUltimo.Enabled = True
           
    Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)
    
    Call RefrescaDataList(False, True)
    estadoAbm = 1 ' el estado del form es "sin cambios"
    
    'Call enabledDesplaz
    
    Call f_Boton_Zorder
    
Else

    'MDIForm1.ActiveForm.Hide
    Unload Me

End If
End Sub



Private Sub cmdImprimir_Click()
Dim strQuery As String

CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_horarios_one.rpt"

strQuery = " {horarios.hrs_idprof} = " & DataList1.BoundText

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub

Private Sub cmdModificar_Click()

If DataList1.BoundText <> "" Then

    Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)
    
    'If MDIForm1.ActiveForm.Data1.Recordset.BOF = True Or MDIForm1.ActiveForm.Data1.Recordset.EOF = True Then
    '    MDIForm1.ActiveForm.Data1.Recordset.MoveFirst
    'End If
    
    'cmdAgregar.Enabled = False
    'cmdBorrar.Enabled = False
    'cmdclose.Enabled = False
    cmdModificar.Enabled = False
    'cmdBuscar.Enabled = False
    cmdAceptar.Visible = True
    cmdCancelar.Visible = True
    'cmdPrimero.Enabled = False
    'cmdAnterior.Enabled = False
    'cmdSiguiente.Enabled = False
    'cmdUltimo.Enabled = False
    
    Profesional = DataList1.BoundText
    
    Call abm("Edit")
    
    Call RefrescaDataList(True, False)
    'MDIForm1.ActiveForm.txtFields(1).SetFocus
    
    cmdAceptar.Default = True
    cmdCancelar.Cancel = True
    
    estadoAbm = 3 ' el estado es modificar
    
    MaskEdBox1(0).SetFocus
    
    Call f_Enabled2Turno
        
    Call f_Boton_Zorder
    
Else

    MsgBox "Debe seleccionar un profesional", vbInformation, "Información"

End If

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

Sub ValorDefault()

For i = 28 To 34
    Text1(i).Text = DataList1.BoundText
Next

For i = 0 To 6
    
    strQuery = "Select * from horarios where hrs_idprof = " & DataList1.BoundText & " and hrs_dia = " & i
    
    With DataDia(i)
        .RecordSource = strQuery
        .Refresh
    End With
    
    Text2(i).Text = i
    
Next

End Sub

Sub abm(Accion As String)
Dim Result
'Result = ""

For i = 0 To 6
    Result = CallByName(DataDia(i).Recordset, Accion, VbMethod)

Next

End Sub

Sub RefrescaDataList(p_Locked As Boolean, p_Enabled As Boolean)
DataList1.Locked = p_Locked
DataList1.Enabled = p_Enabled
DataList1.Refresh

strQuery = "select *, (prf_apell & ', ' & prf_nombre) as nom from profesionales order by prf_apell, prf_nombre"
With Adodc4
    .RecordSource = strQuery
    .Refresh
End With

DataList1.BoundText = Profesional

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

Private Sub Pic_Imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Imprimir

End Sub

Private Sub Text1_GotFocus(Index As Integer)

Text1(Index).SelStart = 0
Text1(Index).SelLength = 50

End Sub

Sub f_Enabled2Turno()

If estadoAbm = 3 Then

    For i = 0 To 6
        Select Case Combo1(i).ListIndex
            Case Is = 0 '"1 turno"
                Select Case i
                    Case Is = 0
                        MaskEdBox1(14).Enabled = False
                        MaskEdBox1(21).Enabled = False
                    Case Is = 1
                        MaskEdBox1(15).Enabled = False
                        MaskEdBox1(22).Enabled = False
                    Case Is = 2
                        MaskEdBox1(16).Enabled = False
                        MaskEdBox1(23).Enabled = False
                    Case Is = 3
                        MaskEdBox1(17).Enabled = False
                        MaskEdBox1(24).Enabled = False
                    Case Is = 4
                        MaskEdBox1(18).Enabled = False
                        MaskEdBox1(25).Enabled = False
                    Case Is = 5
                        MaskEdBox1(19).Enabled = False
                        MaskEdBox1(26).Enabled = False
                    Case Is = 6
                        MaskEdBox1(20).Enabled = False
                        MaskEdBox1(27).Enabled = False
                End Select
            Case Is = 1 '"2 turnos"
                Select Case Index
                    Case Is = 0
                        MaskEdBox1(14).Enabled = True
                        MaskEdBox1(21).Enabled = True
                    Case Is = 1
                        MaskEdBox1(15).Enabled = True
                        MaskEdBox1(22).Enabled = True
                    Case Is = 2
                        MaskEdBox1(16).Enabled = True
                        MaskEdBox1(23).Enabled = True
                    Case Is = 3
                        MaskEdBox1(17).Enabled = True
                        MaskEdBox1(24).Enabled = True
                    Case Is = 4
                        MaskEdBox1(18).Enabled = True
                        MaskEdBox1(25).Enabled = True
                    Case Is = 5
                        MaskEdBox1(19).Enabled = True
                        MaskEdBox1(26).Enabled = True
                    Case Is = 6
                        MaskEdBox1(20).Enabled = True
                        MaskEdBox1(27).Enabled = True
                End Select
        End Select
    Next

End If

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

For i = 0 To 6
        
    Set Me.DataDia(i).Recordset = Nothing

    strQuery = "select * from Horarios where hrs_dia = " & i
    Call f_Data_DatabaseName(DataDia(i), strQuery)
    
Next

Set Me.ado_Consultorio(0).Recordset = Nothing
Set Me.Adodc4.Recordset = Nothing

strQuery = "Consultorios"
Call f_Adodc_ConnectionString(ado_Consultorio(0), strQuery)

'===============================
'En el caso de que el control adodc este enlazado a un datalist o un datacombo se deben blanquear las propiedades
'de enlace de dichos objetos y definirlas mediante codigo como a continuacion
strQuery = "select *, (prf_apell & ', ' & prf_nombre) as nom from profesionales order by prf_apell, prf_nombre"
Call f_Adodc_ConnectionString(Adodc4, strQuery)

Call f_Enlaza_ControlData(DataList1, Adodc4, Adodc4, "prf_codigo", "prf_codigo", "nom")
'Call f_Enlaza_Controles
'===============================
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

