VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_Adm_Diet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion Planes Alimentarios"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   ControlBox      =   0   'False
   Icon            =   "frm_Adm_Diet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11175
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   2520
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   11
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Menú"
      Enabled         =   0   'False
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   7335
      Begin VB.Frame Frame3 
         Caption         =   "Totales:"
         Height          =   615
         Left            =   480
         TabIndex        =   1
         Top             =   3960
         Width           =   6375
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Proporción en Kcal:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   5160
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Subtotal de Kcal:"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   495
         Left            =   2460
         TabIndex        =   29
         Top             =   4560
         Width           =   2295
         Begin VB.CommandButton cmd_Cancelar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Adm_Diet.frx":0ECA
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_Adm_Diet.frx":1042
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Deshacer cambios"
            Top             =   120
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmd_Salir 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Adm_Diet.frx":14C2
            Height          =   375
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_Adm_Diet.frx":1656
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Salir"
            Top             =   120
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmd_Guardar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Adm_Diet.frx":1957
            Enabled         =   0   'False
            Height          =   375
            Left            =   720
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_Adm_Diet.frx":1AB0
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Guardar cambios"
            Top             =   120
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Pic_Cancelar_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1200
            MouseIcon       =   "frm_Adm_Diet.frx":1D6C
            Picture         =   "frm_Adm_Diet.frx":1EBE
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   38
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
            MouseIcon       =   "frm_Adm_Diet.frx":2036
            Picture         =   "frm_Adm_Diet.frx":2188
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   35
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
            MouseIcon       =   "frm_Adm_Diet.frx":231C
            Picture         =   "frm_Adm_Diet.frx":246E
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   36
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
            MouseIcon       =   "frm_Adm_Diet.frx":25C7
            Picture         =   "frm_Adm_Diet.frx":2719
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   33
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
            MouseIcon       =   "frm_Adm_Diet.frx":2A1A
            Picture         =   "frm_Adm_Diet.frx":2B6C
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   37
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
            MouseIcon       =   "frm_Adm_Diet.frx":2FEC
            Picture         =   "frm_Adm_Diet.frx":313E
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   34
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdreporte 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Adm_Diet.frx":33FA
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_Adm_Diet.frx":3552
            Style           =   1  'Graphical
            TabIndex        =   40
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
            MouseIcon       =   "frm_Adm_Diet.frx":39D2
            Picture         =   "frm_Adm_Diet.frx":3B24
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   39
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
            MouseIcon       =   "frm_Adm_Diet.frx":3FA4
            Picture         =   "frm_Adm_Diet.frx":40F6
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   41
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   480
         Picture         =   "frm_Adm_Diet.frx":424E
         ScaleHeight     =   3585
         ScaleWidth      =   6465
         TabIndex        =   20
         Top             =   240
         Width           =   6495
         Begin VB.Label lbl_aviso 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seleccione un Paciente"
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
            Left            =   2280
            TabIndex        =   21
            Top             =   2160
            Width           =   2040
         End
      End
      Begin VB.CommandButton noCheck 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         MouseIcon       =   "frm_Adm_Diet.frx":681CC
         MousePointer    =   99  'Custom
         Picture         =   "frm_Adm_Diet.frx":6831E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Destildar Todo"
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CommandButton Check 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         MouseIcon       =   "frm_Adm_Diet.frx":68685
         MousePointer    =   99  'Custom
         Picture         =   "frm_Adm_Diet.frx":687D7
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Tildar Todo"
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3135
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5530
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   3615
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6376
         MultiRow        =   -1  'True
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   6
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Desayuno"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Almuerzo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Merienda"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Cena"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Colación 1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Colación 2"
               ImageVarType    =   2
            EndProperty
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
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame8 
         Caption         =   "Nombre:"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   2775
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frm_Adm_Diet.frx":688AF
            DataField       =   "Legajo"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "nom"
            BoundColumn     =   "Legajo"
            Text            =   "DataCombo1"
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   495
            Left            =   2160
            TabIndex        =   42
            Top             =   120
            Width           =   495
            Begin VB.CommandButton cmd_Tipito 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_Adm_Diet.frx":688C3
               Height          =   315
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frm_Adm_Diet.frx":68FD3
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Info"
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
               MouseIcon       =   "frm_Adm_Diet.frx":69263
               Picture         =   "frm_Adm_Diet.frx":693B5
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
               MouseIcon       =   "frm_Adm_Diet.frx":69645
               Picture         =   "frm_Adm_Diet.frx":69797
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   45
               Top             =   120
               Width           =   315
            End
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   735
         Left            =   5040
         TabIndex        =   22
         Top             =   120
         Width           =   2175
         Begin VB.CommandButton cmd_Aceptar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Adm_Diet.frx":698C7
            Height          =   375
            Left            =   1080
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_Adm_Diet.frx":69A20
            Picture         =   "frm_Adm_Diet.frx":69B72
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Aceptar"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmd_Cerrar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Adm_Diet.frx":69E2E
            Height          =   375
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_Adm_Diet.frx":69FC2
            Picture         =   "frm_Adm_Diet.frx":6A114
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Cancelar"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.PictureBox Pic_Aceptar_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1080
            MouseIcon       =   "frm_Adm_Diet.frx":6A5C7
            Picture         =   "frm_Adm_Diet.frx":6A719
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   25
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox Pic_Cerrar_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1560
            MouseIcon       =   "frm_Adm_Diet.frx":6A872
            Picture         =   "frm_Adm_Diet.frx":6A9C4
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   26
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox Pic_Aceptar 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1080
            MouseIcon       =   "frm_Adm_Diet.frx":6AB58
            Picture         =   "frm_Adm_Diet.frx":6ACAA
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   24
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox Pic_Cerrar 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1560
            MouseIcon       =   "frm_Adm_Diet.frx":6AF66
            Picture         =   "frm_Adm_Diet.frx":6B0B8
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   23
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmd_Calendario 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_Adm_Diet.frx":6B3B9
            Height          =   375
            Left            =   360
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_Adm_Diet.frx":6B839
            Picture         =   "frm_Adm_Diet.frx":6B98B
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Calendario"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.PictureBox Pic_Calendario 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   360
            MouseIcon       =   "frm_Adm_Diet.frx":6BE4C
            Picture         =   "frm_Adm_Diet.frx":6BF9E
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   46
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox Pic_Calendario_Gris 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   360
            MouseIcon       =   "frm_Adm_Diet.frx":6C45F
            Picture         =   "frm_Adm_Diet.frx":6C5B1
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   48
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Fecha:"
         Height          =   735
         Left            =   3000
         TabIndex        =   18
         Top             =   120
         Width           =   1815
         Begin MSComCtl2.DTPicker DTPicker1 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "d/MMM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   23461889
            CurrentDate     =   37858
         End
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rollback"
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frm_Adm_Diet.frx":6C741
      Height          =   5895
      Left            =   7440
      TabIndex        =   10
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   10398
      _Version        =   393216
      BackColor       =   -2147483628
      Cols            =   4
      GridLinesFixed  =   3
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   4
      _Band(0)._MapCol(0)._Name=   "idcategoria"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "Alimentos"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "Cant ideal (gr/cc)"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(2)._Alignment=   7
      _Band(0)._MapCol(3)._Name=   "Cant real (gr/cc)"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(3)._Alignment=   7
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   10680
      Top             =   1800
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
      Left            =   10680
      Top             =   1440
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
      Left            =   10680
      Top             =   1080
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
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   10680
      Top             =   360
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
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   11775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbl_Porcentaje 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_Porcentaje"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   5
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7560
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   11040
      MouseIcon       =   "frm_Adm_Diet.frx":6C756
      MousePointer    =   99  'Custom
      Picture         =   "frm_Adm_Diet.frx":6C8A8
      Stretch         =   -1  'True
      ToolTipText     =   "Contraer"
      Top             =   2918
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   7320
      MouseIcon       =   "frm_Adm_Diet.frx":6CCEA
      MousePointer    =   99  'Custom
      Picture         =   "frm_Adm_Diet.frx":6CE3C
      Stretch         =   -1  'True
      ToolTipText     =   "Expandir"
      Top             =   2918
      Width           =   135
   End
End
Attribute VB_Name = "frm_Adm_Diet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
Dim tb_FlexGrid As ADODB.Recordset
Dim tempNode As Node
Dim tempNode1 As Node
Dim Kcal
Dim Porcion
Dim AuxTab As Integer 'para saber si tengo que calcular la cantidad que tiene cada categoria / 0 se calcula y 1 no se calcula
Dim nodoChange As Boolean 'true en el caso de que queden cambios pendientes en el treeview y false en caso contrario

Dim nLegajo As Long
Dim idTpoMenu As Integer
Dim fechaMenu As String

Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
                            
Private Sub cmd_aceptar_Click()

MousePointer = vbHourglass

nLegajo = DataCombo1.BoundText

Frame1.Enabled = True
Frame1.BorderStyle = 1

Frame2.Enabled = False

cmdreporte.Enabled = True
cmd_guardar.Visible = True
cmd_Cancelar.Visible = True
cmd_salir.Visible = True

Picture1.Visible = False
lbl_aviso.Visible = False

'================================
'enalza el control Flexgrid
Dim param(1) As Integer
param(0) = nLegajo

Set Me.MSHFlexGrid1.DataSource = f_StaticRecordset(adCmdTable, "[csl_AdmDietaFlexGrid]", param)
'================================

For j = 1 To 3
    MSHFlexGrid1.ColAlignmentFixed(j) = 3
Next

MSHFlexGrid1.Refresh

Legajo = DataCombo1.BoundText
idTpoMenu = TabStrip2.SelectedItem.Index
fechaMenu = "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"

dbdiet.Execute "insert into Menu_tmp select * from Menu where menu.legajo = " & nLegajo & " and Menu.fechaMenu = " & fechaMenu
'dbdiet.Execute "delete * from Menu where menu.legajo = " & Legajo & " and Menu.fechaMenu = " & fechaMenu
'
dbdiet.Execute "insert into platosMenu_tmp select * from platosMenu where platosMenu.legajo = " & nLegajo & " and platosMenu.fechaMenu = " & fechaMenu
'dbdiet.Execute "delete * from platosMenu where platosMenu.legajo = " & Legajo & " and platosMenu.fechaMenu = " & fechaMenu
       
Call buscar
Call proporcionKcal
Call ActualizaKcalTotal
Call CalculaGrPorGrupo
       
Call f_Boton_Zorder

MousePointer = vbDefault
''''''''''''''''''''''''''''

End Sub

Private Sub cmd_Cancelar_Click()
Dim strMsg As String

If nodoChange Then
    strMsg = MsgBox("¿Esta seguro que desea deshacer los cambios realizados?", vbYesNo)
Else
    strMsg = MsgBox("No se han realizado cambios", vbInformation)
End If

If strMsg = vbYes Then
    MousePointer = vbHourglass
    
'    Frame2.Enabled = True
'
'    Frame1.Enabled = False
'    Frame1.BorderStyle = 0
'
'    cmd_Guardar.Visible = False
'    cmd_cancelar.Visible = False
'
'    Picture1.Visible = True
'    lbl_aviso.Visible = True
'
'    cmd_cerrar.SetFocus
    
    If nodoChange Then
        dbdiet.Execute "delete * from Menu_tmp"
        dbdiet.Execute "delete * from PlatosMenu_tmp"
        
        dbdiet.Execute "insert into Menu_tmp select * from Menu where menu.legajo = " & nLegajo & " and Menu.fechaMenu = " & fechaMenu
        dbdiet.Execute "insert into platosMenu_tmp select * from platosMenu where platosMenu.legajo = " & nLegajo & " and platosMenu.fechaMenu = " & fechaMenu
    
        Call buscar
        Call proporcionKcal
        Call ActualizaKcalTotal
        Call CalculaGrPorGrupo
    
    End If
    
    nodoChange = False
    Me.cmd_guardar.Enabled = False
    Me.cmd_Cancelar.Enabled = False

    Call f_Boton_Zorder
    
    MousePointer = vbDefault
End If

End Sub

Private Sub cmd_cerrar_Click()

Call f_Boton_Zorder

'Unload Me
frm_Adm_Diet.Hide

End Sub

Private Sub cmd_guardar_Click()
Dim strMsg As String

If nodoChange Then

    strMsg = MsgBox("¿Esta seguro que desea guardar los cambios realizados?", vbYesNo)
    
    If strMsg = vbYes Then
        
        MousePointer = vbHourglass
        
        dbdiet.Execute "delete * from Menu where menu.legajo = " & nLegajo & " and Menu.fechaMenu = " & fechaMenu
        dbdiet.Execute "insert into Menu select * from Menu_tmp"
        
        dbdiet.Execute "delete * from platosMenu where platosMenu.legajo = " & nLegajo & " and platosMenu.fechaMenu = " & fechaMenu
        dbdiet.Execute "insert into PlatosMenu select * from PlatosMenu_tmp"
        
        'dbdiet.Execute "delete * from Menu_tmp"
        'dbdiet.Execute "delete * from PlatosMenu_tmp"
        
        Call buscar
        Call proporcionKcal
        Call ActualizaKcalTotal
        Call CalculaGrPorGrupo
        
        nodoChange = False
        Me.cmd_guardar.Enabled = False
        Me.cmd_Cancelar.Enabled = False
        
        Call f_Boton_Zorder
        
        MousePointer = vbDefault
        
    End If

Else

    strMsg = MsgBox("No se han realizado cambios", vbInformation)

End If

End Sub

Private Sub cmd_salir_Click()
Dim strMsg As String

If nodoChange Then
    strMsg = MsgBox("¿Esta seguro que desea finalizar la operacion?" & vbCrLf & vbTab & "- Se perderan los cambios realizados", vbYesNo)
Else
    strMsg = MsgBox("¿Esta seguro que desea finalizar la operacion?", vbYesNo)
End If

If strMsg = vbYes Then
    
    '================================
    'enalza el control Flexgrid
    Dim param(1) As Integer
    param(0) = 0
    
    Set Me.MSHFlexGrid1.DataSource = f_StaticRecordset(adCmdTable, "[csl_AdmDietaFlexGrid]", param)
    '================================
        
    Call f_Cancela
    
    Call f_Boton_Zorder

End If

End Sub



Private Sub CmdReporte_Click()
ReporteMenu.Show
End Sub


Private Sub Command1_Click()
'Text1.Text = TabStrip2.SelectedItem.Index
'Text1.Text = TreeView1.SelectedItem.Key
'TreeView1.Nodes.Remove (Val(Text1.Text))
'MSHFlexGrid1.CellBackColor = &HC0C0C0
'Beep
'Dim ali() As String
'ali = Split(TreeView1.SelectedItem.Key, "//")
'MsgBox ali(0) & " -- " & ali(1) & " -- " & ali(2) & " -- " & ali(3)
'MsgBox "FullPath " & TreeView1.SelectedItem.FullPath & " Index " & TreeView1.SelectedItem.Index & " Key " & TreeView1.SelectedItem.Key & " root " & TreeView1.SelectedItem.Root & " text " & TreeView1.SelectedItem.Text

End Sub

Private Sub cmd_tipito_Click()

frmPacientes.Show
frmPacientes.SetFocus
frmPacientes.Data1.Recordset.FindFirst " legajo = " & DataCombo1.BoundText
End Sub


Private Sub Command3_Click()
Rollback
Call buscar
Call proporcionKcal
Call ActualizaKcalTotal
Call CalculaGrPorGrupo
BeginTrans
End Sub

Private Sub Check_Click()
Dim Alimento() As String
Dim Auxiliar As Integer
Dim cTitulo As String

BeginTrans

ProgressBar1.Visible = True

ProgressBar1.ZOrder 0

ProgressBar1.Max = TreeView1.Nodes.Count

cTitulo = Me.Caption

lbl_Porcentaje.Caption = 0
lbl_Porcentaje.Visible = True
lbl_Porcentaje.ZOrder 0
                
'la variable "Auxiliar" contiene 0 si ya están todos los nodos tildados o de lo contrario un 1.
Auxiliar = 0

For i = 1 To TreeView1.Nodes.Count
    
    Alimento = Split(TreeView1.Nodes(i).Key, "//")
    
    If Alimento(2) = "Ingrediente" Then
            
            'obtención de kcal según la selección de platos en el treeview en tiempo real
            If TreeView1.Nodes(i).Checked = False Then
                
                dbdiet.Execute "insert into platosmenu_tmp (legajo, idtpomenu, idplato, fechaMenu) select " & DataCombo1.BoundText & ", " & TabStrip2.SelectedItem.Index & ", " & Alimento(3) & ", " & "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"
                
                Call DevuelvePorcion(Alimento(3), Alimento(1))
                Call subtotalKcal(Alimento(1), Alimento(3))
                Label1(1) = Val(Label1(1).Caption) + Kcal
            
                'calcula cantidad por categoria para luego comparar con lo que debería tener
                Call DevuelveCategoria(Alimento(1))
                Call CantidadPorcion(Alimento(3))
                MSHFlexGrid1.Recordset.MoveFirst
                
                tb_FlexGrid.MoveFirst
                tb_FlexGrid.Find " idcategoria = " & Categoria1
                aux2 = tb_FlexGrid.AbsolutePosition
                MSHFlexGrid1.TextMatrix(aux2, 3) = MSHFlexGrid1.TextMatrix(aux2, 3) + Porcion * CantAux
                Call pintar(aux2)
            
                '---------------------------------------------------------------------------
                
                Set tb = dbdiet.OpenRecordset("platosmenu_tmp", dbOpenDynaset)
                
                tb.FindFirst " idplato = " & Alimento(3)
                Sumador = tb.Fields("cantneta").Value + Porcion
                tb.Close
            
                TreeView1.Nodes(i).Checked = True
                TreeView1.Nodes(i).Parent.Checked = True
                dbdiet.Execute "insert into menu_tmp (legajo, idtpomenu, idalimento, idplato, fechaMenu) select " & DataCombo1.BoundText & ", " & TabStrip2.SelectedItem.Index & ", " & Alimento(1) & ", " & Alimento(3) & ", " & "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"  '& DTPicker1.Value & "#"
            
                Call DevuelveUnidad(Alimento(3))
                If UnidadPlato = 1 Then ' si es por Unidad
                    dbdiet.Execute "update platosmenu_tmp set CantNeta = 1 where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & Alimento(3) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
                Else
                    dbdiet.Execute "update platosmenu_tmp set CantNeta = " & Sumador & " where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & Alimento(3) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
                End If
                
                Auxiliar = 1
            End If
            '------------------
             
    End If
    ProgressBar1.Value = i
        
    lbl_Porcentaje.Caption = i * 100 / TreeView1.Nodes.Count
    
    Me.Caption = cTitulo & " - " & Format(Val(lbl_Porcentaje.Caption), "standard") & "%"
    
Next
ProgressBar1.Visible = False

lbl_Porcentaje.Visible = False
                
Me.Caption = cTitulo
                
If Auxiliar = 1 Then
    msg = MsgBox("¿Está seguro que quiere guardar los cambios?", vbYesNo, "Umana Nutrición")
    
    If msg = vbYes Then
        CommitTrans
    Else
        Rollback
        Call buscar
        Call proporcionKcal
        Call ActualizaKcalTotal
        Call CalculaGrPorGrupo
    End If
Else
    Beep
    MsgBox "Ya están todos los platos tildados", vbInformation, "Umana Nutrición"
    Rollback
End If

End Sub



Private Sub cmd_Calendario_Click()
Dim strQuery As String

strQuery = "DblClick para recuperar los registros de la fecha"
frm_calendario.cargarParametros 1, DataCombo1.BoundText, strQuery

frm_calendario.Show

Me.Hide

End Sub


Private Sub DataCombo1_Change()

If DataCombo1.Text <> "" Then

    Me.Caption = " Gestion Planes Alimentarios " & " - " & DataCombo1.Text
    
End If

End Sub

Private Sub DataCombo1_LostFocus()
If DataCombo1.Text = "" Then
    
    DataCombo1.SetFocus
    MsgBox "Debe seleccionar un Paciente", vbInformation, "Información"
    
End If

End Sub

Private Sub DTPicker1_Change()
'AuxTab = 1
'
'Call buscar
'Call ActualizaKcalTotal
'AuxTab = 0
'Call CalculaGrPorGrupo


End Sub

Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
'Me.Height = 6525
'Me.Width = 7545 '12000
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

'Image1.Visible = True

'esto soluciona el error de que la propiedad boundtext del datacombo
'toma el valor de la propiedad text del mismo
Dim nBundText As Integer

nBundText = 0

nBundText = DataCombo1.BoundText

Adodc.Refresh

DataCombo1.BoundText = nBundText

End Sub

Private Sub Form_Deactivate()
'msg = MsgBox("¿Desea aplicar los cambios en el 'Administrador de Dietas'?", vbYesNo, "Guardar")
    
'    If msg = vbYes Then
'MsgBox "DeActiv"

End Sub

Private Sub Form_Load()
Dim strQuery As String
Dim Valor() As String
Dim valorProgresBar As Integer

nLegajo = 0

Me.Height = 6525
Me.Width = 7545 '12000

Call f_CargarOrigenDatos

'If DataCombo1.BoundText = "" Then
'    DataCombo1.BoundText = 1
'End If

'configura el mshflexgrid

MSHFlexGrid1.Refresh
''For i = 1 To tb_FlexGrid.RecordCount
''    MSHFlexGrid1.TextMatrix(i, 3) = 0
''Next

MSHFlexGrid1.ColWidth(0) = MSHFlexGrid1.ColWidth(0) / 2
MSHFlexGrid1.ColWidth(1) = 1800 'MSHFlexGrid1.ColWidth(1) * 2
MSHFlexGrid1.ColWidth(2) = 700 'MSHFlexGrid1.ColWidth(2) / 2
MSHFlexGrid1.ColWidth(3) = 700
MSHFlexGrid1.ColAlignment(0) = 1
MSHFlexGrid1.RowHeight(0) = 700
'MSHFlexGrid1.TextMatrix(0, 1) = "Alimentos"
'MSHFlexGrid1.TextMatrix(0, 2) = "Cant. ideal (gr/cc)"
'MSHFlexGrid1.TextMatrix(0, 3) = "Cant. real (gr/cc)"
MSHFlexGrid1.WordWrap = True

'MSHFlexGrid1.Row = 0
For j = 1 To 3
    MSHFlexGrid1.ColAlignmentFixed(j) = 3
Next

For i = MSHFlexGrid1.FixedRows To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.TextArray(MSHFlexGrid1.Cols * i) = i
Next
'-----------------------

'establece la fecha actual
DTPicker1.Value = Now
Label1(1).Caption = 0
Adodc1.Refresh

' se genera la carga del treeview
Adodc1.Refresh

frmSplash.ProgressBar1.Max = Val(frmSplash.ProgressBar1.Value) + Adodc1.Recordset.RecordCount

valorProgresBar = Val(frmSplash.ProgressBar1.Value)

For i = 1 To Adodc1.Recordset.RecordCount
    nombre = Adodc1.Recordset.Fields("nombreplato").Value
    Id = Adodc1.Recordset.Fields("idplato").Value
    
    Set tempNode = TreeView1.Nodes.Add(, , nombre & "//" & Id & "//Plato", nombre)
        
    'strquery = " select * from menu_tmp where Legajo = " & DataCombo1.BoundText & " and IdTpoMenu = " & TabStrip2.SelectedItem.Index & " and idPlato = " & Id & " and fechamenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"    'SSTab1.Tab + 1 & " and idPlato = " & id
    strQuery = " select * from menu where Legajo = " & DataCombo1.BoundText & " and IdTpoMenu = " & TabStrip2.SelectedItem.Index & " and idPlato = " & Id & " and fechamenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"    'SSTab1.Tab + 1 & " and idPlato = " & id
        
    With Adodc3
        .RecordSource = strQuery
        .Refresh
    End With

    If Adodc3.Recordset.RecordCount <> 0 Then
        Adodc3.Recordset.MoveFirst
    End If
    
    strQuery = " select * from ingredientesplatos, alimentos, categoria where ingredientesplatos.codalimento = alimentos.codalimento and alimentos.idcategoria = categoria.idcategoria and ingredientesplatos.idplato = " & Id & " order by decripcion, descripAlimento"
    
    With Adodc2
        .RecordSource = strQuery
        .Refresh
    End With
    
    For j = 1 To Adodc2.Recordset.RecordCount
        
        Descripcion = Adodc2.Recordset.Fields("descripAlimento").Value
        Categoria = Adodc2.Recordset.Fields("decripcion").Value
        cod1 = Adodc2.Recordset.Fields("ingredientesplatos.codalimento").Value
                
        Set tempNode1 = TreeView1.Nodes.Add(nombre & "//" & Id & "//Plato", tvwChild, Descripcion & "//" & cod1 & "//Ingrediente//" & Id, Categoria & ", " & Descripcion)
        'compara el nombre de la categoria y el del alimento, si son iguales NO los repite
        Valor = Split(tempNode1.Text, ", ")
        If Valor(0) = Valor(1) Then
            tempNode1.Text = Valor(0)
        End If
                                
        If Adodc3.Recordset.RecordCount <> 0 Then
            For k = 1 To Adodc3.Recordset.RecordCount
                If Adodc3.Recordset.Fields("idalimento").Value = cod1 Then
                    tempNode.Checked = True
                    tempNode1.Checked = True
                    'obtención de kcal según la selección de platos en el treeview en tiempo real
                    Call CantidadPorcion(Id)
                    Call subtotalKcal(cod1, Id)
                    Label1(1) = Val(Label1(1).Caption) + Kcal * CantAux
                       
                    '------------------
                    'calcula cantidad por categoria para luego comparar con lo que debería tener
                    Call DevuelvePorcion(Id, cod1)
                    Call DevuelveCategoria(cod1)
                    
                    tb_FlexGrid.MoveFirst
                    tb_FlexGrid.Find " idcategoria = " & Categoria1
                    aux2 = tb_FlexGrid.AbsolutePosition
                    MSHFlexGrid1.TextMatrix(aux2, 3) = MSHFlexGrid1.TextMatrix(aux2, 3) + Porcion * CantAux
                    Call pintar(aux2)
                    '---------------------------------------------------------------------------
                End If
                If Adodc3.Recordset.EOF = False Then
                    Adodc3.Recordset.MoveNext
                End If
            Next
            Adodc3.Recordset.MoveFirst
        End If
        
        If Adodc2.Recordset.EOF = False Then
            Adodc2.Recordset.MoveNext
        End If
    Next
       
    If Adodc1.Recordset.EOF = False Then
        Adodc1.Recordset.MoveNext
    End If
    
    frmSplash.ProgressBar1.Value = valorProgresBar + i
Next

' Devuelve o establece un valor que determina si los elementos
' se resaltan cuando el puntero del mouse pasa sobre ellos.
TreeView1.HotTracking = True
TreeView1.Sorted = True


Call proporcionKcal

cmd_cerrar.Cancel = True

AuxTab = 0
nodoChange = False

nLegajo = DataCombo1.BoundText
idTpoMenu = TabStrip2.SelectedItem.Index
fechaMenu = "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"

Call f_Boton_Zorder

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Cancel = True

If cmd_salir.Enabled = True Then

    strMsg = MsgBox("¿Esta seguro que desea finalizar la operacion?" & vbCrLf & vbTab & "- Se perderan los cambios realizados", vbYesNo)

    If strMsg = vbYes Then
    
        Call f_Cancela
        
        Call cmd_cerrar_Click
        
    End If

Else

    Call cmd_cerrar_Click
    
End If

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Me.cmd_Tipito.Enabled = True Then
    Me.Pic_Tipito.ZOrder 0
Else
    Me.Pic_Tipito_Gris.ZOrder 0
End If

End Sub

Private Sub Image1_Click()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 6525
Me.Width = 11265
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2
Image1.Visible = False
End Sub

Private Sub Image2_Click()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 6525
Me.Width = 7545 '12000
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Image1.Visible = True
End Sub

Private Sub Label1_Change(Index As Integer)
If Val(Label1(0).Caption) <> 0 Then
If Val(Label1(1).Caption) > (Val(Label1(0).Caption) - 100) And Val(Label1(1).Caption) < (Val(Label1(0).Caption) + 100) Then
    Label1(1).ForeColor = &HFF&
    Else
    Label1(1).ForeColor = &HFF0000
End If
End If
End Sub



Private Sub noCheck_Click()
Dim Alimento() As String
Dim idPlato As Integer
Dim Auxiliar As Integer
BeginTrans

ProgressBar1.Visible = True
ProgressBar1.Max = TreeView1.Nodes.Count
idPlato = 0
Auxiliar = 0
For i = 1 To TreeView1.Nodes.Count
    
    Alimento = Split(TreeView1.Nodes(i).Key, "//")
           
    If Alimento(2) = "Ingrediente" Then
            
            'obtención de kcal según la selección de platos en el treeview en tiempo real
            If TreeView1.Nodes(i).Checked = True Then
                
                Call CantidadPorcion(Alimento(3))
                Call DevuelvePorcion(Alimento(3), Alimento(1))
                Call subtotalKcal(Alimento(1), Alimento(3))
                Label1(1) = Val(Label1(1).Caption) - (Kcal * CantAux)
            
                'calcula cantidad por categoria para luego comparar con lo que debería tener
                Call DevuelveCategoria(Alimento(1))
                'Call CantidadPorcion(Alimento(3))
                tb_FlexGrid.MoveFirst
                tb_FlexGrid.Find " idcategoria = " & Categoria1
                aux2 = tb_FlexGrid.AbsolutePosition
                MSHFlexGrid1.TextMatrix(aux2, 3) = Val(MSHFlexGrid1.TextMatrix(aux2, 3)) - (Porcion * CantAux)
                Call pintar(aux2)
                              
                TreeView1.Nodes(i).Checked = False
                TreeView1.Nodes(i).Parent.Checked = False
                    
                dbdiet.Execute "delete from menu_tmp where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idalimento = " & Alimento(1) & " and idplato = " & Alimento(3) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"           '& "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"
                
                idPlato = Alimento(3)
                Auxiliar = 1
            End If
            '------------------
    Else
        
        If idPlato <> 0 Then
            dbdiet.Execute "delete from platosmenu_tmp where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & idPlato & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"           '& "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"
            idPlato = 0
        End If
        
    End If
    ProgressBar1.Value = i
    
Next
'para que elimine el último plato
If idPlato <> 0 Then
    dbdiet.Execute "delete from platosmenu_tmp where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & idPlato & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"           '& "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"
    idPlato = 0
End If
'-----------------------
ProgressBar1.Visible = False

If Auxiliar = 1 Then
    msg = MsgBox("¿Está seguro que quiere guardar los cambios?", vbYesNo, "Umana Nutrición")
    
    If msg = vbYes Then
        CommitTrans
    Else
        Rollback
        Call buscar
        Call proporcionKcal
        Call ActualizaKcalTotal
        Call CalculaGrPorGrupo
    End If
Else
    Beep
    MsgBox "Ya están todos los platos destildados", vbInformation, "Umana Nutrición"
    Rollback
End If
    
End Sub







Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Aceptar

End Sub

Private Sub Pic_Calendario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Calendario

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

Private Sub TabStrip2_Click()
MousePointer = vbHourglass
idTpoMenu = TabStrip2.SelectedItem.Index

AuxTab = 1

Call buscar
Call proporcionKcal
Call ActualizaKcalTotal

AuxTab = 0

MousePointer = vbDefault

End Sub

Private Sub Text1_DblClick()
Text1.Text = TreeView1.SelectedItem.Index
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Suma = 0
    TotalPorcion = 0
    CodigoPlato = Split(TreeView1.SelectedItem.Key, "//")
    KeyPlato = TreeView1.SelectedItem.Key
    If TreeView1.SelectedItem.Children <> 0 Then
        
        MDIForm1.ingred.Enabled = True
        MDIForm1.platos.Enabled = True
        
        If TreeView1.SelectedItem.Checked = True Then
            MDIForm1.Cantidad.Enabled = True
                      
            incounter = TreeView1.SelectedItem.Index + 1
            'CodigoPlato = Split(TreeView1(SSTab1.Tab).SelectedItem.Key, "/")
                       
            Call CantidadPorcion(CodigoPlato(1))
            
            For i = 1 To TreeView1.SelectedItem.Children
                If TreeView1.Nodes(incounter).Checked = True Then
                    codigoAlimento = Split(TreeView1.Nodes(incounter).Key, "//")
                    
                    Call subtotalKcal(codigoAlimento(1), CodigoPlato(1))
                    Suma = Suma + Kcal
                                        
                End If
                    incounter = incounter + 1
            Next
                       
        Else
        MDIForm1.Cantidad.Enabled = False
                
        End If
    Else
        MDIForm1.ingred.Enabled = False
        MDIForm1.Cantidad.Enabled = False
        MDIForm1.platos.Enabled = True
    End If
    
    If CodigoPlato(2) = "Plato" Then
        MDIForm1.ingred.Enabled = True
    Else
        MDIForm1.ingred.Enabled = False
    End If

    PopupMenu MDIForm1.EdiCion, vbPopupMenuLeftAlign
End If

End Sub

Sub buscar()
Dim codAlimento() As String
Dim codPlato() As String

Label1(1).Caption = 0
'limpia todas las casillas tildadas
For i = 1 To TreeView1.Nodes.Count
    TreeView1.Nodes(i).Checked = False
Next

'strquery = " select * from menu_tmp where Legajo = " & DataCombo1.BoundText & " and IdTpoMenu = " & TabStrip2.SelectedItem.Index & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
        
'strquery = " select menu.* from menu_tmp, menu where menu_tmp.legajo = menu.legajo and menu_tmp.idTpoMenu = menu.idTpoMenu and menu_tmp.idAlimento = menu.idAlimento and menu_tmp.idPlato = menu.idPlato and menu_tmp.fechaMenu = menu.fechaMenu and menu.Legajo = " & DataCombo1.BoundText & " and menu.IdTpoMenu = " & TabStrip2.SelectedItem.Index & " and menu.fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
'strquery = " select menu.* from menu_tmp right join menu on menu_tmp.legajo = menu.legajo and menu_tmp.idTpoMenu = menu.idTpoMenu and menu_tmp.idAlimento = menu.idAlimento and menu_tmp.idPlato = menu.idPlato and menu_tmp.fechaMenu = menu.fechaMenu where menu.Legajo = " & DataCombo1.BoundText & " and menu.IdTpoMenu = " & TabStrip2.SelectedItem.Index & " and menu.fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"

'carga en un registro todas las ocurrencias de menu_tmp y menu que coincidan con legajo, fecha  y tpo de menu

'strquery = "(select * from menu_tmp where menu_tmp.legajo = " & Legajo & " and menu_tmp.IdTpoMenu = " & idTpoMenu & " and menu_tmp.fechaMenu = " & fechaMenu & ") union (select * from menu where menu.legajo = " & Legajo & " and menu.idTpoMenu = " & idTpoMenu & " and menu.fechaMenu = " & fechaMenu & ")"
strQuery = " select * from menu_tmp where Legajo = " & nLegajo & " and IdTpoMenu = " & idTpoMenu & " and fechaMenu = " & fechaMenu

With Adodc3
    .RecordSource = strQuery
    .Refresh
End With

'si el registro no esta vacío
If Adodc3.Recordset.RecordCount <> 0 Then
    Adodc3.Recordset.MoveFirst
'repetir desde 1 a la cantidad de nodos
For i = 1 To TreeView1.Nodes.Count
    'si el nodo actual no tiene hijos (es decir, si es un alimento) entonces
    If TreeView1.Nodes(i).Children = 0 Then
        'obtiene el cod del plato y del alimento de nodo actual
        codPlato = Split(TreeView1.Nodes(i).Parent.Key, "//")
        codAlimento = Split(TreeView1.Nodes(i).Key, "//")
        'repetir desde 1 a la cantidad de ocurrencias en el registro antes cargado
        For k = 1 To Adodc3.Recordset.RecordCount
            'obtiene el cod del plato y del alimento de la ocurrencia actual en el registro mencionado
            plato = Adodc3.Recordset.Fields("idplato").Value
            Alimento = Adodc3.Recordset.Fields("idalimento").Value
            'si alguno de los codigos del registro coinciden con los del nodo actual entoces
            If codPlato(1) = plato And codAlimento(1) = Alimento Then
                'tildar el nodo actual y su padre
                TreeView1.Nodes(i).Checked = True
                TreeView1.Nodes(i).Parent.Checked = True
            
                'obtención de kcal según la selección de platos en el treeview en tiempo real
                        
                        Call subtotalKcal(codAlimento(1), codPlato(1))
                        Label1(1) = Val(Label1(1).Caption) + Kcal
                    
                '------------------
                
                'If AuxTab = 0 Then
                    'calcula cantidad por categoria para luego comparar con lo que debería tener
                '    Call DevuelvePorcion(codPlato(1), codAlimento(1))
                '    Call DevuelveCategoria(codAlimento(1))
                '    Call CantidadPorcion(codPlato(1))
                '    tb_FlexGrid.MoveFirst
                '    tb_FlexGrid.Find " idcategoria = " & Categoria1
                '    aux2 = tb_FlexGrid.AbsolutePosition
                '    MSHFlexGrid1.TextMatrix(aux2, 3) = MSHFlexGrid1.TextMatrix(aux2, 3) + Porcion * CantAux
                '    Call pintar(aux2)
                    '---------------------------------------------------------------------------
                'End If
            End If
            'avanzar un lugar la ocurrencia sin llegar a EOF
            If Adodc3.Recordset.EOF = False Then
                Adodc3.Recordset.MoveNext
            End If
        Next
        'mover a la primer ocurrencia del registro
        Adodc3.Recordset.MoveFirst
    End If
Next
End If

End Sub
Sub proporcionKcal()
Set tb = dbdiet.OpenRecordset("tpoMenu", dbOpenDynaset)
tb.FindFirst "idtpomenu = " & TabStrip2.SelectedItem.Index
porcentaje = tb.Fields("proporcionRct").Value
tb.Close

Set tb = dbdiet.OpenRecordset("pacientes", dbOpenDynaset)
tb.FindFirst "legajo = " & DataCombo1.BoundText
If tb.Fields("Rctideal").Value <> "" Then
    rctideal = tb.Fields("Rctideal").Value
    Else
    rctideal = 0
End If
tb.Close

Label1(0) = Format(rctideal * porcentaje / 100, "standard")

End Sub

Sub subtotalKcal(codAlimento, plato)
Dim KcalAux, PorcionAux
Set tb = dbdiet.OpenRecordset("alimentos", dbOpenDynaset)
tb.FindFirst " codalimento = " & codAlimento
hc = tb.Fields("hc").Value
prot = tb.Fields("prot").Value
lip = tb.Fields("lip").Value
tb.Close
KcalAux = hc * 4 + prot * 4 + lip * 9

Set tb = dbdiet.OpenRecordset("ingredientesplatos", dbOpenDynaset)
tb.FindFirst " codalimento = " & codAlimento & " and idplato = " & plato
PorcionAux = tb.Fields("porcion").Value
tb.Close

Kcal = PorcionAux * KcalAux / 100

End Sub
Sub CantidadPorcion(plato)
CantAux = 0
Set tb = dbdiet.OpenRecordset("platosmenu_tmp", dbOpenDynaset)
tb.FindFirst " legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & plato & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
If tb.RecordCount <> 0 Then
    CantAux = tb.Fields("Cantidad").Value
End If
tb.Close
End Sub
Sub ActualizaKcalTotal()
Dim codAlimento() As String
Dim codPlato() As String

Label1(1).Caption = 0

For i = 1 To TreeView1.Nodes.Count
    If TreeView1.Nodes(i).Checked = True And TreeView1.Nodes(i).Children <> 0 Then
        codPlato = Split(TreeView1.Nodes(i).Key, "//")
        
        Call CantidadPorcion(codPlato(1))
        
        For j = 1 To TreeView1.Nodes(i).Children
            indice = i + j
            
            If TreeView1.Nodes(indice).Checked = True Then
                codAlimento = Split(TreeView1.Nodes(indice).Key, "//")
                Call subtotalKcal(codAlimento(1), codPlato(1))
                Label1(1) = Val(Label1(1).Caption) + (Kcal * CantAux)
                
            End If
        Next
    End If
Next

End Sub
Sub DevuelveUnidad(plato)
Set tb = dbdiet.OpenRecordset("platos", dbOpenDynaset)
tb.FindFirst " idplato = " & plato
UnidadPlato = tb.Fields("idUnidad").Value
tb.Close
End Sub
Sub DevuelveCantNeta(plato)
Set tb = dbdiet.OpenRecordset("platosmenu_tmp", dbOpenDynaset)
tb.FindFirst " legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & plato & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
CantidadNeta = tb.Fields("cantNeta").Value
tb.Close
End Sub
Sub DevuelvePorcion(plato, Alimento)
Set tb = dbdiet.OpenRecordset("ingredientesplatos", dbOpenDynaset)
tb.FindFirst " idplato = " & plato & " and codalimento = " & Alimento
Porcion = tb.Fields("porcion").Value
tb.Close
End Sub
Sub DevuelveCategoria(Alimento)
Set tb = dbdiet.OpenRecordset("alimentos", dbOpenDynaset)
tb.FindFirst " codalimento = " & Alimento
Categoria1 = tb.Fields("idcategoria").Value
tb.Close
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim kidNode As Node
Dim codAlimento() As String
Dim codPlato() As String

MousePointer = vbHourglass

'tilda todos los hijos de un nodo tildado(para este caso funciona bien pero para algo mas complejo hay que mejorarlo un poco)
incounter = Node.Index + 1
codPlato = Split(Node.Key, "//")
indice = TabStrip2.SelectedItem.Index

If codPlato(2) = "Plato" And Node.Children = 0 Then
    
    MsgBox "El Plato no Tiene Ingredientes para Seleccionar", vbInformation
    
Else
'si el nodo esta chekeado
If Node.Checked = True Then
    
    If Node.Children <> 0 Then
        Sumador = 0
        
        dbdiet.Execute "insert into platosmenu_tmp (legajo, idtpomenu, idplato, fechaMenu) select " & DataCombo1.BoundText & ", " & TabStrip2.SelectedItem.Index & ", " & codPlato(1) & ", " & "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"
                
        For i = 1 To Node.Children
            
            'esta función divide el string según el caracter especial "//" y a cada divición la guarda en un array
            codAlimento = Split(TreeView1.Nodes(incounter).Key, "//")
                                    
            'obtención de kcal según la selección de platos en el treeview en tiempo real
            If TreeView1.Nodes(incounter).Checked = False Then
                Call DevuelvePorcion(codPlato(1), codAlimento(1))
                Call subtotalKcal(codAlimento(1), codPlato(1))
                Label1(1).Caption = Val(Label1(1).Caption) + Kcal
            
            End If
            '------------------
                                  
            Sumador = Sumador + Porcion
            
            TreeView1.Nodes(incounter).Checked = True
                   
            dbdiet.Execute "insert into menu_tmp (legajo, idtpomenu, idalimento, idplato, fechaMenu) select " & DataCombo1.BoundText & ", " & TabStrip2.SelectedItem.Index & ", " & codAlimento(1) & ", " & codPlato(1) & ", " & "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"  '& DTPicker1.Value & "#"
                                       
            incounter = incounter + 1
            
            'calcula cantidad por categoria para luego comparar con lo que debería tener
            Call DevuelveCategoria(codAlimento(1))
            Call CantidadPorcion(codPlato(1))
            tb_FlexGrid.MoveFirst
            tb_FlexGrid.Find " idcategoria = " & Categoria1
            aux2 = tb_FlexGrid.AbsolutePosition
            MSHFlexGrid1.TextMatrix(aux2, 3) = MSHFlexGrid1.TextMatrix(aux2, 3) + Porcion * CantAux
            Call pintar(aux2)
            
            '---------------------------------------------------------------------------
        Next
        
        Call DevuelveUnidad(codPlato(1))
        If UnidadPlato = 1 Then ' si es por Unidad
            dbdiet.Execute "update platosmenu_tmp set CantNeta = 1 where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
        Else
            dbdiet.Execute "update platosmenu_tmp set CantNeta = " & Sumador & " where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
        End If
    Else
        Node.Parent.Checked = True 'tilda el nodo padre
        codPlato = Split(Node.Parent.Key, "//")
        codAlimento = Split(Node.Key, "//")
        dbdiet.Execute "insert into menu_tmp (legajo, idtpomenu, idalimento, idplato, fechaMenu) select " & DataCombo1.BoundText & ", " & TabStrip2.SelectedItem.Index & ", " & codAlimento(1) & ", " & codPlato(1) & ", " & "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'" '& DTPicker1.Value & "#"
        dbdiet.Execute "insert into platosmenu_tmp (legajo, idtpomenu, idplato, fechaMenu) select " & DataCombo1.BoundText & ", " & TabStrip2.SelectedItem.Index & ", " & codPlato(1) & ", " & "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'" '& DTPicker1.Value & "#"
                    
        Call CantidadPorcion(codPlato(1))
        'obtención de kcal según la selección de platos en el treeview en tiempo real
        
        Call DevuelvePorcion(codPlato(1), codAlimento(1))
        Call subtotalKcal(codAlimento(1), codPlato(1))
        Label1(1).Caption = Val(Label1(1).Caption) + (Kcal * CantAux)
        
        Call DevuelveUnidad(codPlato(1))
        Call DevuelveCantNeta(codPlato(1))
        If UnidadPlato = 1 Then ' si es por Unidad
            dbdiet.Execute "update platosmenu_tmp set CantNeta = " & CantAux & " where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
        Else
            dbdiet.Execute "update platosmenu_tmp set CantNeta = " & CantidadNeta + (Porcion * CantAux) & " where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
        End If
        '------------------
        '------------------
            'calcula cantidad por categoria para luego comparar con lo que debería tener
            Call DevuelveCategoria(codAlimento(1))
            'Call CantidadPorcion(codPlato(1))
            tb_FlexGrid.MoveFirst
            tb_FlexGrid.Find " idcategoria = " & Categoria1
            aux2 = tb_FlexGrid.AbsolutePosition
            MSHFlexGrid1.TextMatrix(aux2, 3) = MSHFlexGrid1.TextMatrix(aux2, 3) + Porcion * CantAux
            Call pintar(aux2)
            '---------------------------------------------------------------------------
    End If
    
Else
    If Node.Children <> 0 Then
                
        Call CantidadPorcion(codPlato(1))
        
        For i = 1 To Node.Children
                
            'esta función divide el string según el caracter especial "/" y a cada divición la guarda en un array
            codAlimento = Split(TreeView1.Nodes(incounter).Key, "//")
            
            'obtención de kcal según la selección de platos en el treeview en tiempo real
            If TreeView1.Nodes(incounter).Checked = True Then
                Call subtotalKcal(codAlimento(1), codPlato(1))
                Label1(1) = Val(Label1(1).Caption) - (Kcal * CantAux)
            
                'calcula cantidad por categoria para luego comparar con lo que debería tener
                Call DevuelvePorcion(codPlato(1), codAlimento(1))
                Call DevuelveCategoria(codAlimento(1))
                'Call CantidadPorcion(codPlato(1))
                tb_FlexGrid.MoveFirst
                tb_FlexGrid.Find " idcategoria = " & Categoria1
                aux2 = tb_FlexGrid.AbsolutePosition
                MSHFlexGrid1.TextMatrix(aux2, 3) = MSHFlexGrid1.TextMatrix(aux2, 3) - (Porcion * CantAux)
                Call pintar(aux2)
                '---------------------------------------------------------------------------
                
            End If
            '------------------
            '------------------
                        
            TreeView1.Nodes(incounter).Checked = False
                    
            dbdiet.Execute "delete from menu_tmp where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idalimento = " & codAlimento(1) & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"           '& "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"
            dbdiet.Execute "delete from platosmenu_tmp where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"           '& "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"
            incounter = incounter + 1
            
            
        Next
    Else
        Set kidNode = Node.Parent.Child
        Counter = kidNode.FirstSibling.Index
        bandera = 0
        For k = 1 To Node.Parent.Children
            If TreeView1.Nodes(Counter).Checked = True Then
                bandera = 1
            End If
            Counter = Counter + 1
        Next
        If bandera = 0 Then
            Node.Parent.Checked = False
        End If
        
        codPlato = Split(Node.Parent.Key, "//")
        codAlimento = Split(Node.Key, "//")
                
        Call CantidadPorcion(codPlato(1))
        
        dbdiet.Execute "delete from menu_tmp where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idalimento = " & codAlimento(1) & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"      '& "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"
                
        If bandera = 0 Then
            Call CantidadPorcion(codPlato(1))
            dbdiet.Execute "delete from platosmenu_tmp where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"      '& "'" & Format(DTPicker1.Value, "dd/mm/yy") & "'"
        End If
                     
        'obtención de kcal según la selección de platos en el treeview en tiempo real
        Call DevuelvePorcion(codPlato(1), codAlimento(1))
        Call subtotalKcal(codAlimento(1), codPlato(1))
        Label1(1) = Val(Label1(1).Caption) - (Kcal * CantAux)
                     
        '------------------
        'calcula cantidad por categoria para luego comparar con lo que debería tener
        Call DevuelveCategoria(codAlimento(1))
        'Call CantidadPorcion(codPlato(1))
        tb_FlexGrid.MoveFirst
        tb_FlexGrid.Find " idcategoria = " & Categoria1
        aux2 = tb_FlexGrid.AbsolutePosition
        MSHFlexGrid1.TextMatrix(aux2, 3) = MSHFlexGrid1.TextMatrix(aux2, 3) - (Porcion * CantAux)
        Call pintar(aux2)
            '---------------------------------------------------------------------------
            
        If bandera = 1 Then
            Call DevuelveUnidad(codPlato(1))
            Call DevuelveCantNeta(codPlato(1))
            If UnidadPlato = 1 Then ' si es por Unidad
                dbdiet.Execute "update platosmenu_tmp set CantNeta = " & CantAux & " where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
            Else
                dbdiet.Execute "update platosmenu_tmp set CantNeta = " & CantidadNeta - (Porcion * CantAux) & " where legajo = " & DataCombo1.BoundText & " and idtpomenu = " & TabStrip2.SelectedItem.Index & " and idplato = " & codPlato(1) & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
            End If
            
        End If
        '------------------
    End If
 
End If

End If

nodoChange = True
Me.cmd_guardar.Enabled = True
Me.cmd_Cancelar.Enabled = True

Call f_Boton_Zorder

MousePointer = vbDefault

End Sub
Sub CantidadPopup() 'calcula cantidad por categoria luego de modificar la cantidad de un plato
If TreeView1.SelectedItem.Children <> 0 Then
                        
        If TreeView1.SelectedItem.Checked = True Then
            incounter = TreeView1.SelectedItem.Index + 1
            'CodigoPlato = Split(TreeView1(SSTab1.Tab).SelectedItem.Key, "/")
            CodigoPlato = Split(TreeView1.SelectedItem.Key, "//")
'            Call CantidadPorcion(CodigoPlato(1))
            
            
            For i = 1 To TreeView1.SelectedItem.Children
                If TreeView1.Nodes(incounter).Checked = True Then
                    codigoAlimento = Split(TreeView1.Nodes(incounter).Key, "//")
                    
                    'Call subtotalKcal(codigoAlimento(1), CodigoPlato(1))
                    'Suma = Suma + Kcal
                    'calcula cantidad por categoria para luego comparar con lo que debería tener
                    Call DevuelvePorcion(CodigoPlato(1), codigoAlimento(1))
                    Call DevuelveCategoria(codigoAlimento(1))
                    Call CantidadPorcion(CodigoPlato(1))
                    tb_FlexGrid.MoveFirst
                    tb_FlexGrid.Find " idcategoria = " & Categoria1
                    aux2 = tb_FlexGrid.AbsolutePosition
                    MSHFlexGrid1.TextMatrix(aux2, 3) = MSHFlexGrid1.TextMatrix(aux2, 3) + Porcion * (CantAux - PrevCantAux) '(CantAux - PrevCantAux)es la cantidad actual - la cantidad anterior
                    
                    Call pintar(aux2)
                
                    '---------------------------------------------------------------------------
                    
                End If
                    incounter = incounter + 1
            Next
                
        End If
    
    End If

End Sub


Sub pintar(posicion)
MSHFlexGrid1.Row = posicion
                                               
If MSHFlexGrid1.TextMatrix(posicion, 2) <> 0 Then
    If MSHFlexGrid1.TextMatrix(posicion, 3) > (MSHFlexGrid1.TextMatrix(posicion, 2) - 6) And MSHFlexGrid1.TextMatrix(posicion, 3) < (MSHFlexGrid1.TextMatrix(posicion, 2) + 6) Then
        MSHFlexGrid1.CellBackColor = &H8000000F  '&HC0C0C0
    Else
        MSHFlexGrid1.CellBackColor = &H80000014
    End If
End If
End Sub

Sub CalculaGrPorGrupo()

Dim codAlimento() As String
Dim codPlato() As String

'Label1(1).Caption = 0
'limpia todas las casillas tildadas
'For i = 1 To TreeView1.Nodes.Count
'    TreeView1.Nodes(i).Checked = False
'Next

' setea la tercer columna del mshflexgrid
For i = 1 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.TextMatrix(i, 3) = 0
Next
'por cada ficha del TabStrip2
For h = 1 To 6
'carga en un registro todas las ocurrencias de menu_tmp que coincidan con legajo, fecha  y tpo de menu

    strQuery = " select * from menu_tmp where Legajo = " & DataCombo1.BoundText & " and IdTpoMenu = " & h & " and fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
        
    'strquery = " select menu.* from menu_tmp, menu where menu_tmp.legajo = menu.legajo and menu_tmp.idTpoMenu = menu.idTpoMenu and menu_tmp.idAlimento = menu.idAlimento and menu_tmp.idPlato = menu.idPlato and menu_tmp.fechaMenu = menu.fechaMenu and menu.Legajo = " & DataCombo1.BoundText & " and menu.IdTpoMenu = " & h & " and menu.fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
    'strquery = " select menu.* from menu_tmp right join menu on menu_tmp.legajo = menu.legajo and menu_tmp.idTpoMenu = menu.idTpoMenu and menu_tmp.idAlimento = menu.idAlimento and menu_tmp.idPlato = menu.idPlato and menu_tmp.fechaMenu = menu.fechaMenu where menu.Legajo = " & DataCombo1.BoundText & " and menu.IdTpoMenu = " & h & " and menu.fechaMenu = " & "#" & Month(DTPicker1.Value) & "/" & Day(DTPicker1.Value) & "/" & Year(DTPicker1.Value) & "#"
    
    With Adodc3
        .RecordSource = strQuery
        .Refresh
    End With
    
    'si el registro no esta vacío
    If Adodc3.Recordset.RecordCount <> 0 Then
        Adodc3.Recordset.MoveFirst
        'repetir desde 1 a la cantidad de nodos
        For i = 1 To TreeView1.Nodes.Count
            'si el nodo actual no tiene hijos (es decir, si es un alimento) entonces
            If TreeView1.Nodes(i).Children = 0 Then
                'obtiene el cod del plato y del alimento de nodo actual
                codPlato = Split(TreeView1.Nodes(i).Parent.Key, "//")
                codAlimento = Split(TreeView1.Nodes(i).Key, "//")
                'repetir desde 1 a la cantidad de ocurrencias en el registro antes cargado
                For k = 1 To Adodc3.Recordset.RecordCount
                    'obtiene el cod del plato y del alimento de la ocurrencia actual en el registro mencionado
                    plato = Adodc3.Recordset.Fields("idplato").Value
                    Alimento = Adodc3.Recordset.Fields("idalimento").Value
                    'si alguno de los codigos del registro coinciden con los del nodo actual entoces
                    If codPlato(1) = plato And codAlimento(1) = Alimento Then
                        'tildar el nodo actual y su padre
                        'TreeView1.Nodes(i).Checked = True
                        'TreeView1.Nodes(i).Parent.Checked = True
                    
                        'obtención de kcal según la selección de platos en el treeview en tiempo real
                                
                                Call subtotalKcal(codAlimento(1), codPlato(1))
                                'Label1(1) = Val(Label1(1).Caption) + Kcal
                            
                        '------------------
                        
                        If AuxTab = 0 Then
                            'calcula cantidad por categoria para luego comparar con lo que debería tener
                            Call DevuelvePorcion(codPlato(1), codAlimento(1))
                            Call DevuelveCategoria(codAlimento(1))
                            Call CantidadPorcion(codPlato(1))
                            tb_FlexGrid.MoveFirst
                            tb_FlexGrid.Find " idcategoria = " & Categoria1
                            aux2 = tb_FlexGrid.AbsolutePosition
                            MSHFlexGrid1.TextMatrix(aux2, 3) = MSHFlexGrid1.TextMatrix(aux2, 3) + Porcion * CantAux
                            Call pintar(aux2)
                            '---------------------------------------------------------------------------
                        End If
                    End If
                    'avanzar un lugar la ocurrencia sin llegar a EOF
                    If Adodc3.Recordset.EOF = False Then
                        Adodc3.Recordset.MoveNext
                    End If
                Next
                'mover a la primer ocurrencia del registro
                Adodc3.Recordset.MoveFirst
            End If
        Next
    End If

Next
End Sub


Sub f_Cancela()

MousePointer = vbHourglass
    
Frame2.Enabled = True

Frame1.Enabled = False
Frame1.BorderStyle = 0

cmdreporte.Enabled = False
cmd_guardar.Visible = False
cmd_Cancelar.Visible = False
cmd_salir.Visible = False

Picture1.Visible = True
lbl_aviso.Visible = True

cmd_cerrar.SetFocus

If nodoChange Then
    dbdiet.Execute "delete * from Menu_tmp"
    dbdiet.Execute "delete * from PlatosMenu_tmp"

    Call buscar
    Call proporcionKcal
    Call ActualizaKcalTotal
    Call CalculaGrPorGrupo

End If

nodoChange = False

MousePointer = vbDefault
    
End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Data1.Recordset = Nothing
Set Me.Adodc.Recordset = Nothing
Set Me.Adodc1.Recordset = Nothing
Set Me.Adodc2.Recordset = Nothing
Set Me.Adodc3.Recordset = Nothing

strQuery = "SELECT * FROM Pacientes"
Call f_Data_DatabaseName(Data1, strQuery)

strQuery = "select *, (apell & ', ' & nombre) as nom from pacientes order by apell, nombre"
Call f_Adodc_ConnectionString(Adodc, strQuery)

strQuery = "Platos"
Call f_Adodc_ConnectionString(Adodc1, strQuery)

strQuery = "select * from alimentos, ingredientesplatos where alimentos.codalimento = ingredientesplatos.codalimento"
Call f_Adodc_ConnectionString(Adodc2, strQuery)

strQuery = "select * from Menu"
Call f_Adodc_ConnectionString(Adodc3, strQuery)

'=========================================
'Enlaza control Flexgrid
Dim param(1) As Integer
param(0) = nLegajo

Set tb_FlexGrid = f_StaticRecordset(adCmdTable, "[csl_AdmDietaFlexGrid]", param)
Set Me.MSHFlexGrid1.DataSource = tb_FlexGrid 'f_StaticRecordset(adCmdTable, "[csl_AdmDietaFlexGrid]", param)
'=========================================

''strQuery = "csl_AdmDietaFlexGrid"
''Call f_Adodc_ConnectionString(Adodc6, strQuery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo1, Data1, Adodc, "Legajo", "Legajo", "nom")

'==============================================

End Sub

Sub f_Boton_Zorder()

If Me.cmd_Calendario.Enabled = True Then
    Me.Pic_Calendario.ZOrder 0
Else
    Me.Pic_Calendario_Gris.ZOrder 0
End If

If Me.cmd_Tipito.Enabled = True Then
    Me.Pic_Tipito.ZOrder 0
Else
    Me.Pic_Tipito_Gris.ZOrder 0
End If

If Me.cmdreporte.Enabled = True Then
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


Me.cmd_aceptar.ZOrder 1
Me.cmd_cerrar.ZOrder 1

Me.cmdreporte.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
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

Me.cmdreporte.ZOrder 1
Me.cmd_guardar.ZOrder 0
Me.cmd_Cancelar.ZOrder 1
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Cancelar()

Me.cmdreporte.ZOrder 1
Me.cmd_guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 0
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Salir()

Me.cmdreporte.ZOrder 1
Me.cmd_guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_salir.ZOrder 0

End Sub

Sub f_Imprimir()

Me.cmdreporte.ZOrder 0
Me.cmd_guardar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1
Me.cmd_salir.ZOrder 1

End Sub

Sub f_Tipito()

Me.cmd_Tipito.ZOrder 0

End Sub

Sub f_Calendario()

Me.cmd_Calendario.ZOrder 0

End Sub


