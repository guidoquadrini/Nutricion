VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_turnos 
   Caption         =   "Excepciones de horarios de Atencion"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   LinkTopic       =   "Form2"
   ScaleHeight     =   6135
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4575
      Left            =   5280
      TabIndex        =   0
      Top             =   1320
      Width           =   6255
      Begin VB.Frame fme_botones_abm 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   285
         TabIndex        =   12
         Top             =   3840
         Width           =   5535
         Begin VB.CommandButton cmdAceptar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":0000
            Height          =   375
            Left            =   4680
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":0159
            Picture         =   "frm_turnos.frx":02AB
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Aceptar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdCancelar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":0567
            Height          =   375
            Left            =   5160
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":06FB
            Picture         =   "frm_turnos.frx":084D
            Style           =   1  'Graphical
            TabIndex        =   44
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
            MouseIcon       =   "frm_turnos.frx":0D00
            Picture         =   "frm_turnos.frx":0E52
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   43
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
            MouseIcon       =   "frm_turnos.frx":1153
            Picture         =   "frm_turnos.frx":12A5
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   42
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdModificar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":1561
            Height          =   375
            Left            =   4080
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":1682
            Picture         =   "frm_turnos.frx":17D4
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Modificar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdBorrar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":1A47
            Height          =   375
            Left            =   3600
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":1B5D
            Picture         =   "frm_turnos.frx":1CAF
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Eliminar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdAgregar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":1E3E
            Height          =   375
            Left            =   3120
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":1F8B
            Picture         =   "frm_turnos.frx":20DD
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Agregar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdUltimo 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":2517
            Height          =   375
            Left            =   2520
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":26BF
            Picture         =   "frm_turnos.frx":2811
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Ultimo"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdSiguiente 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":2CDC
            Height          =   375
            Left            =   2040
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":2E49
            Picture         =   "frm_turnos.frx":2F9B
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Siguiente"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdBuscar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":3410
            Height          =   375
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":3598
            Picture         =   "frm_turnos.frx":36EA
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Buscar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdAnterior 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":39C7
            Height          =   375
            Left            =   1080
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":3B31
            Picture         =   "frm_turnos.frx":3C83
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Anterior"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdPrimero 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":40F1
            Height          =   375
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":4296
            Picture         =   "frm_turnos.frx":43E8
            Style           =   1  'Graphical
            TabIndex        =   34
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
            MouseIcon       =   "frm_turnos.frx":48A3
            Picture         =   "frm_turnos.frx":49F5
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   33
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
            MouseIcon       =   "frm_turnos.frx":4EB0
            Picture         =   "frm_turnos.frx":5002
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   32
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
            MouseIcon       =   "frm_turnos.frx":5470
            Picture         =   "frm_turnos.frx":55C2
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   31
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
            MouseIcon       =   "frm_turnos.frx":589F
            Picture         =   "frm_turnos.frx":59F1
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   30
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
            MouseIcon       =   "frm_turnos.frx":5E66
            Picture         =   "frm_turnos.frx":5FB8
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   29
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
            MouseIcon       =   "frm_turnos.frx":6483
            Picture         =   "frm_turnos.frx":65D5
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   28
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
            MouseIcon       =   "frm_turnos.frx":6A0F
            Picture         =   "frm_turnos.frx":6B61
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   27
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
            MouseIcon       =   "frm_turnos.frx":6CF0
            Picture         =   "frm_turnos.frx":6E42
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   26
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
            MouseIcon       =   "frm_turnos.frx":70B5
            Picture         =   "frm_turnos.frx":7207
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
            MouseIcon       =   "frm_turnos.frx":7328
            Picture         =   "frm_turnos.frx":747A
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   24
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
            MouseIcon       =   "frm_turnos.frx":7590
            Picture         =   "frm_turnos.frx":76E2
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   23
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
            MouseIcon       =   "frm_turnos.frx":782F
            Picture         =   "frm_turnos.frx":7981
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   22
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
            MouseIcon       =   "frm_turnos.frx":7B29
            Picture         =   "frm_turnos.frx":7C7B
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   21
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
            MouseIcon       =   "frm_turnos.frx":7DE8
            Picture         =   "frm_turnos.frx":7F3A
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   20
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
            MouseIcon       =   "frm_turnos.frx":80C2
            Picture         =   "frm_turnos.frx":8214
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   19
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
            MouseIcon       =   "frm_turnos.frx":837E
            Picture         =   "frm_turnos.frx":84D0
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   18
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
            MouseIcon       =   "frm_turnos.frx":8675
            Picture         =   "frm_turnos.frx":87C7
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   17
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
            MouseIcon       =   "frm_turnos.frx":8920
            Picture         =   "frm_turnos.frx":8A72
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   16
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdImprimir 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":8C06
            Height          =   375
            Left            =   0
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_turnos.frx":8D5E
            Style           =   1  'Graphical
            TabIndex        =   15
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
            MouseIcon       =   "frm_turnos.frx":91DE
            Picture         =   "frm_turnos.frx":9330
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   14
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
            MouseIcon       =   "frm_turnos.frx":9488
            Picture         =   "frm_turnos.frx":95DA
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   13
            Top             =   120
            Width           =   375
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   0
            Top             =   480
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
      End
      Begin VB.Frame fme_SelectProf 
         Caption         =   "Seleccionar Profesional"
         Height          =   735
         Left            =   45
         TabIndex        =   5
         Top             =   0
         Width           =   6135
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   495
            Left            =   3360
            TabIndex        =   6
            Top             =   120
            Width           =   615
            Begin VB.PictureBox Pic_Tipito_Gris 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               MouseIcon       =   "frm_turnos.frx":9A5A
               Picture         =   "frm_turnos.frx":9BAC
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   9
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
               MouseIcon       =   "frm_turnos.frx":9CDC
               Picture         =   "frm_turnos.frx":9E2E
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   8
               Top             =   120
               Width           =   315
            End
            Begin VB.CommandButton cmd_Tipito 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_turnos.frx":A0BE
               Height          =   315
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frm_turnos.frx":A7CE
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Agregar"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   315
            End
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frm_turnos.frx":AA5E
            DataField       =   "ehr_idProf"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1560
            TabIndex        =   10
            Top             =   240
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
         Begin VB.Label Label4 
            Caption         =   "Profesional:"
            DragMode        =   1  'Automatic
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   270
            Width           =   855
         End
      End
      Begin VB.Frame fme_ActualHr 
         Caption         =   "Actualizar Horarios"
         Height          =   3015
         Left            =   45
         TabIndex        =   1
         Top             =   720
         Width           =   6135
         Begin MSComCtl2.DTPicker DTP_Fecha 
            DataField       =   "ehr_fecha"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   960
            TabIndex        =   46
            Top             =   1080
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   66191361
            CurrentDate     =   37867
         End
         Begin MSComCtl2.DTPicker DTP_hrHasta 
            Height          =   285
            Left            =   3720
            TabIndex        =   2
            Top             =   1080
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
            Left            =   2640
            TabIndex        =   3
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Format          =   66191362
            CurrentDate     =   38495
         End
         Begin MSDataGridLib.DataGrid Datagrid1 
            Bindings        =   "frm_turnos.frx":AA73
            Height          =   1695
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   2990
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
            Left            =   4080
            Top             =   2520
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
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
   End
   Begin VB.Frame Frame6 
      Caption         =   "Actualizar Horarios"
      Height          =   1455
      Left            =   240
      TabIndex        =   55
      Top             =   1320
      Width           =   4695
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         ItemData        =   "frm_turnos.frx":AA8E
         Left            =   240
         List            =   "frm_turnos.frx":AA98
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   840
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "hrs_hrHasta"
         DataSource      =   "DataDia(0)"
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   57
         Top             =   840
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
         DataSource      =   "DataDia(0)"
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   58
         Top             =   840
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
         DataSource      =   "DataDia(0)"
         Height          =   285
         Index           =   21
         Left            =   3720
         TabIndex        =   59
         Top             =   840
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
         DataSource      =   "DataDia(0)"
         Height          =   285
         Index           =   14
         Left            =   3000
         TabIndex        =   60
         Top             =   840
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
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Index           =   13
         Left            =   3720
         TabIndex        =   64
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   14
         Left            =   3000
         TabIndex        =   65
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Index           =   9
         Left            =   2280
         TabIndex        =   66
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   67
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Primer Turno"
         Height          =   195
         Left            =   1695
         TabIndex        =   63
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Segundo Turno"
         Height          =   195
         Left            =   3030
         TabIndex        =   62
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Turnos"
         Height          =   195
         Index           =   15
         Left            =   360
         TabIndex        =   61
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   8175
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   1440
         TabIndex        =   69
         Top             =   2880
         Width           =   2175
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   720
            MouseIcon       =   "frm_turnos.frx":AAAF
            Picture         =   "frm_turnos.frx":AC01
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   81
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            MouseIcon       =   "frm_turnos.frx":AD22
            Picture         =   "frm_turnos.frx":AE74
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   80
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1200
            MouseIcon       =   "frm_turnos.frx":B008
            Picture         =   "frm_turnos.frx":B15A
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   79
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            MouseIcon       =   "frm_turnos.frx":B2B3
            Picture         =   "frm_turnos.frx":B405
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   78
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1680
            MouseIcon       =   "frm_turnos.frx":B885
            Picture         =   "frm_turnos.frx":B9D7
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   77
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1200
            MouseIcon       =   "frm_turnos.frx":BCD8
            Picture         =   "frm_turnos.frx":BE2A
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   76
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   720
            MouseIcon       =   "frm_turnos.frx":C0E6
            Picture         =   "frm_turnos.frx":C238
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   75
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            MouseIcon       =   "frm_turnos.frx":C4AB
            Picture         =   "frm_turnos.frx":C5FD
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   74
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":C755
            Height          =   375
            Left            =   720
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":C876
            Picture         =   "frm_turnos.frx":C9C8
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Modificar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":CC3B
            Height          =   375
            Left            =   1200
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":CD94
            Picture         =   "frm_turnos.frx":CEE6
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Aceptar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton Command4 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":D1A2
            Height          =   375
            Left            =   1680
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_turnos.frx":D336
            Picture         =   "frm_turnos.frx":D488
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Cancelar"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton Command5 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_turnos.frx":D93B
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frm_turnos.frx":DA93
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Imprimir"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Seleccionar Profesional"
         Height          =   1095
         Left            =   240
         TabIndex        =   48
         Top             =   120
         Width           =   4695
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "ehr_fecha"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   1560
            TabIndex        =   68
            Top             =   600
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Format          =   66191361
            CurrentDate     =   37867
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   495
            Left            =   3360
            TabIndex        =   49
            Top             =   120
            Width           =   615
            Begin VB.CommandButton Command1 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_turnos.frx":DF13
               Height          =   315
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frm_turnos.frx":E623
               Style           =   1  'Graphical
               TabIndex        =   52
               ToolTipText     =   "Agregar"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   315
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               MouseIcon       =   "frm_turnos.frx":E8B3
               Picture         =   "frm_turnos.frx":EA05
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   51
               Top             =   120
               Width           =   315
            End
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               MouseIcon       =   "frm_turnos.frx":EC95
               Picture         =   "frm_turnos.frx":EDE7
               ScaleHeight     =   315
               ScaleWidth      =   315
               TabIndex        =   50
               Top             =   120
               Width           =   315
            End
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "frm_turnos.frx":EF17
            DataField       =   "ehr_idProf"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1560
            TabIndex        =   53
            Top             =   240
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
         Begin VB.Label Label6 
            Caption         =   "Fecha:"
            DragMode        =   1  'Automatic
            Height          =   255
            Left            =   360
            TabIndex        =   82
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Profesional:"
            DragMode        =   1  'Automatic
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   270
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frm_turnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Tipito_Click()

Unload frm_abm_prof
frm_abm_prof.Show
frm_abm_prof.Data1.Recordset.FindFirst " prf_codigo = " & DataCombo1.BoundText

End Sub

Private Sub Datagrid1_ButtonClick(ByVal ColIndex As Integer)
Select Case ColIndex
    Case Is = 1
    
            With Me.DTP_Fecha
                
                .Left = Datagrid1.Columns(ColIndex).Left + Datagrid1.Left
                .Width = Datagrid1.Columns(ColIndex).Width '+ 50
                        
                .Top = Datagrid1.Top + Datagrid1.RowTop(Datagrid1.Row) + Datagrid1.RowHeight
                
                'en caso de error continuo con la siguiente intruccion
                'ya que cuando estoy agregando un registro la siguiente
                'intruccion provoca un error
                On Error Resume Next
                .Value = Datagrid1.Columns("Fecha").Value
                
                .Visible = True
                .SetFocus
            End With
    
    Case Is = 2
            
            With Me.DTP_hrDesde
            
                .Left = Datagrid1.Columns(ColIndex).Left + Datagrid1.Left
                .Width = Datagrid1.Columns(ColIndex).Width
            
                .Top = Datagrid1.Top + Datagrid1.RowTop(Datagrid1.Row) + Datagrid1.RowHeight
                                                
                'en caso de error continuo con la siguiente intruccion
                'ya que cuando estoy agregando un registro la siguiente
                'intruccion provoca un error
                On Error Resume Next
                .Value = Datagrid1.Columns("hr. Desde").Value
                                
                .Visible = True
                .SetFocus
            End With
    
    Case Is = 3
            
            With Me.DTP_hrHasta
            
                .Left = Datagrid1.Columns(ColIndex).Left + Datagrid1.Left
                .Width = Datagrid1.Columns(ColIndex).Width '+ 50
            
                .Top = Datagrid1.Top + Datagrid1.RowTop(Datagrid1.Row) + Datagrid1.RowHeight
                                                
                'en caso de error continuo con la siguiente intruccion
                'ya que cuando estoy agregando un registro la siguiente
                'intruccion provoca un error
                On Error Resume Next
                .Value = Datagrid1.Columns("hr. Hasta").Value
                                
                .Visible = True
                .SetFocus
            End With

End Select
End Sub

Private Sub Datagrid1_Click()

Me.DTP_Fecha.Visible = False
Me.DTP_hrDesde.Visible = False
Me.DTP_hrHasta.Visible = False

End Sub

Private Sub DTP_Fecha_LostFocus()

'Fecha
Datagrid1.Columns("Fecha").Value = Me.DTP_Fecha.Value

End Sub

Private Sub DTP_hrDesde_LostFocus()

'Hr. Desde
Datagrid1.Columns("hr. Desde").Value = Me.DTP_hrDesde.Value

End Sub

Private Sub DTP_hrHasta_LostFocus()

'Hr. Hasta
Datagrid1.Columns("hr. Hasta").Value = Me.DTP_hrHasta.Value

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
