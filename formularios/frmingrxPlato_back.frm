VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmingrxPlato_back 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingredientes por Platos"
   ClientHeight    =   3240
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frmingrxPlato_back.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6180
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   -240
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6600
      Top             =   2160
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
      RecordSource    =   $"frmingrxPlato_back.frx":0ECA
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
   Begin VB.Frame Frame2 
      Caption         =   "Nro."
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Label Label1 
         Caption         =   "label1"
         DataField       =   "idPlato"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalles"
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtFields 
         DataField       =   "Porcion"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   2
         Top             =   1725
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ac&tualizar"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         ToolTipText     =   "Mostrar Todos"
         Top             =   2160
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmingrxPlato_back.frx":1026
         DataField       =   "idPlato"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmingrxPlato_back.frx":103B
         DataField       =   "CodAlimento"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "nom"
         BoundColumn     =   "CodAlimento"
         Text            =   "DataCombo2"
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
      Begin VB.Label lblLabels 
         Caption         =   "Plato:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   495
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ingrediente:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porción:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1740
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6000
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad en grs. o cc _________________________"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   3855
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6600
      Top             =   1680
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
      RecordSource    =   "select * from platos order by nombreplato"
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
   Begin VB.Frame fme_botones_abm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   323
      TabIndex        =   11
      Top             =   2640
      Width           =   5535
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":1050
         Height          =   375
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":11A9
         Picture         =   "frmingrxPlato_back.frx":12FB
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Aceptar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":15B7
         Height          =   375
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":174B
         Picture         =   "frmingrxPlato_back.frx":189D
         Style           =   1  'Graphical
         TabIndex        =   38
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
         MouseIcon       =   "frmingrxPlato_back.frx":1D50
         Picture         =   "frmingrxPlato_back.frx":1EA2
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   37
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
         MouseIcon       =   "frmingrxPlato_back.frx":21A3
         Picture         =   "frmingrxPlato_back.frx":22F5
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   36
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":25B1
         Height          =   375
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":26D2
         Picture         =   "frmingrxPlato_back.frx":2824
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Modificar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":2A97
         Height          =   375
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":2BAD
         Picture         =   "frmingrxPlato_back.frx":2CFF
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Eliminar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":2E8E
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":2FDB
         Picture         =   "frmingrxPlato_back.frx":312D
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Agregar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUltimo 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":3567
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":370F
         Picture         =   "frmingrxPlato_back.frx":3861
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ultimo"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSiguiente 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":3D2C
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":3E99
         Picture         =   "frmingrxPlato_back.frx":3FEB
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Siguiente"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":4460
         Height          =   375
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":45E8
         Picture         =   "frmingrxPlato_back.frx":473A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Buscar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":4A17
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":4B81
         Picture         =   "frmingrxPlato_back.frx":4CD3
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Anterior"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimero 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":5141
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "frmingrxPlato_back.frx":52E6
         Picture         =   "frmingrxPlato_back.frx":5438
         Style           =   1  'Graphical
         TabIndex        =   20
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
         MouseIcon       =   "frmingrxPlato_back.frx":58F3
         Picture         =   "frmingrxPlato_back.frx":5A45
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   19
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
         MouseIcon       =   "frmingrxPlato_back.frx":5F00
         Picture         =   "frmingrxPlato_back.frx":6052
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   18
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
         MouseIcon       =   "frmingrxPlato_back.frx":64C0
         Picture         =   "frmingrxPlato_back.frx":6612
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   17
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
         MouseIcon       =   "frmingrxPlato_back.frx":68EF
         Picture         =   "frmingrxPlato_back.frx":6A41
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   16
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
         MouseIcon       =   "frmingrxPlato_back.frx":6EB6
         Picture         =   "frmingrxPlato_back.frx":7008
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   15
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
         MouseIcon       =   "frmingrxPlato_back.frx":74D3
         Picture         =   "frmingrxPlato_back.frx":7625
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   14
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
         MouseIcon       =   "frmingrxPlato_back.frx":7A5F
         Picture         =   "frmingrxPlato_back.frx":7BB1
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   13
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
         MouseIcon       =   "frmingrxPlato_back.frx":7D40
         Picture         =   "frmingrxPlato_back.frx":7E92
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
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
         MouseIcon       =   "frmingrxPlato_back.frx":8105
         Picture         =   "frmingrxPlato_back.frx":8257
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   28
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
         MouseIcon       =   "frmingrxPlato_back.frx":8378
         Picture         =   "frmingrxPlato_back.frx":84CA
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   29
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
         MouseIcon       =   "frmingrxPlato_back.frx":85E0
         Picture         =   "frmingrxPlato_back.frx":8732
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   30
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
         MouseIcon       =   "frmingrxPlato_back.frx":887F
         Picture         =   "frmingrxPlato_back.frx":89D1
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   31
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
         MouseIcon       =   "frmingrxPlato_back.frx":8B79
         Picture         =   "frmingrxPlato_back.frx":8CCB
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   32
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
         MouseIcon       =   "frmingrxPlato_back.frx":8E38
         Picture         =   "frmingrxPlato_back.frx":8F8A
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   33
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
         MouseIcon       =   "frmingrxPlato_back.frx":9112
         Picture         =   "frmingrxPlato_back.frx":9264
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   34
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
         MouseIcon       =   "frmingrxPlato_back.frx":93CE
         Picture         =   "frmingrxPlato_back.frx":9520
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   35
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
         MouseIcon       =   "frmingrxPlato_back.frx":96C5
         Picture         =   "frmingrxPlato_back.frx":9817
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   40
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
         MouseIcon       =   "frmingrxPlato_back.frx":9970
         Picture         =   "frmingrxPlato_back.frx":9AC2
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   41
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "frmingrxPlato_back.frx":9C56
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmingrxPlato_back.frx":9DAE
         Style           =   1  'Graphical
         TabIndex        =   43
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
         MouseIcon       =   "frmingrxPlato_back.frx":A22E
         Picture         =   "frmingrxPlato_back.frx":A380
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   42
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
         MouseIcon       =   "frmingrxPlato_back.frx":A800
         Picture         =   "frmingrxPlato_back.frx":A952
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   44
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
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from IngredientesPlatos order by idplato, codalimento"
      Top             =   2895
      Visible         =   0   'False
      Width           =   6180
   End
End
Attribute VB_Name = "frmingrxPlato_back"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msg As String
Dim tb As Recordset
'Public estadoAbm As Integer ' define el estado de un formulario de abm
'                             1 = sin cambios; 2 = agregar; 3 = modificar
'el modulo "fSetEnableFields(MDIForm1.ActiveForm, vbFalse)" se debe agregar al proyecto
Dim Titulo As String ' titulo del form


Private Sub cmdAceptar_Click()
If estadoAbm = 2 Or estadoAbm = 3 Then 'si el estado es agregar o modificar
'------------
    If Not Val(txtFields(3).Text) > 0 Then
        MsgBox "EL valor de la porcion debe ser mayor a cero", vbInformation
        txtFields(3).SetFocus
    
    Else
    
        If DataCombo1.Text = "" Then
            MsgBox "Debe completar el nombre del plato"
            DataCombo1.SetFocus
        Else
            If DataCombo2.Text = "" Then
                MsgBox "Debe completar el nombre del ingrediente"
                DataCombo2.SetFocus
            Else
                If txtFields(3).Text = "" Then
                    MsgBox "Debe completar la cantidad por porción"
                    txtFields(3).SetFocus
                Else
                    
                    strquery = "select * from ingredientesplatos where idplato = " & DataCombo1.BoundText & " and codalimento = " & DataCombo2.BoundText

                    Set tb = dbdiet.OpenRecordset(strquery)
        
                    If tb.RecordCount = 0 Then
    
                            MDIForm1.ActiveForm.Data1.UpdateRecord
                            MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
                        
                            txtFields(3).Enabled = False
                            
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
                
                        Else
    
                            If estadoAbm = 2 Then
                                MsgBox "El ingrediente seleccionado ya fue incluído dentro del plato con " & tb.Fields("porcion").Value & " gr/cc por porción"
                            Else
                                MDIForm1.ActiveForm.Data1.UpdateRecord
                                MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
                                
                                txtFields(3).Enabled = False
                                
                                cmdBuscar.Enabled = True
                                cmdAgregar.Enabled = True
                                cmdBorrar.Enabled = True
                                'cmdClose.Enabled = True
                                cmdModificar.Enabled = True
                                
                                cmdAgregar.SetFocus
                                cmdAgregar.Default = True
                                cmdCancelar.Cancel = True
                                
                                cmdPrimero.Enabled = True
                                cmdAnterior.Enabled = True
                                cmdSiguiente.Enabled = True
                                cmdUltimo.Enabled = True
                               
                                Call fSetEnableFields(MDIForm1.ActiveForm, vbFalse)
                            
                                estadoAbm = 1 ' el estado del form es "sin cambios"
                                
                                Call enabledDesplaz
        
                            End If
                                                    
                        End If
                        tb.Close
                               
                End If
            End If
        End If
               
    End If

'-------------
    'MDIForm1.ActiveForm.Data1.UpdateRecord
    'MDIForm1.ActiveForm.Data1.Recordset.Bookmark = MDIForm1.ActiveForm.Data1.Recordset.LastModified
    
    'condiciones extras
        'If estadoAbm = 2 Then
        '    dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(MDIForm1.ActiveForm.Label1.Caption) & ", codalimento from alimentos where estado = true"
        'End If
        
    Call f_Boton_Zorder
    
Else

    If Not MDIForm1.ActiveForm Is Nothing Then
    
        MDIForm1.ActiveForm.Hide
        
    End If

End If
End Sub

Private Sub cmdAgregar_Click()
Dim c
c = DataCombo1.BoundText

Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

estadoAbm = 2 ' el estado es agregar

strquery = "select * from IngredientesPlatos order by idplato, codalimento"
With Data1
    .RecordSource = strquery
    .Refresh
End With

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

Unload frm_Adm_Diet

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
        
                    strquery = "select * from menu where idalimento = " & DataCombo2.BoundText & " and idplato = " & DataCombo1.BoundText
        Set tb = dbdiet.OpenRecordset(strquery)
        If tb.RecordCount = 0 Then
            Data1.Recordset.Delete
            Data1.Recordset.MovePrevious
        Else
            MsgBox "No se puede eliminar el registro actual porque puede afectar la integridad del Sistema", , "Información"
        End If
        tb.Close
        
        Call f_Boton_Zorder
        
    Else
        cmdAgregar.SetFocus
    End If
End If

End Sub

Private Sub cmdBuscar_Click()
Dim strquery As String

strquery = "select * from IngredientesPlatos order by idplato, codalimento"
With Data1
    .RecordSource = strquery
    .Refresh
End With

'aclare campo por el cual buscar
'falta programar bien
msg = InputBox("Ingrese nombre del plato:", "Buscar por Nombre")

'strquery = " select * from ingredientesplatos, platos where ingredientesplatos.idplato = platos.idplato and nombreplato like '" & msg & "*' order by nombreplato"
'strquery = "select ingredientesplatos.idplato as idplato, ingredientesplatos.codalimento as codalimento, ingredientesplatos.porcion as porcion,alimentos.codalimento, alimentos.idcategoria, alimentos.descripalimento, platos.idplato, platos.nombreplato, categoria.idcategoria, categoria.decripcion from ingredientesplatos, alimentos, platos, categoria where ingredientesplatos.idplato = platos.idplato and ingredientesplatos.codalimento = alimentos.codalimento and alimentos.idcategoria = categoria.idcategoria and nombreplato like '" & msg & "*' order by nombreplato, decripcion, descripalimento"
If msg <> "" Then

    strquery = "select IngredientesPlatos.idplato as idplato, IngredientesPlatos.codalimento as codalimento, IngredientesPlatos.porcion as porcion from IngredientesPlatos, platos where ingredientesplatos.idplato = platos.idplato and nombreplato like '" & msg & "*' order by ingredientesplatos.idplato, codalimento"
    
    With MDIForm1.ActiveForm.Data1
        .RecordSource = strquery
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
    
        MDIForm1.ActiveForm.Hide
    
    End If

End If
End Sub



Private Sub cmdImprimir_Click()
Dim strquery As String
'aclare el filtro para imprimir
msg = MsgBox("¿Desea imprimir todos los registros?", vbYesNo, "Imprimir")
  
CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_alimxplato_one.rpt"

If msg = vbYes Then
    
    strquery = ""
    
Else
    
    strquery = " {ingredientesplatos.idplato} = " & Val(Label1.Caption)
    
End If

Call f_print(CrystalReport1, strquery, crptToWindow)

End Sub

Private Sub cmdModificar_Click()

Call fSetEnableFields(MDIForm1.ActiveForm, vbTrue)

plato = DataCombo1.BoundText
Alimento = DataCombo2.BoundText

strquery = "select * from IngredientesPlatos order by idplato, codalimento"

With Data1
    .RecordSource = strquery
    .Refresh
End With

If MDIForm1.ActiveForm.Data1.Recordset.BOF = True Or MDIForm1.ActiveForm.Data1.Recordset.EOF = True Then
    MDIForm1.ActiveForm.Data1.Recordset.MoveFirst
Else
    If plato <> "" And Alimento <> "" Then
        Data1.Recordset.FindFirst " idplato = " & plato & " and codalimento = " & Alimento
    End If
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

Private Sub Form_Activate()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 3615
Me.Width = 6270
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Call enabledDesplaz

Call f_Boton_Zorder

End Sub




Private Sub Command1_Click()
Dim strquery As String

'strquery = "select ingredientesplatos.idplato, ingredientesplatos.codalimento, ingredientesplatos.porcion,alimentos.codalimento, alimentos.idcategoria, alimentos.descripalimento, platos.idplato, platos.nombreplato, categoria.idcategoria, categoria.decripcion from ingredientesplatos, alimentos, platos, categoria where ingredientesplatos.idplato = platos.idplato and ingredientesplatos.codalimento = alimentos.codalimento and alimentos.idcategoria = categoria.idcategoria order by nombreplato, decripcion, descripalimento"
strquery = "select * from IngredientesPlatos order by idplato, codalimento"
With Data1
    .RecordSource = strquery
    .Refresh
End With

Call enabledDesplaz

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
    MsgBox "Debe Completar el Nombre del Plato", vbInformation, "Información"
End If

End Sub

Private Sub DataCombo2_LostFocus()
If DataCombo2.Text = "" Then
    DataCombo2.SetFocus
    MsgBox "Debe Completar el Nombre del Ingrediente", vbInformation, "Información"
End If

End Sub

Private Sub Form_Load()

'Data1.DatabaseName = Lugar

Call f_CargarOrigenDatos

txtFields(3).Enabled = False

estadoAbm = 1

Titulo = Me.Caption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdCancelar_Click

End Sub

Private Sub Label1_Change()
Me.Caption = Titulo & " - Nro. " & Val(Label1.Caption)
End Sub

Private Sub Pic_Imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Imprimir

End Sub

Private Sub txtFields_GotFocus(Index As Integer)

    txtFields(3).SelStart = 0
    txtFields(3).SelLength = 50


End Sub

Private Sub txtFields_Validate(Index As Integer, Cancel As Boolean)

If Not IsNumeric(txtFields(Index)) Or Not Val(txtFields(Index).Text) > 0 Then
    MsgBox "EL valor de la porcion debe ser numerico mayor a cero", vbInformation
    'cancel= Valor que indica si el control pierde el foco.
    'Establecer cancel con el valor True indica que el control mantiene el foco.
    Cancel = True
End If

End Sub

Sub f_CargarOrigenDatos()
Dim strquery As String
strquery = ""

strquery = "select * from IngredientesPlatos order by idplato, codalimento"
Call f_Data_DatabaseName(Data1, strquery)

strquery = "select * from platos order by nombreplato"
Call f_Adodc_ConnectionString(Adodc1, strquery)

strquery = "select alimentos.codalimento, alimentos.idcategoria, alimentos.descripalimento, alimentos.hc, alimentos.prot, alimentos.lip, alimentos.estado, (categoria.decripcion & ' ,  ' & alimentos.descripalimento) as nom from alimentos, categoria where alimentos.idcategoria = categoria.idcategoria order by categoria.decripcion, alimentos.descripalimento"
Call f_Adodc_ConnectionString(Adodc2, strquery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo1, Data1, Adodc1, "idPlato", "idPlato", "NombrePlato")

Call f_Enlaza_ControlData(DataCombo2, Data1, Adodc2, "CodAlimento", "CodAlimento", "nom")

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

