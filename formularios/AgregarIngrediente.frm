VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form AgregarIngrediente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "AgregarIngrediente.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalles"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   2
         Top             =   1245
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "AgregarIngrediente.frx":0ECA
         DataField       =   "CodAlimento"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label2 
         Caption         =   "Cantidad en grs. o cc _________________________"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   3855
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6000
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porción:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1260
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ingrediente:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6240
      Top             =   840
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
      Begin VB.PictureBox Pic_Aceptar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         MouseIcon       =   "AgregarIngrediente.frx":0EDF
         Picture         =   "AgregarIngrediente.frx":1031
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Pic_Cancelar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         MouseIcon       =   "AgregarIngrediente.frx":12ED
         Picture         =   "AgregarIngrediente.frx":143F
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Pic_Cancelar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         MouseIcon       =   "AgregarIngrediente.frx":1740
         Picture         =   "AgregarIngrediente.frx":1892
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Pic_Aceptar_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         MouseIcon       =   "AgregarIngrediente.frx":1A26
         Picture         =   "AgregarIngrediente.frx":1B78
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "AgregarIngrediente.frx":1CD1
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "AgregarIngrediente.frx":1E2A
         Picture         =   "AgregarIngrediente.frx":1F7C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Aceptar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "AgregarIngrediente.frx":2238
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "AgregarIngrediente.frx":23CC
         Picture         =   "AgregarIngrediente.frx":251E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancelar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   375
      End
   End
End
Attribute VB_Name = "AgregarIngrediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
IngredAgregado = DataCombo2.BoundText
PorcionAgregado = Val(txtFields(3).Text)
Unload Me
'frm_Adm_Diet.Show

End Sub

Private Sub cmdCancelar_Click()
IngredAgregado = 0
PorcionAgregado = 0
Unload Me
'frm_Adm_Diet.Show

End Sub

Private Sub Form_Load()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 2745
Me.Width = 6240
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Call f_CargarOrigenDatos

cmdAceptar.Default = True
cmdCancelar.Cancel = True

Call f_Boton_Zorder

'DataCombo2.BoundText = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdCancelar_Click

End Sub

Sub f_CargarOrigenDatos()
Dim strquery As String
strquery = ""

strquery = "select alimentos.codalimento, alimentos.idcategoria, alimentos.descripalimento, alimentos.hc, alimentos.prot, alimentos.lip, alimentos.estado, (categoria.decripcion & ' ,  ' & alimentos.descripalimento) as nom from alimentos, categoria where alimentos.idcategoria = categoria.idcategoria order by categoria.decripcion, alimentos.descripalimento"
Call f_Data_DatabaseName(Data1, strquery)

strquery = "select alimentos.codalimento, alimentos.idcategoria, alimentos.descripalimento, alimentos.hc, alimentos.prot, alimentos.lip, alimentos.estado, (categoria.decripcion & ' ,  ' & alimentos.descripalimento) as nom from alimentos, categoria where alimentos.idcategoria = categoria.idcategoria order by categoria.decripcion, alimentos.descripalimento"
Call f_Adodc_ConnectionString(Adodc2, strquery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo2, Adodc2, Adodc2, "CodAlimento", "CodAlimento", "nom")

'==============================================

End Sub

Sub f_Boton_Zorder()

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

Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Sub f_Aceptar()

Me.cmdAceptar.ZOrder 0

Me.cmdCancelar.ZOrder 1

End Sub

Sub f_Cancelar()

Me.cmdCancelar.ZOrder 0

Me.cmdAceptar.ZOrder 1

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub



Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call f_Aceptar

End Sub


Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call f_Cancelar

End Sub
