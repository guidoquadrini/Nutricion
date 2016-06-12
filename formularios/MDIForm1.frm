VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H00FFFFFF&
   Caption         =   "Omnia"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   150
   ClientWidth     =   10770
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0ECA
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList Toolbar_Umana_Disabled 
      Left            =   1440
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   41
      ImageHeight     =   38
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E3B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10920
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13037
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15B81
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18215
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   926
      ButtonWidth     =   1138
      ButtonHeight    =   873
      Appearance      =   1
      Style           =   1
      ImageList       =   "Toolbar_Umana_Enabled"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pacientes"
            Object.ToolTipText     =   "Pacientes"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "FinalizarCosulta"
            Object.ToolTipText     =   "Finalizar consulta actual"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "FormulaSintetica"
            Object.ToolTipText     =   "Fórmula Sintética"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "FormulaDesarrollada"
            Object.ToolTipText     =   "Fórmula Desarrollada"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "GestionPlanAlimentario"
            Object.ToolTipText     =   "Gestion de Planes Alimentarios"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "MDIForm1.frx":18667
      OLEDropMode     =   1
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   6780
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13335
            Text            =   "Estado"
            TextSave        =   "Estado"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "21/09/2011"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "03:28 p.m."
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Toolbar_Umana_Enabled 
      Left            =   1560
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   36
      ImageHeight     =   27
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":187C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":18E04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":19462
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":19867
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":19DFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A334
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3360
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIconsDoc 
      Left            =   3840
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A786
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A898
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A9AA
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1AABC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1ABCE
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1ACE0
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1ADF2
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1AF04
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B016
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B128
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B23A
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B34C
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B45E
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBarDoc 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIconsDoc"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abrir"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cortar"
            Object.ToolTipText     =   "Cortar"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pegar"
            Object.ToolTipText     =   "Pegar"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Negrita"
            Object.ToolTipText     =   "Negrita"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cursiva"
            Object.ToolTipText     =   "Cursiva"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Subrayado"
            Object.ToolTipText     =   "Subrayado"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Alinear a la izquierda"
            Object.ToolTipText     =   "Alinear a la izquierda"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Centrar"
            Object.ToolTipText     =   "Centrar"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Alinear a la derecha"
            Object.ToolTipText     =   "Alinear a la derecha"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox CrystalReport1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   10710
      TabIndex        =   3
      Top             =   945
      Width           =   10770
   End
   Begin VB.Menu arc 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "A&brir..."
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu profesionales 
         Caption         =   "Pro&fesionales"
      End
      Begin VB.Menu consultorios 
         Caption         =   "&Consultorios"
      End
      Begin VB.Menu mnu_MantAlimentos 
         Caption         =   "Man&tenimiento Alimentos"
         Begin VB.Menu cat 
            Caption         =   "&Grupos de Alimentos"
         End
         Begin VB.Menu ali 
            Caption         =   "&Alimentos"
         End
         Begin VB.Menu plato 
            Caption         =   "&Platos"
         End
         Begin VB.Menu ingr 
            Caption         =   "&Ingredietes de plato"
         End
      End
      Begin VB.Menu porcentajeComida1 
         Caption         =   "P&orcentaje de Kcal por Comida"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu cerrar 
         Caption         =   "&Cerrar"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Guardar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "G&uardar como..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "C&onfigurar página..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "&Vista preliminar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "I&mprimir..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu sal 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edición"
      Visible         =   0   'False
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Deshacer"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cor&tar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnu_Pacientes 
      Caption         =   "&Pacientes"
      Begin VB.Menu mnu_ModoAcceso 
         Caption         =   "Modo de ACCESO"
         Begin VB.Menu mnu_ConsultaNueva 
            Caption         =   "Nueva Consulta..."
         End
         Begin VB.Menu mnu_ConsultaExistente 
            Caption         =   "Consulta Existente..."
         End
      End
      Begin VB.Menu mnu_FinalizarCosulta 
         Caption         =   "Finalizar consulta actual"
      End
      Begin VB.Menu bar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Mantenimiento 
         Caption         =   "Mantenimiento"
      End
      Begin VB.Menu mnu_Evaluacion_Subjetiva 
         Caption         =   "&Evaluacion global subjetiva"
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_formulas 
         Caption         =   "&Formulas"
         Begin VB.Menu mnu_FormulaSintetica 
            Caption         =   "&Sintética"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnu_FormulaDesarollada 
            Caption         =   "&Desarrollada"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnu_GestionPlanAlimentario 
         Caption         =   "&Gestion Planes Alimentarios"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_Turnos 
      Caption         =   "&Turnos"
      Begin VB.Menu mnu_agen_turno 
         Caption         =   "&Agenda de Turnos"
      End
      Begin VB.Menu horarios 
         Caption         =   "&Horarios de Profesionales"
      End
      Begin VB.Menu bar4 
         Caption         =   "-"
      End
      Begin VB.Menu excepciones 
         Caption         =   "E&xcepciones Horarias"
      End
      Begin VB.Menu feriado 
         Caption         =   "&Feriados"
      End
   End
   Begin VB.Menu mnu_Herramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuInforme 
         Caption         =   "&Editor de texto"
      End
   End
   Begin VB.Menu mnu_Documentos 
      Caption         =   "&Documentos"
      Begin VB.Menu equivalencias 
         Caption         =   "E&quivalencias"
      End
      Begin VB.Menu mnu_HD 
         Caption         =   "&Historia dietetica"
         Begin VB.Menu mnu_Anamnesis 
            Caption         =   "&Anamnesis Alimentaria"
         End
         Begin VB.Menu mnu_info_Complementaria 
            Caption         =   "Informacion Complementaria"
         End
      End
      Begin VB.Menu mnu_Reg7Dias 
         Caption         =   "Planilla &Registro de 7 dias"
      End
      Begin VB.Menu mnu_Piramide_Alimentaria 
         Caption         =   "&Piramide Alimentaria"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "&Ver"
      WindowList      =   -1  'True
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Barra de herramientas"
         Begin VB.Menu mnutexto 
            Caption         =   "&Barra de texto"
         End
         Begin VB.Menu mnuherramientas 
            Caption         =   "&Barra de herramientas"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "B&arra de estado"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu cascada 
         Caption         =   "&Cascada"
      End
      Begin VB.Menu mosaico 
         Caption         =   "&Mosaico"
      End
      Begin VB.Menu organizar 
         Caption         =   "&Organizar iconos"
      End
   End
   Begin VB.Menu EdiCion 
      Caption         =   "EdiCion"
      Visible         =   0   'False
      Begin VB.Menu Cantidad 
         Caption         =   "Cantidad"
         Enabled         =   0   'False
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu Agregar 
         Caption         =   "Agregar"
         Begin VB.Menu platos 
            Caption         =   "Plato"
            Enabled         =   0   'False
         End
         Begin VB.Menu ingred 
            Caption         =   "Ingrediente"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu eliminar 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu popupFeriado 
      Caption         =   "popupFeriado"
      Visible         =   0   'False
      Begin VB.Menu Expandir 
         Caption         =   "Ver todo el año"
      End
      Begin VB.Menu Contraer 
         Caption         =   "Ver por mes"
      End
   End
   Begin VB.Menu popupTurnos 
      Caption         =   "popupTurnos"
      Visible         =   0   'False
      Begin VB.Menu tur_imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu tur_eliminar 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu ay 
      Caption         =   "Ay&uda"
      Begin VB.Menu Contenido 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnu_Creditos 
         Caption         =   "Cre&ditos"
      End
      Begin VB.Menu bar5 
         Caption         =   "-"
      End
      Begin VB.Menu Acerca 
         Caption         =   "&Acerca de"
      End
   End
   Begin VB.Menu mnuEliminar 
      Caption         =   "Eliminar"
      Visible         =   0   'False
      Begin VB.Menu mnuEli 
         Caption         =   "Eliminar seleccionados"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
Dim tb1 As Recordset
'CantVentDoc mantiene la cantidad de ventanas de documentos abiertas para poder
'determinar en que momento tengo una sola ventana para poder, al cerrar, cambiar el menu actual
Public CantVentDoc As Integer

Private Sub Acerca_Click()
frmAbout.Show
End Sub

Private Sub ali_Click()
frm_abm_Alimentos.Show

End Sub

Private Sub Cantidad_Click()
Dim cant As String

cant = InputBox("Ingrese Cantidad", "Actualizar")
If IsNumeric(cant) Then

    CantFinal = Int(cant)
    If CantFinal <> 0 Then
        frm_Adm_Diet.DevuelveUnidad (CodigoPlato(1))
        frm_Adm_Diet.DevuelveCantNeta (CodigoPlato(1))
                    
        If UnidadPlato = 1 Then
            'dbdiet.Execute "update platosmenu set cantidad = " & CantFinal & ", cantNeta = " & CantFinal & " where legajo = " & frm_Adm_Diet.DataCombo1.BoundText & " and idtpomenu = " & frm_Adm_Diet.TabStrip2.SelectedItem.Index & " and idplato = " & CodigoPlato(1) & " and fechaMenu = " & "#" & Month(frm_Adm_Diet.DTPicker1.Value) & "/" & Day(frm_Adm_Diet.DTPicker1.Value) & "/" & Year(frm_Adm_Diet.DTPicker1.Value) & "#"
            dbdiet.Execute "update platosmenu_tmp set cantidad = " & CantFinal & ", cantNeta = " & CantFinal & " where legajo = " & frm_Adm_Diet.DataCombo1.BoundText & " and idtpomenu = " & frm_Adm_Diet.TabStrip2.SelectedItem.Index & " and idplato = " & CodigoPlato(1) & " and fechaMenu = " & "#" & Month(frm_Adm_Diet.DTPicker1.Value) & "/" & Day(frm_Adm_Diet.DTPicker1.Value) & "/" & Year(frm_Adm_Diet.DTPicker1.Value) & "#"
        Else
            'dbdiet.Execute "update platosmenu set cantidad = " & CantFinal & ", cantNeta = " & CantFinal * CantidadNeta & " where legajo = " & frm_Adm_Diet.DataCombo1.BoundText & " and idtpomenu = " & frm_Adm_Diet.TabStrip2.SelectedItem.Index & " and idplato = " & CodigoPlato(1) & " and fechaMenu = " & "#" & Month(frm_Adm_Diet.DTPicker1.Value) & "/" & Day(frm_Adm_Diet.DTPicker1.Value) & "/" & Year(frm_Adm_Diet.DTPicker1.Value) & "#"
            dbdiet.Execute "update platosmenu_tmp set cantidad = " & CantFinal & ", cantNeta = " & CantFinal * CantidadNeta & " where legajo = " & frm_Adm_Diet.DataCombo1.BoundText & " and idtpomenu = " & frm_Adm_Diet.TabStrip2.SelectedItem.Index & " and idplato = " & CodigoPlato(1) & " and fechaMenu = " & "#" & Month(frm_Adm_Diet.DTPicker1.Value) & "/" & Day(frm_Adm_Diet.DTPicker1.Value) & "/" & Year(frm_Adm_Diet.DTPicker1.Value) & "#"
        End If
        CantFinal = CantFinal - CantAux
        PrevCantAux = CantAux
        frm_Adm_Diet.Label1(1) = Val(frm_Adm_Diet.Label1(1).Caption) + (Suma * CantFinal)
        frm_Adm_Diet.CantidadPopup
    End If
Else

    MsgBox "Debe ingresar un valor numerico. Se considera solo la parte entera", vbInformation

End If

End Sub

Private Sub cascada_Click()
MDIForm1.Arrange vbcascada
End Sub

Private Sub cat_Click()
frmcateg.Show

End Sub

Private Sub cerrar_Click()
If mnuEdit.Visible = True Then
            
    If CantVentDoc = 1 Then
    
        MDIForm1.tbToolBarDoc.Visible = False
        MDIForm1.tbToolBar.Visible = True
    
        MDIForm1.mnuFileNew.Visible = False
        MDIForm1.mnuFileOpen.Visible = False
        MDIForm1.mnuFileSave.Visible = False
        MDIForm1.mnuFileSaveAs.Visible = False
        MDIForm1.mnuFileBar1.Visible = False
        MDIForm1.mnuFilePageSetup.Visible = False
        MDIForm1.mnuFilePrintPreview.Visible = False
        MDIForm1.mnuFilePrint.Visible = False
        MDIForm1.mnuFileBar2.Visible = False
        MDIForm1.mnuEdit.Visible = False
        
        MDIForm1.mnutexto.Checked = False
        MDIForm1.mnuherramientas.Checked = True
        
        CantVentDoc = 0
        
        mnuInforme.Enabled = True
    Else
        CantVentDoc = CantVentDoc - 1
    End If
    
End If

'If ActiveForm Is Nothing Then
'Else
'ActiveForm.Hide
'End If

If Not MDIForm1.ActiveForm Is Nothing Then
    
    MDIForm1.ActiveForm.Hide
    
End If
    
End Sub

Private Sub consultorios_Click()
frm_abm_consul.Show
End Sub

Private Sub Contenido_Click()
frmBrowser.Show
'Unload frmAbout

End Sub

Private Sub Contraer_Click()

Call frm_calendario.Size("mes")

End Sub



Private Sub eliminar_Click()
Dim cgoplato() As String
cgoplato = Split(frm_Adm_Diet.TreeView1.SelectedItem.Key, "//")
indice = frm_Adm_Diet.TreeView1.SelectedItem.Index

'If frm_Adm_Diet.TreeView1.SelectedItem.Children = 0 Then
If cgoplato(2) = "Ingrediente" Then
    msg = MsgBox("¿Desea Eliminar el Ingrediente Seleccionado?", vbYesNo, "Eliminar")
    
    If msg = vbYes Then
        'verifica que se pueda eliminar sin problemas y no perder integridad
        strQuery = "select * from menu where idalimento = " & cgoplato(1) & " and idplato = " & cgoplato(3)
        Set tb = dbdiet.OpenRecordset(strQuery)
        If tb.RecordCount = 0 Then
            dbdiet.Execute "delete from ingredientesplatos where idplato = " & cgoplato(3) & " and codalimento = " & cgoplato(1)
            'Data1.Recordset.Delete
            'Data1.Recordset.MovePrevious
            frm_Adm_Diet.TreeView1.Nodes.Remove (indice)
        Else
            MsgBox "No se puede eliminar el ingrediente seleccionado porque hay otro paciente que lo está usando", , "Información"
        End If
        tb.Close
                
    Else
        frm_Adm_Diet.TreeView1.SetFocus
    End If
    
    'For i = 0 To 3
    '    MsgBox cgoplato(i)
    'Next
Else
    
    msg = MsgBox("¿Desea Eliminar el Plato Seleccionado?", vbYesNo, "Eliminar")
    
    If msg = vbYes Then
        'verifica que se pueda eliminar sin problemas y no perder integridad
        strQuery = "select * from menu where idplato = " & cgoplato(1)
        Set tb = dbdiet.OpenRecordset(strQuery)
        If tb.RecordCount = 0 Then
            dbdiet.Execute "delete from platos where idplato = " & cgoplato(1)
            dbdiet.Execute "delete from ingredientesplatos where idplato = " & cgoplato(1)
            'Data1.Recordset.Delete
            'Data1.Recordset.MovePrevious
            frm_Adm_Diet.TreeView1.Nodes.Remove (indice)
        Else
            MsgBox "No se puede eliminar el plato seleccionado porque puede hay otro paciente que lo está usando", , "Información"
        End If
        tb.Close
        
    Else
        frm_Adm_Diet.TreeView1.SetFocus
    End If
    
    'For i = 0 To 2
    '    MsgBox cgoplato(i)
    'Next
End If
End Sub

Private Sub equivalencias_Click()
Dim strQuery As String

CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_equivalencias_one.rpt"

strQuery = ""

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub

Private Sub excepciones_Click()
frm_ExcepcionesHrs.Show
End Sub

Private Sub Expandir_Click()

Call frm_calendario.Size("año")

End Sub

Private Sub feriado_Click()
Dim strQuery As String

strQuery = "DblClick para Agregar o Eliminar un feriado"

frm_calendario.cargarParametros 0, , strQuery
frm_calendario.Show

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
'--------------
'Cierra el origen de datos OLE DB
'Cn.Close
'Set Cn = Nothing
'--------------
dbdiet.Close

End
End Sub

Private Sub mnu_ConsultaExistente_Click()

Set frm_ModoAtencion.TabStrip1.SelectedItem = frm_ModoAtencion.TabStrip1.Tabs(2)
frm_ModoAtencion.Show vbModal

End Sub

Private Sub mnu_ConsultaNueva_Click()

Set frm_ModoAtencion.TabStrip1.SelectedItem = frm_ModoAtencion.TabStrip1.Tabs(1)
frm_ModoAtencion.Show vbModal

End Sub

Private Sub mnu_Creditos_Click()

frm_Creditos.Show vbModal

End Sub

Private Sub mnu_Evaluacion_Subjetiva_Click()
frm_Evaluacion_Subjetiva.Show

End Sub

Private Sub horarios_Click()
frm_abm_horarios.Show

End Sub

Private Sub ingr_Click()
frmingrxPlato.Show

End Sub

Private Sub ingred_Click()
Dim tempNode As Node
Dim Descripcion, Categoria As String

AgregarIngrediente.Show vbModal

If IngredAgregado <> 0 And PorcionAgregado > 0 Then
    
    dbdiet.Execute "insert into IngredientesPlatos (idplato, codalimento, porcion) values (" & CodigoPlato(1) & ", " & IngredAgregado & ", " & PorcionAgregado & ")"
    
    strQuery = " select * from alimentos, categoria where alimentos.idcategoria = categoria.idcategoria and alimentos.codalimento = " & IngredAgregado
    Set tb = dbdiet.OpenRecordset(strQuery)
    Descripcion = tb.Fields("descripalimento").Value
    Categoria = tb.Fields("decripcion").Value
    tb.Close
       
    Set tempNode = frm_Adm_Diet.TreeView1.Nodes.Add(KeyPlato, tvwChild, Descripcion & "//" & IngredAgregado & "//Ingrediente//" & CodigoPlato(1), Categoria & ", " & Descripcion)
    
    'para poder reindexar los nodos en orden creciente para luego poder seleccionarlos correctamente
    Unload frm_Adm_Diet
    frm_Adm_Diet.Show
End If

End Sub

Private Sub MDIForm_Click()
'ActiveForm.ActiveForm.Print "hola"
End Sub

Private Sub MDIForm_Load()
'lugar = "c:\dietetica\db1nueva prueba anterior sin replica.mdb"
'Set dbdiet = OpenDatabase(lugar)
Me.Acerca.Caption = Me.Acerca.Caption & " " & App.Title & "..."
CantVentDoc = 0


'Titulo Main ===============================
Select Case sUserNivel
           
    Case Is = "Nutricionista"
        Me.Caption = "Omnia - Usuario: " & f_GetUserLoging(nUserLoging)
    
    Case Else
        Me.Caption = "Omnia - Usuario: " & sUserNivel

End Select
'===========================================

'Modo de ACCESO standard
nCodAtencion = 0
Call f_GetPermisos

End Sub
Private Sub Form_Paint()
    lvListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    Select Case lvListView.View
        Case lvwIcon
            tbToolBar.Buttons(LISTVIEW_MODE0).Value = tbrPressed
        Case lvwSmallIcon
            tbToolBar.Buttons(LISTVIEW_MODE1).Value = tbrPressed
        Case lvwList
            tbToolBar.Buttons(LISTVIEW_MODE2).Value = tbrPressed
        Case lvwReport
            tbToolBar.Buttons(LISTVIEW_MODE3).Value = tbrPressed
    End Select
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'--------------
'Cierra el origen de datos OLE DB
'Cn.Close
'Set Cn = Nothing
'--------------
'dbdiet.Close
End Sub

Private Sub mnu_Anamnesis_Click()
Dim strQuery As String

CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_HD_Anamnesis_Alimentaria_one.rpt"

strQuery = ""

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub

Private Sub mnu_GestionPlanAlimentario_Click()

frm_Adm_Diet.Show

End Sub


Private Sub mnu_agen_turno_Click()

frm_agendaTurnos.Show

End Sub




Private Sub mnu_FinalizarCosulta_Click()
Dim sMsg As String

sMsg = MsgBox("¿Desea finalizar la consulta actual?", vbYesNo)

If sMsg = vbYes Then

    'modo de acceso standard
    nCodAtencion = 0
    
    Call f_GetPermisos

End If

End Sub

Private Sub mnu_info_Complementaria_Click()
Dim strQuery As String

CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_HD_Info_Complementaria_one.rpt"

strQuery = ""

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub

Private Sub mnu_Piramide_Alimentaria_Click()
Dim strQuery As String

CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_Piramide_Alimentaria_one.rpt"

strQuery = ""

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub

Private Sub mnu_Reg7Dias_Click()
Dim strQuery As String

CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_Registro7Dias_one.rpt"

strQuery = ""

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub



Private Sub mnuEli_Click()
Dim fileSys As New FileSystemObject, fil As File
Dim sFile As String

msg = MsgBox("¿Desea eliminar los elementos seleccionados?", vbYesNo, "Eliminar")

If msg = vbYes Then

    For i = 0 To ReporteMenu.List1.ListCount - 1
       
        If ReporteMenu.List1.Selected(i) = True Then
            ReporteMenu.List1.ListIndex = i
            
            ArmaArchivo = Split(ReporteMenu.List1.Text, " ")
            ArmaArchivo1 = Split(ArmaArchivo(1), "/")
            
            sFile = App.Path & "\Reportes\Rep_" & ReporteMenu.DataCombo1.BoundText & "_" & ArmaArchivo1(2) & "_" & ArmaArchivo1(1) & "_" & ArmaArchivo1(0) & "_.doc"
            
            Set fil = fileSys.GetFile(sFile)
            
            fil.Delete True
                           
        Else
                    
        End If
    Next
End If

ReporteMenu.BuscarInformes
End Sub

Private Sub mnuherramientas_Click()

    mnuherramientas.Checked = Not mnuherramientas.Checked
    tbToolBar.Visible = mnuherramientas.Checked
    
End Sub

Public Sub mnuInforme_Click()

LoadNewDoc
tbToolBarDoc.Visible = True
tbToolBar.Visible = False

' visualizar el menu de texto
mnuFileNew.Visible = True
mnuFileOpen.Visible = True
mnuFileSave.Visible = True
mnuFileSaveAs.Visible = True
mnuFileBar1.Visible = True
mnuFilePageSetup.Visible = True
mnuFilePrintPreview.Visible = True
mnuFilePrint.Visible = True
mnuFileBar2.Visible = True
mnuEdit.Visible = True

mnutexto.Checked = True
mnuherramientas.Checked = False

mnuInforme.Enabled = False

End Sub

Private Sub mnutexto_Click()
    
    mnutexto.Checked = Not mnutexto.Checked
    tbToolBarDoc.Visible = mnutexto.Checked

End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    

End Sub

Private Sub mosaico_Click()
MDIForm1.Arrange vbTileHorizontal
End Sub

Private Sub organizar_Click()
MDIForm1.Arrange vbArrangeIcons
End Sub

Private Sub mnu_Mantenimiento_Click()
frmPacientes.Show

End Sub

Private Sub plato_Click()
frmplato.Show

End Sub

Private Sub platos_Click()
Dim tempNode As Node

AgregarPlato.Show vbModal

If PlatoAgregado <> "" And unidadAgregado <> 0 Then
    dbdiet.Execute " insert into platos (nombreplato, idunidad) values ('" & PlatoAgregado & "', " & unidadAgregado & ")"

    Set tb = dbdiet.OpenRecordset("platos")
    tb.MoveLast
    Id = tb.Fields("idplato").Value
    tb.Close
    
    Set tempNode = frm_Adm_Diet.TreeView1.Nodes.Add(, , PlatoAgregado & "//" & Id & "//Plato", PlatoAgregado)
    
End If


End Sub





Private Sub porcentajeComida1_Click()

PorcentajeComida.Show vbModal


End Sub

Private Sub mnu_FormulaSintetica_Click()
'PrincipalFrm.Show
frm_FormulaSintetica.Show

End Sub

Private Sub profesionales_Click()
'frm_abm_prof.Show
frm_abm_Profesionales.Show
End Sub

Private Sub sal_Click()
End
End Sub

Private Sub mnu_FormulaDesarollada_Click()
frm_formulaDesarrollada.Show

End Sub

Private Sub Timer1_Timer()


End Sub

Private Sub usuario_Click()
frmpru.Show

End Sub





Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    'MsgBox Button & " y  " & Button.Key
    On Error Resume Next
    Select Case Button.Key
        Case "pacientes"
            'TareasPendientes: Agregar código de botón 'Atrás'.
            frmPacientes.Show
            'MsgBox "Agregar código de botón 'Atrás'."
        Case "FinalizarCosulta"
            Call mnu_FinalizarCosulta_Click
        Case "FormulaSintetica"
            'TareasPendientes: Agregar código de botón 'Adelante'.
            'PrincipalFrm.Show
            Call mnu_FormulaSintetica_Click
            'MsgBox "Agregar código de botón 'Adelante'."
        Case "FormulaDesarrollada"
            mnu_FormulaDesarollada_Click
            'mnuEditCut_Click
        Case "GestionPlanAlimentario"
            mnu_GestionPlanAlimentario_Click
            'mnuEditCopy_Click
        Case "Salir"
            sal_Click
        
    End Select
    
End Sub

Private Sub tbToolBarDoc_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Nuevo"
            LoadNewDoc
        Case "Abrir"
            mnuFileOpen_Click
        Case "Guardar"
            mnuFileSave_Click
        Case "Imprimir"
            mnuFilePrint_Click
        Case "Cortar"
            mnuEditCut_Click
        Case "Copiar"
            mnuEditCopy_Click
        Case "Pegar"
            mnuEditPaste_Click
        Case "Negrita"
            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
            Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
        Case "Cursiva"
            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
            Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Subrayado"
            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
            Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Alinear a la izquierda"
            ActiveForm.rtfText.SelAlignment = rtfLeft
        Case "Centrar"
            ActiveForm.rtfText.SelAlignment = rtfCenter
        Case "Alinear a la derecha"
            ActiveForm.rtfText.SelAlignment = rtfRight
    End Select
End Sub
'------------------------------------
Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF
    ActiveForm.rtfText.SelText = vbNullString

End Sub

Private Sub mnuEditUndo_Click()
    'TareasPendientes: Agregar código 'mnuEditUndo'.
    MsgBox "Agregar código 'mnuEditUndo'."
End Sub


Private Sub mnuFileExit_Click()
    'descargar el formulario
    Unload ActiveForm

End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Imprimir"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
    'TareasPendientes: Agregar código 'mnuFilePrintPreview'.
    MsgBox "Agregar código 'mnuFilePrintPreview'."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Configurar página"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    

    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Guardar como"
        .DefaultExt = ".rtf"
        .CancelError = False
        'Pendiente: establecer los indicadores y atributos del control common dialog
        .Filter = "Todos los archivos (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.Caption = sFile
    ActiveForm.rtfText.SaveFile sFile

End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With dlgCommonDialog
            .DialogTitle = "Guardar"
            .CancelError = False
            'Pendiente: establecer los indicadores y atributos del control common dialog
            .Filter = "Todos los archivos (*.rtf)|*.rtf"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub


Private Sub mnuFileOpen_Click()
    Dim sFile As String


    If ActiveForm Is Nothing Then LoadNewDoc
    

    With dlgCommonDialog
        .DialogTitle = "Abrir"
        .CancelError = False
        'Pendiente: establecer los indicadores y los atributos del control common dialog
        .Filter = "Todos los archivos (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.rtfText.LoadFile sFile
    ActiveForm.Caption = sFile

End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
        
End Sub

Public Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Document " & lDocumentCount
    frmD.Show
    CantVentDoc = CantVentDoc + 1
End Sub


Private Sub tur_eliminar_Click()

Call frm_agendaTurnos.f_eliminarTurno

End Sub

Private Sub tur_imprimir_Click()
Dim strQuery As String


frm_agendaTurnos.CrystalReport1.Reset

frm_agendaTurnos.CrystalReport1.ReportFileName = App_Path & "\rpts\rep_turnos_one.rpt"

strQuery = " {turnos.tur_idprof} = " & frm_agendaTurnos.DataCombo1.BoundText & _
           " and {turnos.tur_fecha} = Date (" & Year(frm_agendaTurnos.DTPicker1.Value) & ", " & Month(frm_agendaTurnos.DTPicker1.Value) & ", " & Day(frm_agendaTurnos.DTPicker1.Value) & ")" & _
           " and {Turnos.tur_hrDesde} = '" & frm_agendaTurnos.DataGrid1.Columns("Horario").Value & "'"

Call f_print(frm_agendaTurnos.CrystalReport1, strQuery, crptToWindow)



End Sub
