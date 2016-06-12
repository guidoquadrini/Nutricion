VERSION 5.00
Begin VB.UserControl ctrlPrueba 
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   ScaleHeight     =   825
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdPrimero 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":0000
      Height          =   375
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":0710
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAnterior 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":0E14
      Height          =   375
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":1524
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBuscar 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":1C28
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":2338
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdSiguiente 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":2A3C
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":314C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdUltimo 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":3850
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":3F60
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAgregar 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":4664
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":4D74
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBorrar 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":5478
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":5B88
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdModificar 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":628C
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":698E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":7092
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":77A2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":7EA6
      Height          =   375
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":85B6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      DisabledPicture =   "ctrlPrueba.ctx":8CBA
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "ctrlPrueba.ctx":93CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Label ContenedorBotones 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5745
   End
End
Attribute VB_Name = "ctrlPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
