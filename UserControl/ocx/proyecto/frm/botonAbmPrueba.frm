VERSION 5.00
Begin VB.Form botonAbmPrueba 
   Caption         =   "Form1"
   ClientHeight    =   705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":0000
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":0710
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":0E14
      Height          =   375
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":1524
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":1C28
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":2338
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdModificar 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":2A3C
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":313E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBorrar 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":3842
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":3F52
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAgregar 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":4656
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":4D66
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdUltimo 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":546A
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":5B7A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdSiguiente 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":627E
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":698E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBuscar 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":7092
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":77A2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAnterior 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":7EA6
      Height          =   375
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":85B6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdPrimero 
      Appearance      =   0  'Flat
      DisabledPicture =   "botonAbmPrueba.frx":8CBA
      Height          =   375
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "botonAbmPrueba.frx":93CA
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
Attribute VB_Name = "botonAbmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
