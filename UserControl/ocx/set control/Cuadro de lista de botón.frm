VERSION 5.00
Begin VB.Form frmListButtons 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   2880
   ClientTop       =   3210
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   5865
   Begin VB.CommandButton cmdPrimero 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":0000
      Height          =   375
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":0710
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAnterior 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":0E14
      Height          =   375
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":1524
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBuscar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":1C28
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":2338
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdSiguiente 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":2A3C
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":314C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdUltimo 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":3850
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":3F60
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAgregar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":4664
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":4D74
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdBorrar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":5478
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":5B88
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdModificar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":628C
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":698E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":7092
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":77A2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":7EA6
      Height          =   375
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":85B6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimir 
      Appearance      =   0  'Flat
      DisabledPicture =   "Cuadro de lista de botón.frx":8CBA
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Cuadro de lista de botón.frx":93CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.ListBox lstItems 
      DragIcon        =   "Cuadro de lista de botón.frx":9ACE
      Height          =   2895
      IntegralHeight  =   0   'False
      Left            =   330
      TabIndex        =   4
      Top             =   765
      Width           =   2280
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2670
      Picture         =   "Cuadro de lista de botón.frx":9F10
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "5011"
      Top             =   1815
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2670
      Picture         =   "Cuadro de lista de botón.frx":A012
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "5012"
      Top             =   3735
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdDelete 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2670
      Picture         =   "Cuadro de lista de botón.frx":A114
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "5010"
      Top             =   1335
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   330
      Left            =   2670
      Picture         =   "Cuadro de lista de botón.frx":A216
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "5009"
      Top             =   855
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.Label ContenedorBotones 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   5745
   End
End
Attribute VB_Name = "frmListButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
  Dim sTmp As String
  sTmp = InputBox("Escriba el nuevo elemento que desea agregar:")
  If Len(sTmp) = 0 Then Exit Sub
  lstItems.AddItem sTmp
End Sub

Private Sub cmdDelete_Click()
  If lstItems.ListIndex > -1 Then
    If MsgBox("Eliminar '" & lstItems.Text & "'?", vbQuestion + vbYesNo) = vbYes Then
      lstItems.RemoveItem lstItems.ListIndex
    End If
  End If
End Sub

Private Sub cmdUp_Click()
  On Error Resume Next
  Dim nItem As Integer
  
  With lstItems
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'no se puede subir el primer elemento
    'subir elemento
    .AddItem .Text, nItem - 1
    'quitar elemento antiguo
    .RemoveItem nItem + 1
    'seleccionar el elemento que se ha movido
    .Selected(nItem - 1) = True
  End With
End Sub

Private Sub cmdDown_Click()
  On Error Resume Next
  Dim nItem As Integer
  
  With lstItems
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'no se puede bajar el último elemento
    'bajar elemento
    .AddItem .Text, nItem + 2
    'quitar elemento antiguo
    .RemoveItem nItem
    'seleccionar el elemento que se ha movido
    .Selected(nItem + 1) = True
  End With
End Sub

Private Sub Form_Load()

End Sub

Private Sub lstItems_DragDrop(Source As Control, X As Single, Y As Single)
  Dim i As Integer
  Dim nID As Integer
  Dim sTmp As String
  
  If Source.Name <> "lstItems" Then Exit Sub
  If lstItems.ListCount = 0 Then Exit Sub

  With lstItems
    i = (Y \ TextHeight("A")) + .TopIndex
    If i = .ListIndex Then
      'colocar encima de sí mismo
      Exit Sub
    End If
    If i > .ListCount - 1 Then i = .ListCount - 1
    nID = .ListIndex
    sTmp = .Text
    If (nID > -1) Then
      sTmp = .Text
    .RemoveItem nID
    .AddItem sTmp, i
    .ListIndex = .NewIndex
    End If
  End With
  SetListButtons
End Sub

Sub lstItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then lstItems.Drag
End Sub

Private Sub lstItems_Click()
  SetListButtons
End Sub

Sub SetListButtons()
  Dim i As Integer
  i = lstItems.ListIndex
  'establecer el estado de los botones de desplazamiento
  cmdUp.Enabled = (i > 0)
  cmdDown.Enabled = ((i > -1) And (i < (lstItems.ListCount - 1)))
  cmdDelete.Enabled = (i > -1)
  ContenedorBotones.Enabled = True
End Sub

