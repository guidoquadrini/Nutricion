VERSION 5.00
Begin VB.UserControl Cls_Button_ABM_Comportamiento 
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   720
   ScaleWidth      =   5805
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdImprimir 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":0000
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":0710
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Imprimir"
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":0E14
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":0F66
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   10
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":1421
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":1573
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   9
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":19E1
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":1B33
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   8
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":1E10
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":1F62
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   7
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":23D7
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":2529
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   6
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":29F4
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":2B46
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   5
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":2F80
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":30D2
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   4
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":3261
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":33B3
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   3
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":3626
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":3778
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   2
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox Pic_Cancelar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5160
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":3A34
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":3B86
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   1
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":3E87
         Height          =   375
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":401B
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":416D
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancelar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":4620
         Height          =   375
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":4779
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":48CB
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Aceptar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":4B87
         Height          =   375
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":4CA8
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":4DFA
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Modificar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBorrar 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":506D
         Height          =   375
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":5183
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":52D5
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Eliminar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAgregar 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":5464
         Height          =   375
         Left            =   3120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":55B1
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":5703
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Agregar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdUltimo 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":5B3D
         Height          =   375
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":5CE5
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":5E37
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Ultimo"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdSiguiente 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":6302
         Height          =   375
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":646F
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":65C1
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Siguiente"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":6A36
         Height          =   375
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":6BBE
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":6D10
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Buscar"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAnterior 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":6FED
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":7157
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":72A9
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Anterior"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdPrimero 
         Appearance      =   0  'Flat
         DisabledPicture =   "Cls_Button_ABM_Comportamiento.ctx":7717
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":78BC
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":7A0E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Primero"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox Pic_Primero_Gris 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":7EC9
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":801B
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   11
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":81C0
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":8312
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   12
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":847C
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":85CE
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   13
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":8756
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":88A8
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   14
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":8A15
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":8B67
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   15
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":8D0F
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":8E61
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   16
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":8FAE
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":9100
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   17
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":9216
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":9368
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":9489
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":95DB
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   19
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
         MouseIcon       =   "Cls_Button_ABM_Comportamiento.ctx":9734
         Picture         =   "Cls_Button_ABM_Comportamiento.ctx":9886
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   20
         Top             =   120
         Width           =   375
      End
   End
End
Attribute VB_Name = "Cls_Button_ABM_Comportamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Sub cmdImprimir_Click()

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call f_Boton_Zorder

End Sub

Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAceptar.ZOrder 0

cmdPrimero.ZOrder 1
cmdAnterior.ZOrder 1
cmdBuscar.ZOrder 1
cmdSiguiente.ZOrder 1
cmdUltimo.ZOrder 1
cmdAgregar.ZOrder 1
cmdBorrar.ZOrder 1
cmdModificar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Agregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAgregar.ZOrder 0

cmdPrimero.ZOrder 1
cmdAnterior.ZOrder 1
cmdBuscar.ZOrder 1
cmdSiguiente.ZOrder 1
cmdUltimo.ZOrder 1
cmdBorrar.ZOrder 1
cmdModificar.ZOrder 1
cmdAceptar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Anterior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAnterior.ZOrder 0

cmdPrimero.ZOrder 1
cmdBuscar.ZOrder 1
cmdSiguiente.ZOrder 1
cmdUltimo.ZOrder 1
cmdAgregar.ZOrder 1
cmdBorrar.ZOrder 1
cmdModificar.ZOrder 1
cmdAceptar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Borrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdBorrar.ZOrder 0

cmdPrimero.ZOrder 1
cmdAnterior.ZOrder 1
cmdBuscar.ZOrder 1
cmdSiguiente.ZOrder 1
cmdUltimo.ZOrder 1
cmdAgregar.ZOrder 1
cmdModificar.ZOrder 1
cmdAceptar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Buscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdBuscar.ZOrder 0

cmdPrimero.ZOrder 1
cmdAnterior.ZOrder 1
cmdSiguiente.ZOrder 1
cmdUltimo.ZOrder 1
cmdAgregar.ZOrder 1
cmdBorrar.ZOrder 1
cmdModificar.ZOrder 1
cmdAceptar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub





Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancelar.ZOrder 0

cmdPrimero.ZOrder 1
cmdAnterior.ZOrder 1
cmdBuscar.ZOrder 1
cmdSiguiente.ZOrder 1
cmdUltimo.ZOrder 1
cmdAgregar.ZOrder 1
cmdBorrar.ZOrder 1
cmdModificar.ZOrder 1
cmdAceptar.ZOrder 1

End Sub

Private Sub Pic_Modificar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdModificar.ZOrder 0

cmdPrimero.ZOrder 1
cmdAnterior.ZOrder 1
cmdBuscar.ZOrder 1
cmdSiguiente.ZOrder 1
cmdUltimo.ZOrder 1
cmdAgregar.ZOrder 1
cmdBorrar.ZOrder 1
cmdAceptar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Primero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPrimero.ZOrder 0

cmdAnterior.ZOrder 1
cmdBuscar.ZOrder 1
cmdSiguiente.ZOrder 1
cmdUltimo.ZOrder 1
cmdAgregar.ZOrder 1
cmdBorrar.ZOrder 1
cmdModificar.ZOrder 1
cmdAceptar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Siguiente_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSiguiente.ZOrder 0

cmdPrimero.ZOrder 1
cmdAnterior.ZOrder 1
cmdBuscar.ZOrder 1
cmdUltimo.ZOrder 1
cmdAgregar.ZOrder 1
cmdBorrar.ZOrder 1
cmdModificar.ZOrder 1
cmdAceptar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub

Private Sub Pic_Ultimo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdUltimo.ZOrder 0

cmdPrimero.ZOrder 1
cmdAnterior.ZOrder 1
cmdBuscar.ZOrder 1
cmdSiguiente.ZOrder 1
cmdAgregar.ZOrder 1
cmdBorrar.ZOrder 1
cmdModificar.ZOrder 1
cmdAceptar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub

Sub f_Boton_Zorder()

cmdPrimero.ZOrder 1
cmdAnterior.ZOrder 1
cmdBuscar.ZOrder 1
cmdSiguiente.ZOrder 1
cmdUltimo.ZOrder 1
cmdAgregar.ZOrder 1
cmdBorrar.ZOrder 1
cmdModificar.ZOrder 1
cmdAceptar.ZOrder 1
cmdCancelar.ZOrder 1

End Sub

Private Sub UserControl_Initialize()

'f_Boton_Zorder

End Sub
