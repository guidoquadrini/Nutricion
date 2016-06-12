VERSION 5.00
Begin VB.Form frm_formulaDesarrollada_Print 
   Caption         =   "Opciones - Imprimir"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2565
   Icon            =   "frm_formulaDesarrollada_Print.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   2565
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   735
         Left            =   735
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
         Begin VB.CommandButton cmdAceptar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_formulaDesarrollada_Print.frx":0ECA
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_formulaDesarrollada_Print.frx":1023
            Picture         =   "frm_formulaDesarrollada_Print.frx":1175
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Aceptar"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton cmdCancelar 
            Appearance      =   0  'Flat
            DisabledPicture =   "frm_formulaDesarrollada_Print.frx":1431
            Height          =   375
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            MouseIcon       =   "frm_formulaDesarrollada_Print.frx":15C5
            Picture         =   "frm_formulaDesarrollada_Print.frx":1717
            Style           =   1  'Graphical
            TabIndex        =   9
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
            Left            =   120
            MouseIcon       =   "frm_formulaDesarrollada_Print.frx":1BCA
            Picture         =   "frm_formulaDesarrollada_Print.frx":1D1C
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   8
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
            MouseIcon       =   "frm_formulaDesarrollada_Print.frx":1E75
            Picture         =   "frm_formulaDesarrollada_Print.frx":1FC7
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   7
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox Pic_Aceptar 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DrawMode        =   16  'Merge Pen
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            MouseIcon       =   "frm_formulaDesarrollada_Print.frx":215B
            Picture         =   "frm_formulaDesarrollada_Print.frx":22AD
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   6
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
            MouseIcon       =   "frm_formulaDesarrollada_Print.frx":2569
            Picture         =   "frm_formulaDesarrollada_Print.frx":26BB
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.OptionButton opt_graficos 
         Caption         =   "Graficos comparativos"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton opt_descripcion 
         Caption         =   "Informe con descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton opt_cantidades 
         Caption         =   "Informe de  cantidades "
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frm_formulaDesarrollada_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub f_Boton_Zorder()

If Me.cmdCancelar.Enabled = True Then
    Me.Pic_Cancelar.ZOrder 0
Else
    Me.Pic_Cancelar_Gris.ZOrder 0
End If

If Me.cmdAceptar.Enabled = True Then
    Me.Pic_Aceptar.ZOrder 0
Else
    Me.Pic_Aceptar_Gris.ZOrder 0
End If

Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 1

End Sub

Sub f_Aceptar()

Me.cmdAceptar.ZOrder 0
Me.cmdCancelar.ZOrder 1

End Sub

Sub f_Cancelar()

Me.cmdAceptar.ZOrder 1
Me.cmdCancelar.ZOrder 0

End Sub

Private Sub cmdAceptar_Click()

Dim strquery, sMsg As String

'Resets the value of all properties (except DataSource Property) to their default values.
frm_formulaDesarrollada.CrystalReport1.Reset
 
If opt_cantidades.Value = True Then

    frm_formulaDesarrollada.CrystalReport1.ReportFileName = App_Path & "\rpts\rep_formdesarrollada_one.rpt"
               
    strquery = " {pacientes.legajo} = " & frm_formulaDesarrollada.DataCombo1.BoundText
            
Else
    
    If opt_descripcion.Value = True Then
    
        frm_formulaDesarrollada.CrystalReport1.ReportFileName = App_Path & "\rpts\rep_formdesarrolladaSinCant_one.rpt"
                       
        strquery = " {pacientes.legajo} = " & frm_formulaDesarrollada.DataCombo1.BoundText
        
    Else
        
        If opt_graficos.Value = True Then
        
            frm_formulaDesarrollada.CrystalReport1.ReportFileName = App_Path & "\rpts\rep_graf_macronutrientes_one.rpt"
                       
            strquery = " {pacientes.legajo} = " & frm_formulaDesarrollada.DataCombo1.BoundText '& " and {csl_macronutrientes_fs.legajo} = " & frm_formulaDesarrollada.DataCombo1.BoundText
            '& " and {csl_macronutrientes_fd.legajo} = " & frm_formulaDesarrollada.DataCombo1.BoundText & " AND {csl_macronutrientes_graf_linea.legajo} = " & frm_formulaDesarrollada.DataCombo1.BoundText
            
        End If
    
    End If
    
End If


Call f_print(frm_formulaDesarrollada.CrystalReport1, strquery, crptToWindow)


End Sub

Private Sub cmdCancelar_Click()

Unload Me

End Sub

Private Sub Form_Load()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 2625
Me.Width = 2685
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Aceptar

End Sub

Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cancelar

End Sub

