VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PorcentajeComida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comidas"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "Porcentaje.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4500
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3735
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Porcentaje.frx":0ECA
         Height          =   2175
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BorderStyle     =   0
         HeadLines       =   3
         RowHeight       =   15
         RowDividerStyle =   0
         FormatLocked    =   -1  'True
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
         Caption         =   "Porcentaje de Kcal por comida"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "idTpoMenu"
            Caption         =   "idTpoMenu"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "DescTpoMenu"
            Caption         =   "Comida"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "proporcionRct"
            Caption         =   "Porcentaje (%)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   0
            BeginProperty Column00 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   2160
         TabIndex        =   4
         Top             =   2520
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2880
         TabIndex        =   3
         Top             =   2520
         Width           =   45
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
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
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
      Begin VB.PictureBox Pic_Cancelar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DrawMode        =   16  'Merge Pen
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         MouseIcon       =   "Porcentaje.frx":0EDF
         Picture         =   "Porcentaje.frx":1031
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   6
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
         MouseIcon       =   "Porcentaje.frx":1332
         Picture         =   "Porcentaje.frx":1484
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
         MouseIcon       =   "Porcentaje.frx":1740
         Picture         =   "Porcentaje.frx":1892
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
         MouseIcon       =   "Porcentaje.frx":1A26
         Picture         =   "Porcentaje.frx":1B78
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Appearance      =   0  'Flat
         DisabledPicture =   "Porcentaje.frx":1CD1
         Height          =   375
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Porcentaje.frx":1E65
         Picture         =   "Porcentaje.frx":1FB7
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancelar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         DisabledPicture =   "Porcentaje.frx":246A
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Porcentaje.frx":25C3
         Picture         =   "Porcentaje.frx":2715
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Aceptar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   375
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Detalles:"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "PorcentajeComida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
                            
Private Sub cmdAceptar_Click()
Dim strQuery As String
Sumador = 0

If estadoAbm = 3 Then

    If f_calculaSuma = 100 Then
        Unload Me
    Else
        Beep
        MsgBox "Debe corregir los porcentajes. El total debe sumar 100"
    End If

End If

estadoAbm = 1

End Sub

Private Sub cmdCancelar_Click()
Dim strMsg As String

strMsg = vbNo

If f_calculaSuma = 100 Then
    
    strMsg = MsgBox("¿Esta seguro que desea finalizar la operacion?", vbYesNo)
    
    If strMsg = vbYes Then
        
        estadoAbm = 1 ' el estado del form es "sin cambios"
           
        Unload Me
        
    End If

Else
    Beep
    MsgBox "Debe corregir los porcentajes. El total debe sumar 100"
End If

End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)

Label1.Caption = f_calculaSuma

End Sub

Private Sub DataGrid1_Change()

estadoAbm = 3

End Sub

Private Sub Form_Load()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 4500
Me.Width = 4590
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

estadoAbm = 1

Call f_CargarOrigenDatos

Label1.Caption = f_calculaSuma

Call f_Boton_Zorder

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call cmdAceptar_Click

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Adodc1.Recordset = Nothing

strQuery = "select * from TpoMenu"
Call f_Adodc_ConnectionString(Adodc1, strQuery)

End Sub

Function f_calculaSuma() As Long
Dim Sumador As Long
Sumador = 0

Adodc1.Recordset.MoveFirst
For i = 1 To 6
    Sumador = Sumador + DataGrid1.Columns(2).Value
    If Adodc1.Recordset.EOF = False Then
        Adodc1.Recordset.MoveNext
    End If
Next

If Sumador = 100 Then
    Label1.ForeColor = &HFF0000
Else
    Label1.ForeColor = &H80000012
End If

f_calculaSuma = Sumador

End Function

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

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Boton_Zorder

End Sub

Private Sub Pic_Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Aceptar

End Sub

Private Sub Pic_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call f_Cancelar

End Sub
