VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_ModoAtencion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Omnia - Modo de Acceso"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frm_ModoAtencion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fme_ModoAtencion 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Frame fme_main 
         Height          =   4695
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7215
         Begin VB.Frame fme_ConsultasExistentes 
            Caption         =   "Consultas Existentes"
            Height          =   3255
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   6735
            Begin MSDataGridLib.DataGrid grd_ConsultasExistentes 
               Height          =   2895
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   5106
               _Version        =   393216
               AllowUpdate     =   0   'False
               ColumnHeaders   =   -1  'True
               HeadLines       =   1
               RowHeight       =   15
               RowDividerStyle =   6
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
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
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
                  DataField       =   ""
                  Caption         =   ""
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
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Frame fme_ConsultaNueva 
            Caption         =   "Seleccionar paciente"
            Height          =   3255
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   6735
            Begin VB.Timer Timer1 
               Enabled         =   0   'False
               Interval        =   100
               Left            =   6840
               Top             =   120
            End
            Begin VB.Frame Frame 
               BorderStyle     =   0  'None
               Caption         =   "Paciente:"
               Height          =   735
               Index           =   0
               Left            =   1200
               TabIndex        =   4
               Top             =   720
               Width           =   4455
               Begin VB.Frame Frame10 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame10"
                  Height          =   495
                  Left            =   3795
                  TabIndex        =   5
                  Top             =   120
                  Width           =   615
                  Begin VB.CommandButton cmd_Tipito 
                     Appearance      =   0  'Flat
                     DisabledPicture =   "frm_ModoAtencion.frx":0ECA
                     Height          =   315
                     Left            =   120
                     MaskColor       =   &H00FFFFFF&
                     Picture         =   "frm_ModoAtencion.frx":15DA
                     Style           =   1  'Graphical
                     TabIndex        =   8
                     ToolTipText     =   "Info"
                     Top             =   120
                     UseMaskColor    =   -1  'True
                     Width           =   315
                  End
                  Begin VB.PictureBox Pic_Tipito_Gris 
                     Appearance      =   0  'Flat
                     BorderStyle     =   0  'None
                     DrawMode        =   16  'Merge Pen
                     ForeColor       =   &H80000008&
                     Height          =   315
                     Left            =   120
                     MouseIcon       =   "frm_ModoAtencion.frx":186A
                     Picture         =   "frm_ModoAtencion.frx":19BC
                     ScaleHeight     =   315
                     ScaleWidth      =   315
                     TabIndex        =   7
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
                     MouseIcon       =   "frm_ModoAtencion.frx":1AEC
                     Picture         =   "frm_ModoAtencion.frx":1C3E
                     ScaleHeight     =   315
                     ScaleWidth      =   315
                     TabIndex        =   6
                     Top             =   120
                     Width           =   315
                  End
               End
               Begin MSDataListLib.DataCombo cbo_Pacientes 
                  Height          =   315
                  Left            =   195
                  TabIndex        =   9
                  Top             =   240
                  Width           =   3615
                  _ExtentX        =   6376
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  ListField       =   ""
                  BoundColumn     =   ""
                  Text            =   ""
               End
            End
            Begin VB.Label lbl_nroDoc 
               Caption         =   "lbl_nroDoc"
               Height          =   195
               Left            =   3600
               TabIndex        =   31
               Top             =   2040
               Width           =   1305
            End
            Begin VB.Label lbl_Telefono 
               Caption         =   "lbl_Telefono"
               Height          =   195
               Left            =   2640
               TabIndex        =   30
               Top             =   2520
               Width           =   2385
            End
            Begin VB.Label lbl_Direccion 
               Caption         =   "lbl_Direccion"
               Height          =   195
               Left            =   2640
               TabIndex        =   29
               Top             =   2280
               Width           =   2385
            End
            Begin VB.Label lbl_tpoDoc 
               Caption         =   "lbl_tpoDoc"
               Height          =   195
               Left            =   2640
               TabIndex        =   28
               Top             =   2040
               Width           =   825
            End
            Begin VB.Label lbl_HC 
               Caption         =   "lbl_HC"
               Height          =   195
               Left            =   2640
               TabIndex        =   27
               Top             =   1800
               Width           =   2385
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Historia Clinica:"
               Height          =   195
               Left            =   600
               TabIndex        =   26
               Top             =   1800
               Width           =   1080
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Telefono:"
               Height          =   195
               Left            =   600
               TabIndex        =   25
               Top             =   2520
               Width           =   675
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Direccion:"
               Height          =   195
               Left            =   600
               TabIndex        =   24
               Top             =   2280
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tpo. y nro. Documento:"
               Height          =   195
               Left            =   600
               TabIndex        =   23
               Top             =   2040
               Width           =   1680
            End
            Begin VB.Label lbl_hora 
               AutoSize        =   -1  'True
               Caption         =   "10:00:00"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "HH:mm"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   4
               EndProperty
               Height          =   195
               Left            =   3120
               TabIndex        =   13
               Top             =   360
               Width           =   630
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
               Height          =   195
               Left            =   2640
               TabIndex        =   12
               Top             =   360
               Width           =   390
            End
            Begin VB.Label lbl_fechaConsulta 
               AutoSize        =   -1  'True
               Caption         =   "01/01/2005"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   3
               EndProperty
               Height          =   195
               Left            =   1200
               TabIndex        =   11
               Top             =   360
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fecha:"
               Height          =   195
               Left            =   600
               TabIndex        =   10
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   495
            Left            =   3008
            TabIndex        =   14
            Top             =   4080
            Width           =   1215
            Begin VB.CommandButton cmd_Cancelar 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_ModoAtencion.frx":1ECE
               Height          =   375
               Left            =   600
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frm_ModoAtencion.frx":2062
               Picture         =   "frm_ModoAtencion.frx":21B4
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Cancelar"
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.CommandButton cmd_Aceptar 
               Appearance      =   0  'Flat
               DisabledPicture =   "frm_ModoAtencion.frx":2667
               Height          =   375
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "frm_ModoAtencion.frx":27C0
               Picture         =   "frm_ModoAtencion.frx":2912
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Aceptar"
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
               Left            =   600
               MouseIcon       =   "frm_ModoAtencion.frx":2BCE
               Picture         =   "frm_ModoAtencion.frx":2D20
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   19
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Aceptar 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               MouseIcon       =   "frm_ModoAtencion.frx":3021
               Picture         =   "frm_ModoAtencion.frx":3173
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   20
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Cancelar_Gris 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   600
               MouseIcon       =   "frm_ModoAtencion.frx":342F
               Picture         =   "frm_ModoAtencion.frx":3581
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   17
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Pic_Aceptar_Gris 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DrawMode        =   16  'Merge Pen
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               MouseIcon       =   "frm_ModoAtencion.frx":3715
               Picture         =   "frm_ModoAtencion.frx":3867
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   18
               Top             =   120
               Width           =   375
            End
         End
         Begin MSComctlLib.TabStrip TabStrip1 
            Height          =   3735
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   6588
            ShowTips        =   0   'False
            HotTracking     =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Nueva Consulta"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "Consultas Existentes"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frm_ModoAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb_cbo_Pacientes As ADODB.Recordset
Dim tb_grd_ConsultasExistentes As ADODB.Recordset

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

'=========================================
strQuery = "select *, (apell & ', ' & nombre) as nom from pacientes order by apell, nombre"

Set tb_cbo_Pacientes = f_StaticRecordset(adCmdText, strQuery)
'=========================================
'=========================================
strQuery = "[csl_ConsultasExistentes]"

Set tb_grd_ConsultasExistentes = f_StaticRecordset(adCmdTable, strQuery)
Set grd_ConsultasExistentes.DataSource = tb_grd_ConsultasExistentes
'=========================================

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(cbo_Pacientes, tb_cbo_Pacientes, tb_cbo_Pacientes, "Legajo", "Legajo", "nom")

'enlaza los lbls de datos de pacientes
Call cbo_Pacientes_Click(1)

End Sub



Private Sub cbo_Pacientes_Click(Area As Integer)
Dim tb_lbls_Pacientes As ADODB.Recordset
Set tb_lbls_Pacientes = New ADODB.Recordset

'=========================================
strQuery = "SELECT Legajo, pac_tpoDoc, pac_nroDoc, Dir, Tel FROM pacientes WHERE legajo = " & cbo_Pacientes.BoundText

Set tb_lbls_Pacientes = f_StaticRecordset(adCmdText, strQuery)
'=========================================

If tb_lbls_Pacientes.RecordCount > 0 Then

    '=========Enlaza los lbls de datos de pacientes===============
    Set Me.lbl_HC.DataSource = tb_lbls_Pacientes
    Me.lbl_HC.DataField = "Legajo"
    
    Set Me.lbl_tpoDoc.DataSource = tb_lbls_Pacientes
    Me.lbl_tpoDoc.DataField = "pac_tpoDoc"
    
    Set Me.lbl_nroDoc.DataSource = tb_lbls_Pacientes
    Me.lbl_nroDoc.DataField = "pac_nroDoc"
    
    Set Me.lbl_Direccion.DataSource = tb_lbls_Pacientes
    Me.lbl_Direccion.DataField = "Dir"
    
    Set Me.lbl_Telefono.DataSource = tb_lbls_Pacientes
    Me.lbl_Telefono.DataField = "Tel"
    '=============================================================

End If

Set tb_lbls_Pacientes = Nothing

End Sub

Private Sub cmd_aceptar_Click()

Dim sTabStripIndex As String
sTabStripIndex = TabStrip1.SelectedItem.Index

Select Case sTabStripIndex
    
    Case Is = 1 'Nueva Consulta
        Dim strQuery As String
        Dim nLegajo As Long
        Dim dFecha As Date
        Dim dHora As Date
                
        dFecha = Format(Now, "Short Date")
        dHora = Time
        
        If cbo_Pacientes.Text <> "" Then
            
            nLegajo = cbo_Pacientes.BoundText
            
            '=======Inserta nueva consulta=====================
            strQuery = "INSERT INTO consultas (atc_Legajo, atc_fecha, atc_Hora, atc_codprf) VALUES(" & nLegajo & ", '" & dFecha & "', '" & dHora & "', " & nUserLoging & ")"
        
            Call f_UpdateInsertRecordset(adCmdText, strQuery)
            '==================================================
            
            '=======Recupera el codigo de consulta recien generado=======
            Dim tb_Consulta As ADODB.Recordset
            Set tb_Consulta = New ADODB.Recordset
            
            strQuery = "SELECT atc_codigo FROM Consultas WHERE atc_legajo = " & nLegajo & " AND atc_fecha = #" & Month(dFecha) & "/" & Day(dFecha) & "/" & Year(dFecha) & "# AND atc_Hora = #" & dHora & "# AND atc_codprf = " & nUserLoging
            
            Set tb_Consulta = f_StaticRecordset(adCmdText, strQuery)
            
            nCodAtencion = tb_Consulta.Fields("atc_codigo").Value
            
            Set tb_Consulta = Nothing
            '==================================================
            
            Me.Hide
            
        Else
        
            MsgBox "Debe seleccionar un paciente", vbInformation
            
        End If
        
    Case Is = 2 'Consulta Existente
                
        If tb_grd_ConsultasExistentes.RecordCount > 0 Then
            'seleccionando la consulta
            
            nCodAtencion = tb_grd_ConsultasExistentes.Fields("atc_codigo").Value
                                   
            Me.Hide
            
        Else
        
            MsgBox "No hay ninguna consulta seleccionada.", vbInformation
        
        End If
        
End Select

Call f_GetPermisos

If nCodAtencion > 0 Then
       
    'Titulo Main ===============================
    Select Case sUserNivel
               
        Case Is = "Nutricionista"
            MDIForm1.Caption = "Omnia - Usuario: " & f_GetUserLoging(nUserLoging) & " - Modo Consulta"
        
        Case Else
            MDIForm1.Caption = "Omnia - Usuario: " & sUserNivel & " - Modo Consulta"
    
    End Select
    '===========================================
            
End If

End Sub

Private Sub cmd_Cancelar_Click()
'se ACCEDE en modo Standard
nCodAtencion = 0

Call f_GetPermisos

'Titulo Main ===============================
Select Case sUserNivel
           
    Case Is = "Nutricionista"
        MDIForm1.Caption = "Omnia - Usuario: " & f_GetUserLoging(nUserLoging)
    
    Case Else
        MDIForm1.Caption = "Omnia - Usuario: " & sUserNivel

End Select
'===========================================
    
Me.Hide

End Sub



Private Sub Form_Load()
'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 5085
Me.Width = 7320
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

'permisos ==================================
Select Case sUserNivel
    Case Is = "Administrador" 'habilito Modo atencion
        Me.cmd_Aceptar.Enabled = True
        
    Case Is = "Nutricionista" 'habilito Modo atencion
        Me.cmd_Aceptar.Enabled = True
    
    Case Else 'deshabilito Modo atencion
        Me.cmd_Aceptar.Enabled = False

End Select
'===========================================
    
Call f_CargarOrigenDatos

Call TabStrip1_Click

Call f_DatagridProperties

Me.lbl_fechaConsulta = Format(Now, "short date")

Timer1.Enabled = True
Me.lbl_hora = Time

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set tb_cbo_Pacientes = Nothing
Set tb_grd_ConsultasExistentes = Nothing

Timer1.Enabled = False

End Sub

Sub f_Boton_Zorder()

If Me.cmd_Cancelar.Enabled = True Then
    Me.Pic_Cancelar.ZOrder 0
Else
    Me.Pic_Cancelar_Gris.ZOrder 0
End If

If Me.cmd_Aceptar.Enabled = True Then
    Me.Pic_Aceptar.ZOrder 0
Else
    Me.Pic_Aceptar_Gris.ZOrder 0
End If

Me.cmd_Aceptar.ZOrder 1
Me.cmd_Cancelar.ZOrder 1

End Sub

Sub f_Aceptar()

Me.cmd_Aceptar.ZOrder 0
Me.cmd_Cancelar.ZOrder 1

End Sub

Sub f_Cancelar()

Me.cmd_Aceptar.ZOrder 1
Me.cmd_Cancelar.ZOrder 0

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


Private Sub TabStrip1_Click()

'define los valores de los frame correspondientes para que funcionen con el tabstrip
Dim a As String
a = TabStrip1.SelectedItem.Index

Select Case a
    Case Is = 1
    
        Me.fme_ConsultaNueva.ZOrder 0
        Me.fme_ConsultasExistentes.ZOrder 1
        
    Case Is = 2
    
        Me.fme_ConsultaNueva.ZOrder 1
        Me.fme_ConsultasExistentes.ZOrder 0
        
End Select

End Sub

Private Sub Timer1_Timer()
Me.lbl_hora = Time
End Sub

Sub f_DatagridProperties()

Me.grd_ConsultasExistentes.MarqueeStyle = dbgHighlightRow

Me.grd_ConsultasExistentes.Columns("atc_codigo").Width = 750.0473
Me.grd_ConsultasExistentes.Columns("atc_codigo").Visible = False

Me.grd_ConsultasExistentes.Columns("HC").Width = 500

Me.grd_ConsultasExistentes.Columns("Paciente").Width = 1915.0237

Me.grd_ConsultasExistentes.Columns("Fecha").Width = 1000
Me.grd_ConsultasExistentes.Columns("Fecha").NumberFormat = "dd/MM/yyyy"

Me.grd_ConsultasExistentes.Columns("Hs").Width = 750.0473
Me.grd_ConsultasExistentes.Columns("Hs").NumberFormat = "HH:mm:ss"

Me.grd_ConsultasExistentes.Columns("prf_codigo").Width = 750.0473
Me.grd_ConsultasExistentes.Columns("prf_codigo").Visible = False

Me.grd_ConsultasExistentes.Columns("Profesional").Width = 1915.0237
'Me.grd_ConsultasExistentes.Columns("Profesional").Visible = False

End Sub
