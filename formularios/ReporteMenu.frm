VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form ReporteMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador de  Planes Alimentarios"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "ReporteMenu.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5565
   Begin VB.Frame Frame2 
      Caption         =   "Opciones:"
      Height          =   2775
      Left            =   2640
      TabIndex        =   8
      Top             =   0
      Width           =   2895
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000005&
         Height          =   930
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   12
         ToolTipText     =   "Tildar para eliminar"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox ChckArchivo 
         Caption         =   "Guardar archivo"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "Guarda el informe para poder modificarlo"
         Top             =   600
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox ChckPantalla 
         Caption         =   "Informe por pantalla"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Sólo muestra el informe por pantalla"
         Top             =   240
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CommandButton cmd_Imprimir 
         Caption         =   "&Generar"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Imprimir"
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Lista de archivos existentes:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Para abrir el archivo hacer doble click"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   2670
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "db1nueva prueba anterior sin replica.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pacientes"
      Top             =   3480
      Visible         =   0   'False
      Width           =   11775
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "D:\Dietetica\rpts\informe1.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ReporteMenu.frx":0ECA
      Height          =   2535
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.Frame Frame1 
      Caption         =   "Ingresar Rango de Fechas"
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2415
      Begin MSComCtl2.DTPicker DTdesde 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66977793
         CurrentDate     =   37858
      End
      Begin MSComCtl2.DTPicker DThasta 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66977793
         CurrentDate     =   37860
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Desde:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Paciente:"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "ReporteMenu.frx":0EDF
         DataField       =   "Legajo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "nom"
         BoundColumn     =   "Legajo"
         Text            =   "DataCombo1"
      End
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   1680
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3000
      Top             =   120
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc4"
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
End
Attribute VB_Name = "ReporteMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command4_Click()

End Sub


Private Sub cmd_Imprimir_Click()
'Resets the value of all properties (except DataSource Property) to their default values.
CrystalReport1.Reset

CrystalReport1.ReportFileName = App_Path & "\rpts\rep_AdminMenu_one.rpt"
    
If ChckArchivo.Value = 1 And ChckPantalla.Value = 1 Then
    repArchivo
    repPantalla
Else
    If ChckArchivo.Value = 1 And ChckPantalla.Value = 0 Then
        repArchivo
    Else
        If ChckArchivo.Value = 0 And ChckPantalla.Value = 1 Then
            repPantalla
        Else
            If ChckArchivo.Value = 0 And ChckPantalla.Value = 0 Then
                MsgBox "Debe tildar al menos una de las dos opciones"
                ChckArchivo.Value = 1
                ChckPantalla.Value = 1
            End If
        End If
    End If
End If

End Sub

Private Sub DataCombo1_Click(Area As Integer)
BuscarInformes

End Sub

Private Sub DataCombo1_LostFocus()
If DataCombo1.Text = "" Then
    DataCombo1.SetFocus
    MsgBox "Debe Completar el Nombre del Paciente", vbInformation, "Información"
End If



End Sub

Private Sub Form_Activate()
BuscarInformes
End Sub

Private Sub Form_Load()

'Data1.DatabaseName = Lugar

'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 3195
Me.Width = 5655
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

Call f_CargarOrigenDatos

DTdesde.Value = Now - 7
DThasta.Value = Now

BuscarInformes
End Sub

Private Sub List1_DblClick()
Dim sFile As String
ArmaArchivo = Split(List1.Text, " ")
ArmaArchivo1 = Split(ArmaArchivo(1), "/")

sFile = App.Path & "\Reportes\Rep_" & DataCombo1.BoundText & "_" & ArmaArchivo1(2) & "_" & ArmaArchivo1(1) & "_" & ArmaArchivo1(0) & "_.doc"

'MDIForm1.mnuInforme_Click

frm_ole_document.Show

If MDIForm1.ActiveForm Is Nothing Then MDIForm1.LoadNewDoc
  
'MDIForm1.ActiveForm.rtfText.LoadFile sFile
MDIForm1.ActiveForm.rtfText.CreateLink (sFile)
MDIForm1.ActiveForm.Caption = sFile

MsgBox "Para modificar el archivo actual debe hacer doble click sobre el mismo." & vbCrLf & vbTab & " - Debe tener instalado Microsoft Word" & vbCrLf & vbTab & " de lo contrario se producira un error", vbInformation

End Sub
Sub repArchivo()
Dim StringDir As String
Dim strQuery As String

''CrystalReport1.SelectionFormula = " {platosmenu.legajo} = " & DataCombo1.BoundText & " and {platosmenu.fechaMenu} in Date(" & Year(DTdesde.Value) & ", " & Month(DTdesde.Value) & ", " & Day(DTdesde.Value) & ") to Date(" & Year(DThasta.Value) & ", " & Month(DThasta.Value) & ", " & Day(DThasta.Value) & ") "
''CrystalReport1.Destination = crptToFile
''CrystalReport1.PrintFileType = crptRTF

strQuery = " {platosmenu.legajo} = " & DataCombo1.BoundText & " and {platosmenu.fechaMenu} in Date(" & Year(DTdesde.Value) & ", " & Month(DTdesde.Value) & ", " & Day(DTdesde.Value) & ") to Date(" & Year(DThasta.Value) & ", " & Month(DThasta.Value) & ", " & Day(DThasta.Value) & ") "

'CrystalReport1.PrintFileType = crptRTF

CrystalReport1.PrintFileType = crptWinWord

''StringDir = Dir$(App_Path & "\Reportes\Rep_" & DataCombo1.BoundText & "_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_.rtf")
''
''If StringDir = "Rep_" & DataCombo1.BoundText & "_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_.rtf" Then
''    Kill App_Path & "\Reportes\Rep_" & DataCombo1.BoundText & "_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_.rtf"
''End If
''
''CrystalReport1.PrintFileName = App_Path & "\Reportes\Rep_" & DataCombo1.BoundText & "_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_.rtf"

StringDir = Dir$(App_Path & "\Reportes\Rep_" & DataCombo1.BoundText & "_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_.doc")

If StringDir = "Rep_" & DataCombo1.BoundText & "_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_.doc" Then
    Kill App_Path & "\Reportes\Rep_" & DataCombo1.BoundText & "_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_.doc"
End If

CrystalReport1.PrintFileName = App_Path & "\Reportes\Rep_" & DataCombo1.BoundText & "_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_.doc"

Call f_print(CrystalReport1, strQuery, crptToFile)

Call BuscarInformes

MsgBox "El archivo se ha generado con la fecha de hoy y agregado a la lista de archivos existentes con éxito.", vbInformation

End Sub

Sub repPantalla()

Dim FecRep() As String
Dim strQuery As String

strQuery = " {platosmenu.legajo} = " & DataCombo1.BoundText & " and {platosmenu.fechaMenu} in Date(" & Year(DTdesde.Value) & ", " & Month(DTdesde.Value) & ", " & Day(DTdesde.Value) & ") to Date(" & Year(DThasta.Value) & ", " & Month(DThasta.Value) & ", " & Day(DThasta.Value) & ") "

Call f_print(CrystalReport1, strQuery, crptToWindow)

End Sub
Public Sub BuscarInformes()
Dim FecRep() As String

List1.Clear

StringDir = Dir$(App_Path & "\Reportes\Rep_" & DataCombo1.BoundText & "_*")

While StringDir <> ""
   
    FecRep = Split(StringDir, "_")
    año = FecRep(2)
    mes = FecRep(3)
    dia = FecRep(4)
    List1.AddItem "Informe " & dia & "/" & mes & "/" & año 'StringDir
    StringDir = Dir$
Wend

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim band As Integer

If Button = 2 Then
    band = 0
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            band = 1
        End If
    Next
           
    If band = 0 Then
        MDIForm1.mnuEli.Enabled = False
    Else
        MDIForm1.mnuEli.Enabled = True
    End If
    
    PopupMenu MDIForm1.mnuEliminar, vbPopupMenuLeftAlign
End If

End Sub

Sub f_CargarOrigenDatos()
Dim strQuery As String
strQuery = ""

Set Me.Data1.Recordset = Nothing
Set Me.Adodc.Recordset = Nothing
Set Me.Adodc4.Recordset = Nothing

strQuery = "select * from Pacientes"
Call f_Data_DatabaseName(Data1, strQuery)

strQuery = "select *, (apell & ', ' & nombre) as nom from pacientes order by apell, nombre"
Call f_Adodc_ConnectionString(Adodc, strQuery)

strQuery = "select * from menu"
Call f_Adodc_ConnectionString(Adodc4, strQuery)

'Define propiedades de los controles enlazados
Call f_Enlaza_ControlData(DataCombo1, Data1, Adodc, "Legajo", "Legajo", "nom")

'==============================================

End Sub
