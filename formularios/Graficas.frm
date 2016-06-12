VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Graficas 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   240
      TabIndex        =   20
      Top             =   4440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3413
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   16744576
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "Graficas.frx":0000
      DataField       =   "Legajo"
      DataSource      =   "Data1"
      Height          =   2400
      Left            =   6480
      TabIndex        =   22
      Top             =   3480
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4233
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16744576
      ForeColor       =   16777215
      ListField       =   "nom"
      BoundColumn     =   "Legajo"
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   3840
      TabIndex        =   21
      Top             =   3480
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      MonthBackColor  =   16744576
      StartOfWeek     =   66977793
      TitleBackColor  =   8388608
      TitleForeColor  =   16777215
      TrailingForeColor=   4194304
      CurrentDate     =   38011
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   240
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Graficas.frx":0014
      Height          =   1935
      Left            =   4200
      TabIndex        =   18
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "legajo"
         Caption         =   "legajo"
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
         DataField       =   "Fecha de Consulta"
         Caption         =   "Fecha de Consulta"
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
         DataField       =   "Peso"
         Caption         =   "Peso"
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
         BeginProperty Column00 
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6960
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
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
      OLEDBFile       =   ""
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
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "db1nueva prueba anterior sin replica.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pesos"
      Top             =   1080
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9960
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   855
      Left            =   0
      OleObjectBlob   =   "Graficas.frx":0029
      TabIndex        =   17
      Top             =   3480
      Width           =   855
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "Graficas.frx":1D51
      TabIndex        =   16
      Top             =   2880
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Dietetica\db1nueva prueba anterior sin replica.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pacientes"
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   10920
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3240
      TabIndex        =   14
      ToolTipText     =   "Cancelar"
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   3240
      TabIndex        =   13
      ToolTipText     =   "Aceptar"
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   315
      Left            =   3240
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   315
      Left            =   3240
      TabIndex        =   11
      ToolTipText     =   "Eliminar"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   9000
      TabIndex        =   10
      ToolTipText     =   "Buscar"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdmod 
      Caption         =   "&Modificar"
      Height          =   315
      Left            =   3240
      TabIndex        =   9
      ToolTipText     =   "Modificar"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Agregar"
      Height          =   315
      Left            =   3240
      TabIndex        =   8
      ToolTipText     =   "Agregar"
      Top             =   840
      Width           =   855
   End
   Begin VB.Frame Frame8 
      Caption         =   "Detalles:"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      Begin VB.TextBox Text1 
         DataField       =   "peso"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         MouseIcon       =   "Graficas.frx":409E
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   480
         Width           =   315
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Graficas.frx":41F0
         DataField       =   "Legajo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "nom"
         BoundColumn     =   "Legajo"
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "fechaPeso"
         DataSource      =   "Data2"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66977793
         CurrentDate     =   37858
      End
      Begin VB.Label Label3 
         Caption         =   "Peso (Kg):"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre y apellido:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   9960
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
End
Attribute VB_Name = "Graficas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb1 As Recordset

Private Sub cmdAdd_Click()
aux = 0
Data2.Recordset.AddNew
CmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdClose.Enabled = False
cmdmod.Enabled = False
'cmdbuscar.Enabled = False
cmdAceptar.Visible = True
cmdcancel.Visible = True
'DataCombo1.Enabled = True

'Command1.Enabled = False
DTPicker1.Enabled = True

Text2.Text = DataCombo1.BoundText

Text1.Enabled = True

Text1.SetFocus

cmdAceptar.Default = True
cmdcancel.Cancel = True
End Sub

Private Sub cmdBuscar_Click()
Dim strQuery As String

msg = InputBox("Ingrese apellido del paciente:", "Buscar por Apellido")

strQuery = " select * from pacientes where apell like '" & msg & "*' order by apell, nombre"

With Data1
    .RecordSource = strQuery
    .Refresh
End With

End Sub

Private Sub cmdDelete_Click()
  'esto puede producir un error si elimina el último
  'registro o el único registro del recordset

If Data2.Recordset.RecordCount > 0 And Data2.Recordset.EOF = False And Data2.Recordset.BOF = False Then
    msg = MsgBox("¿Desea Eliminar el registro actual?", vbYesNo, "Eliminar")
    
    If msg = vbYes Then
            Data2.Recordset.Delete
            Data2.Recordset.MovePrevious
    Else
        CmdAdd.SetFocus
    End If
End If
End Sub
Private Sub cmdClose_Click()
'Unload Me
Graficas.Hide

End Sub

Private Sub cmdAceptar_Click()
Data2.UpdateRecord
Data2.Recordset.Bookmark = Data2.Recordset.LastModified

    Text1.Enabled = False


'If aux = 0 Then
'    dbdiet.Execute "insert into alimenxpaciente (legajo, codalimento) select " & Val(Label1.Caption) & ", codalimento from alimentos where estado = true"
'End If

cmdAceptar.Visible = False
cmdcancel.Visible = False
'cmdbuscar.Enabled = True
CmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdClose.Enabled = True
cmdmod.Enabled = True
'DataCombo1.Enabled = False

DTPicker1.Enabled = False
CmdAdd.SetFocus
CmdAdd.Default = True
cmdClose.Cancel = True

MSChart2.Refresh
aux = 1
End Sub

Private Sub cmdcancel_Click()
Data2.Recordset.CancelUpdate


    Text1.Enabled = False



cmdAceptar.Visible = False
cmdcancel.Visible = False
'cmdbuscar.Enabled = True
CmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdClose.Enabled = True
cmdmod.Enabled = True
'DataCombo1.Enabled = False


DTPicker1.Enabled = False
CmdAdd.SetFocus
CmdAdd.Default = True
cmdClose.Cancel = True

aux = 1
End Sub

Private Sub cmdmod_Click()

If Data2.Recordset.BOF = True Or Data2.Recordset.EOF = True Then
    Data2.Recordset.MoveFirst
End If

CmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdClose.Enabled = False
cmdmod.Enabled = False
cmdAceptar.Visible = True
cmdcancel.Visible = True
'cmdbuscar.Enabled = False
'DataCombo1.Enabled = True


DTPicker1.Enabled = True


    Text1.Enabled = True

Data2.Recordset.Edit
Text1.SetFocus

cmdAceptar.Default = True
cmdcancel.Cancel = True

End Sub

Private Sub Command2_Click()

frmPacientes.Show
frmPacientes.SetFocus
frmPacientes.Data1.Recordset.FindFirst " legajo = " & DataCombo1.BoundText

End Sub

Private Sub Command3_Click()
'Dim chart1 As Chart
Call graphi


End Sub

Private Sub DataCombo1_Click(Area As Integer)
Adodc2.RecordSource = " select legajo, fechapeso as [Fecha de Consulta] , Peso  from pesos where pesos.legajo= " & DataCombo1.BoundText & " order by pesos.fechapeso"
Adodc2.Refresh

Datagrid1.ReBind
Datagrid1.Refresh


Call graficar

'DataGrid1.AddNewMode = a
'MsgBox DataGrid1.AddNewMode

End Sub

Private Sub DataGrid1_OnAddNew()
Datagrid1.Columns(0).Text = DataCombo1.BoundText
End Sub

Private Sub Form_Load()
Lugar = App.Path & "\db1nueva prueba anterior sin replica.mdb"
Set dbdiet = OpenDatabase(Lugar)

Text1.Enabled = False
DTPicker1.Enabled = False
'DataCombo1.Enabled = False

DataCombo1.BoundText = 1


Adodc2.RecordSource = " select legajo, fechapeso as [Fecha de Consulta] , Peso  from pesos where pesos.legajo= " & DataCombo1.BoundText & " order by pesos.fechapeso"
Adodc2.Refresh

Datagrid1.ReBind
Datagrid1.Refresh

Call graficar

Call graphi

aux = 1

End Sub

Sub graficar()
Dim hc, lip, prot

Set tb1 = dbdiet.OpenRecordset("pacientes", dbOpenDynaset)
tb1.FindFirst " legajo = " & DataCombo1.BoundText
hc = tb1.Fields("hckcal").Value
lip = tb1.Fields("lipkcal").Value
prot = tb1.Fields("protkcal").Value
tb1.Close


MSChart1.Column = 1
MSChart1.Data = hc

MSChart1.Column = 2
MSChart1.Data = lip

MSChart1.Column = 3
MSChart1.Data = prot

Adodc1.RecordSource = " select * from pesos where legajo = " & DataCombo1.BoundText
Adodc1.Refresh
If Adodc1.Recordset.RecordCount <> 0 Then

    Adodc1.Recordset.MoveFirst

    MSChart2.RowCount = Adodc1.Recordset.RecordCount

    For i = 1 To Adodc1.Recordset.RecordCount
        MSChart2.Row = i
        MSChart2.RowLabel = Adodc1.Recordset.Fields("fechapeso").Value
        MSChart2.Data = Adodc1.Recordset.Fields("peso").Value
        
        Adodc1.Recordset.MoveNext
    Next
    Adodc1.Recordset.Close
Else
    MSChart2.RowCount = 1
    MSChart2.Row = 1
    MSChart2.RowLabel = ""
    MSChart2.Data = 0
End If

End Sub
Sub graphi()
Dim Char As Chart
Set Char = CreateObject("msgraph.chart")

Char.Application.Visible = True
Char.Application.DataSheet.Range("00:G10").ClearContents
Char.ChartArea.ClearContents
Char.Application.DataSheet.Columns(3).Delete
Char.Application.DataSheet.Columns(4).Delete
Char.Application.DataSheet.Columns(5).Delete

Char.ChartType = -4102   '&& xl3DPie

Char.Application.DataSheet.Range("A0").Value = "Area"
Char.Application.DataSheet.Range("B0").Value = "Empresa"
Char.Application.DataSheet.Range("01").Value = "Legajos"

Char.Application.DataSheet.Range("A1").Value = 500
Char.Application.DataSheet.Range("B1").Value = 100
Char.HasTitle = True
Char.ChartTitle.Text = "Hola"
Char.Application.DataSheet.Activate

End Sub
