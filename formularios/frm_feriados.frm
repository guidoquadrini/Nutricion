VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_feriados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   2910
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin MSComCtl2.MonthView MonthView1 
         BeginProperty DataFormat 
            Type            =   4
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   8
         EndProperty
         Height          =   2370
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "DblClick para Agregar o Eliminar un feriado"
         Top             =   240
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483639
         Appearance      =   1
         MonthBackColor  =   12632256
         StartOfWeek     =   68026375
         TitleBackColor  =   8421504
         TitleForeColor  =   -2147483639
         CurrentDate     =   38231
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2760
         Width           =   2505
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "FILE NAME=Alimentos anterior sin replica.UDL"
      OLEDBString     =   ""
      OLEDBFile       =   "Alimentos anterior sin replica.UDL"
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from feriados"
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
Attribute VB_Name = "frm_feriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
Dim Estilo As Integer
Dim Legajo As Long

Private Sub Form_Load()

Me.Caption = Titulo 'llama a la function Titulo

Call Size("año")

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

If Estilo = 0 Then
    Call CompruebaFecha(DateClicked)
Else
    If CompruebaFecha(DateClicked) Then
        Label1.Caption = "Doble Click para recuperar el item"
    Else
        Label1.Caption = ""
    End If
End If

End Sub

Private Sub MonthView1_DateDblClick(ByVal DateDblClicked As Date)

If Estilo = 0 Then 'solo si es feriado
    If CompruebaFecha(DateDblClicked) = True Then
        
        msg = MsgBox("¿Desea Eliminar el Feriado Seleccionado?", vbYesNo, "Eliminar")
        
        If msg = vbYes Then
                
             Adodc1.Recordset.Delete
            
            MonthView1.DayBold(DateDblClicked) = False
         
        End If
    
    Else
    
        msg = MsgBox("¿Desea Asignar Nuevo Feriado en la Fecha Seleccionada?", vbYesNo, "Agregar")
        
        If msg = vbYes Then
            
            msg = InputBox("Ingrese Descripcion del Feriado:", "Feriado")
   
            If msg <> "" Then
                               
                dbdiet.Execute "insert into feriados (fdo_fecha, descripcion) select " & "'" & Format(DateDblClicked, "dd/mm/yy") & "'" & ", '" & msg & "'"
                                
                MonthView1.DayBold(DateDblClicked) = True
                                  
            End If
        
        End If
                    
    End If

Else ' si es plan alimentario
    
    MousePointer = vbHourglass
    
    Form1.DataCombo1.BoundText = Legajo
    Form1.DTPicker1.Value = DateDblClicked
    Form1.cmd_aceptar.Value = True
    Unload Me
    Form1.Show
    
    MousePointer = vbDefault

End If
End Sub

Private Sub MonthView1_GetDayBold(ByVal StartDate As Date, ByVal Count As Integer, State() As Boolean)
'si es por mes o por año
If MonthView1.MonthColumns = 1 And MonthView1.MonthRows = 1 Then
    
    Call Fdo_Mes(StartDate, Count, State())

Else

    Call Fdo_Año(StartDate, Count, State())

End If

End Sub

Private Sub Fdo_Mes(ByVal StartDate As Date, ByVal Count As Integer, State() As Boolean)

Dim strQuery As String
Dim nferiado As Integer 'índice que indica el feriado dentro de la "matriz" mes actual

If Estilo = 0 Then 'si se trata de feriados
    Frame1.Caption = "Feriados del mes"
       
    'selecciona todos los feriados del mes actual
    strQuery = " select * , month(fdo_fecha) as Mes from feriados where month(fdo_fecha) = " & MonthView1.Month
    
    Set Label1.DataSource = Adodc1
    Label1.DataField = "descripcion"
    
Else 'de lo contrario si se trata de planes alimentarios
    Frame1.Caption = "Planes alimentarios del mes"
       
    'dado un determinado legajo selecciona todos los dias que tiene realizados planes alimentarios del mes actual
    strQuery = "select distinct fechamenu, month(fechamenu) as mes from menu where legajo = " & Legajo & " and month(fechamenu) = " & MonthView1.Month
        
End If

nferiado = 0

With Adodc1
    .RecordSource = strQuery
    .Refresh
End With

'si al menos hay 1 feriado/plan alimentario continuo
If Adodc1.Recordset.RecordCount <> 0 Then
    
    Adodc1.Recordset.MoveFirst
    
    'por cada feriado del recorset
    For i = 1 To Adodc1.Recordset.RecordCount
        'obtengo la diferencia en dias entre el feriado y la primer fecha que muestra
        'el MonthView para utilizarlo como índice en el array "state()"
        If Estilo = 0 Then 'si es feriado
            nferiado = Abs(DateDiff("d", Adodc1.Recordset.Fields("fdo_fecha").Value, StartDate))
        Else 'si es plan alimentario
            nferiado = Abs(DateDiff("d", Adodc1.Recordset.Fields("fechaMenu").Value, StartDate))
        End If
        
        If nferiado <= 42 Then 'si la fecha se visualiza
        
            State(nferiado) = True 'es un array booleano que establece que fecha se marcara con negrita
        
        End If
        
        If Adodc1.Recordset.EOF = False Or Adodc1.Recordset.BOF = False Then
            Adodc1.Recordset.MoveNext
        End If
    
    Next

End If

End Sub

Private Sub Fdo_Año(ByVal StartDate As Date, ByVal Count As Integer, State() As Boolean)

Dim strQuery As String
Dim nferiado As Integer 'índice que indica el feriado dentro de la "matriz" mes actual

If Estilo = 0 Then
    Frame1.Caption = "Feriados del año"

    'selecciona todos los feriados del año actual
    strQuery = " select *, year(fdo_fecha) as año from feriados where year(fdo_fecha) = " & MonthView1.Year

    Set Label1.DataSource = Adodc1
    Label1.DataField = "descripcion"
    
Else 'de lo contrario si se trata de planes alimentarios
    Frame1.Caption = "Planes alimentarios del año"
       
    'dado un determinado legajo selecciona todos los dias que tiene realizados planes alimentarios del año actual
    strQuery = "select distinct fechamenu, year(fechamenu) as año from menu where legajo = " & Legajo & " and year(fechamenu) = " & MonthView1.Year

End If

nferiado = 0

With Adodc1
    .RecordSource = strQuery
    .Refresh
End With

'si al menos hay 1 feriado continuo
If Adodc1.Recordset.RecordCount <> 0 Then
    
    Adodc1.Recordset.MoveFirst
    
    'por cada feriado del recorset
    For i = 1 To Adodc1.Recordset.RecordCount
        'obtengo la diferencia en dias entre el feriado y la primer fecha que muestra
        'el MonthView para utilizarlo como índice en el array "state()"
        If Estilo = 0 Then ' si es feriado
            nferiado = Abs(DateDiff("d", Adodc1.Recordset.Fields("fdo_fecha").Value, StartDate))
        Else ' si es plan alimentario
            nferiado = Abs(DateDiff("d", Adodc1.Recordset.Fields("fechaMenu").Value, StartDate))
        End If
        
        If nferiado <= 378 Then 'si la fecha se visualiza
        
            State(nferiado) = True 'es un array booleano que establece que fecha se marcara con negrita
        
        End If
        
        If Adodc1.Recordset.EOF = False Or Adodc1.Recordset.BOF = False Then
            Adodc1.Recordset.MoveNext
        End If
    
    Next
    
End If

End Sub


Function CompruebaFecha(Fecha As Date) As Boolean
'dada una fecha comprueba si ya esta cargada en la tabla "feriados" devolviendo true, de lo contrario false

Dim strQuery

If Estilo = 0 Then ' si es feriado
    strQuery = " select * from feriados where fdo_fecha = " & "#" & Format(Fecha, "Mm/dd/yy") & "#"
Else ' si es plan alimentario
    strQuery = " select distinct legajo, fechamenu from menu where fechaMenu = " & "#" & Format(Fecha, "Mm/dd/yy") & "#"
End If

With Adodc1
    .RecordSource = strQuery
    .Refresh
End With

If Adodc1.Recordset.RecordCount = 0 Then
    
    CompruebaFecha = False
    
Else
    
    CompruebaFecha = True
    
End If

End Function

Public Sub Size(Mes_Año As String)

'define valores de apariencia segun se visualise todo el año o solo el mes
If Mes_Año = "año" Then

    MonthView1.MonthColumns = 4
    MonthView1.MonthRows = 3
    Label1.Top = 6960
    Label1.Left = 7920
    Me.Height = 7740
    Me.Width = 10800
    Frame1.Height = 7335
    Frame1.Width = 10695
    
    'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
    Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
    Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

    
Else

    MonthView1.MonthColumns = 1
    MonthView1.MonthRows = 1
    Label1.Top = 2760
    Label1.Left = 120
    Me.Height = 3555
    Me.Width = 3030
    Frame1.Height = 3135
    Frame1.Width = 2895
    
    'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
    Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
    Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

    
End If

End Sub

Private Sub MonthView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'desplega el popup menu para poder seleccionar la visualizacion del control monthview
If Button = 2 Then
    
    PopupMenu MDIForm1.popupFeriado, vbPopupMenuLeftAlign
    
End If

End Sub

Public Sub cargarParametros(estiloParam As Integer, Optional LegajoParam As Long)
'estilo representa si se va a tratar de feriados o de planes alimentarios
'estilo = 0 es feriados
'estilo = 1 es planes alimentarios
Estilo = estiloParam
Legajo = LegajoParam

End Sub

Function Titulo() As String
Dim strQuery As String

If Estilo = 1 Then

    'strQuery = "select *, (apell & ',' & nombre) as nom from pacientes where legajo = " & Legajo
    strQuery = "select *, (apell & ', ' & nombre) as nom from pacientes where legajo = " & Legajo
    
    Set tb = dbdiet.OpenRecordset(strQuery)
    Titulo = "Calendario - " & tb.Fields("nom").Value
    tb.Close
    
Else
    
    Titulo = "Calendario"
    
End If

End Function
