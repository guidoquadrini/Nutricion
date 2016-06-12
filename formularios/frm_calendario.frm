VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_calendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   Icon            =   "frm_calendario.frx":0000
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
         MonthBackColor  =   16777215
         StartOfWeek     =   24576007
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
Attribute VB_Name = "frm_calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tb As Recordset
'=====parametros de entrada del form========
Dim Estilo As Integer
Dim Legajo As Long
Dim cToolTipText As String
'===========================================

Dim estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar
                            
Private Sub Form_Load()

Call f_CargarOrigenDatos

Me.Caption = Titulo 'llama a la function Titulo

Me.MonthView1.ToolTipText = cToolTipText

Me.MonthView1.Value = Now

Call Size("año")

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

Select Case Estilo
    
    Case Is = 0 ' feriados
    
        Call CompruebaFecha(DateClicked)
    
    Case Is = 1 ' plan alimentario

        If CompruebaFecha(DateClicked) Then
            Label1.Caption = "Doble Click para recuperar el item"
        Else
            Label1.Caption = ""
        End If

    Case Is = 2 'turnos
    
        If CompruebaFecha(DateClicked) Then
            Label1.Caption = "Doble Click para recuperar el item"
        Else
            Label1.Caption = ""
        End If
        
End Select

End Sub

Private Sub MonthView1_DateDblClick(ByVal DateDblClicked As Date)

Select Case Estilo
    
    Case Is = 0 'solo si es feriado
         
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
                                                                           
                    'agregamos un nuevo feriado
                    Adodc1.Recordset.AddNew
                    Adodc1.Recordset.Fields("fdo_Dia").Value = Day(DateDblClicked)
                    Adodc1.Recordset.Fields("fdo_Mes").Value = Month(DateDblClicked)
                    Adodc1.Recordset.Fields("descripcion").Value = msg
                    Adodc1.Recordset.Update
                                        
                    MonthView1.DayBold(DateDblClicked) = True
                                      
                End If
            
            End If
                         
         End If

    Case Is = 1 ' si es plan alimentario
        
        MousePointer = vbHourglass
        
        frm_Adm_Diet.DataCombo1.BoundText = Legajo
        frm_Adm_Diet.DTPicker1.Value = DateDblClicked
        frm_Adm_Diet.cmd_aceptar.Value = True
        Unload Me
        frm_Adm_Diet.Show
        
        MousePointer = vbDefault
    
    Case Is = 2 ' si es turnos
        
        MousePointer = vbHourglass
        
        frm_agendaTurnos.DataCombo1.BoundText = Legajo
        frm_agendaTurnos.DTPicker1.Value = DateDblClicked
        
        Unload Me
        frm_agendaTurnos.Show
        
        frm_agendaTurnos.cmd_aceptar.Value = True
        
        MousePointer = vbDefault
        
End Select

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

Select Case Estilo

    Case Is = 0 'si se trata de feriados
        
        Frame1.Caption = "Feriados del mes"
       
        'selecciona todos los feriados del mes actual
        strQuery = " select * from feriados where fdo_Mes = " & MonthView1.Month
        
        Set Label1.DataSource = Adodc1
        Label1.DataField = "descripcion"
    
    Case Is = 1 'planes alimentarios
    
        Frame1.Caption = "Planes alimentarios del mes"
       
        'dado un determinado legajo selecciona todos los dias que tiene realizados planes alimentarios del mes actual
        strQuery = "select distinct fechamenu, month(fechamenu) as mes from menu where legajo = " & Legajo & " and month(fechamenu) = " & MonthView1.Month
            
    Case Is = 2
    
        Frame1.Caption = "Turnos asignados durante el mes"
       
        'dado un determinado legajo del prof selecciona todos los dias que tiene asignado turno del mes actual
        strQuery = "select distinct tur_fecha, month(tur_fecha) as mes from turnos where tur_idprof = " & Legajo & " and month(tur_fecha) = " & MonthView1.Month
    
End Select

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
        Select Case Estilo
        
            Case Is = 0 'si es feriado
                Dim fNewFecha As Date
                
                fNewFecha = Adodc1.Recordset.Fields("fdo_Dia").Value & "/" & Adodc1.Recordset.Fields("fdo_Mes").Value & "/" & Me.MonthView1.Year
                
                nferiado = Abs(DateDiff("d", fNewFecha, StartDate))
            
            Case Is = 1 'si es plan alimentario
                
                nferiado = Abs(DateDiff("d", Adodc1.Recordset.Fields("fechaMenu").Value, StartDate))
                
            Case Is = 2
            
                nferiado = Abs(DateDiff("d", Adodc1.Recordset.Fields("tur_fecha").Value, StartDate))
            
        End Select
        
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

Select Case Estilo
    
    Case Is = 0 ' feriados
        
        Frame1.Caption = "Feriados del año"
    
        'selecciona todos los feriados del año
        strQuery = "select * from feriados"
    
        Set Label1.DataSource = Adodc1
        Label1.DataField = "descripcion"
        
    Case Is = 1 'plan alimentario
    
        Frame1.Caption = "Planes alimentarios del año"
           
        'dado un determinado legajo selecciona todos los dias que tiene realizados planes alimentarios del año actual
        strQuery = "select distinct fechamenu, year(fechamenu) as año from menu where legajo = " & Legajo & " and year(fechamenu) = " & MonthView1.Year

    Case Is = 2 ' turnos
    
        Frame1.Caption = "Turnos asignados durante el año"
           
        'dado un determinado legajo selecciona todos los dias que tiene turnos signados en el año actual
        strQuery = "select distinct tur_fecha, year(tur_fecha) as año from turnos where tur_idprof = " & Legajo & " and year(tur_fecha) = " & MonthView1.Year
            
End Select

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
        Select Case Estilo
            
            Case Is = 0 ' si es feriado
            
                Dim fNewFecha As Date
                
                fNewFecha = Adodc1.Recordset.Fields("fdo_Dia").Value & "/" & Adodc1.Recordset.Fields("fdo_Mes").Value & "/" & Me.MonthView1.Year
            
                nferiado = Abs(DateDiff("d", fNewFecha, StartDate))
            
            Case Is = 1 ' si es plan alimentario
            
                nferiado = Abs(DateDiff("d", Adodc1.Recordset.Fields("fechaMenu").Value, StartDate))
            
            Case Is = 2 ' si es turno
            
                nferiado = Abs(DateDiff("d", Adodc1.Recordset.Fields("tur_fecha").Value, StartDate))
                
        End Select
        
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

Select Case Estilo
    
    Case Is = 0 ' si es feriado
        
        strQuery = " select * from feriados where fdo_Dia = " & Day(Fecha) & " and fdo_Mes = " & Month(Fecha)
    
    Case Is = 1 ' si es plan alimentario
        
        strQuery = " select distinct legajo, fechamenu from menu where legajo = " & Legajo & " and fechaMenu = " & "#" & Format(Fecha, "Mm/dd/yy") & "#"
    
    Case Is = 2 ' si es turnos

        strQuery = " select distinct * from turnos where tur_idprof = " & Legajo & " and tur_fecha = " & "#" & Format(Fecha, "Mm/dd/yy") & "#"
        
End Select

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

Public Sub cargarParametros(estiloParam As Integer, Optional LegajoParam As Long, Optional ToolTipTextParam As String)
'estilo representa si se va a tratar de feriados o de planes alimentarios
'estilo = 0 es feriados
'estilo = 1 es planes alimentarios
'estilo = 2 es Turnos
Estilo = estiloParam
Legajo = LegajoParam
cToolTipText = ToolTipTextParam

End Sub

Function Titulo() As String
Dim strQuery As String

Select Case Estilo
    Case Is = 0
   
        Titulo = "Calendario"
        
    Case Is = 1
        
        'strQuery = "select *, (apell & ',' & nombre) as nom from pacientes where legajo = " & Legajo
        strQuery = "select *, (apell & ', ' & nombre) as nom from pacientes where legajo = " & Legajo
        
        Set tb = dbdiet.OpenRecordset(strQuery)
        Titulo = "Calendario - " & tb.Fields("nom").Value
        tb.Close
    
    Case Is = 2
        
        Titulo = "Calendario de turnos "
                
End Select

End Function

Sub f_CargarOrigenDatos()
Dim strQuery As String

strQuery = ""

Set Me.Adodc1.Recordset = Nothing

Select Case Estilo
    
    Case Is = 0 ' si es feriado
        
        strQuery = " select * from feriados "
    
    Case Is = 1 ' si es plan alimentario
        
        strQuery = " select distinct legajo, fechamenu from menu where legajo = " & Legajo
    
    Case Is = 2 ' si es turnos

        strQuery = " select distinct * from turnos where tur_idprof = " & Legajo
        
End Select

Call f_Adodc_ConnectionString(Adodc1, strQuery)

End Sub
