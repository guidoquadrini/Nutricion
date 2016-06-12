Attribute VB_Name = "FuncGrls"


Sub DatagridRefresh(cAdodc As Adodc, cDatagrid As DataGrid)

Dim Result As String

'Ejecuta un método de un objeto o establece o devuelve una propiedad de un objeto.
Result = CallByName(cAdodc, "Refresh", VbMethod)
Result = CallByName(cDatagrid, "ReBind", VbMethod)
Result = CallByName(cDatagrid, "Refresh", VbMethod)

End Sub

Function f_isFeriado(Fecha As Date) As Boolean
'dada una fecha comprueba si ya esta cargada en la tabla "feriados" devolviendo true, de lo contrario false
Dim strQuery
Dim tb As Recordset

strQuery = " select * from feriados where fdo_Dia = " & Day(Fecha) & " and fdo_Mes = " & Month(Fecha)

Set tb = dbdiet.OpenRecordset(strQuery)

If tb.RecordCount = 0 Then

    f_isFeriado = False
    
Else
    
    f_isFeriado = True
    
End If

tb.Close

End Function

Function f_Tiene_ExcepcionHrs(hrs_idProf As Long, hrs_Fecha As Date) As Boolean
'devuelve si un profesional en una fecha dad tiene o no una excepcion horaria para los turnos
Dim tb As Recordset
Dim strQuery As String

strQuery = " select * from excepcionesHrs where ehr_idProf = " & hrs_idProf & " and ehr_fecha = #" & hrs_Fecha & "#"

Set tb = dbdiet.OpenRecordset(strQuery)

If tb.RecordCount = 0 Then

    f_Tiene_ExcepcionHrs = False

Else
    
    f_Tiene_ExcepcionHrs = True
    
End If

tb.Close

End Function

Function f_StaticRecordset(ByVal oCommandType As CommandTypeEnum, ByVal sCommandText As String, Optional ByVal param As Variant) As ADODB.Recordset   'param() As Integer) As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim pmt As ADODB.Parameter
Dim i As Long

i = 0

Set cmd = New ADODB.Command
Set rs = New ADODB.Recordset

With cmd
    .ActiveConnection = m_ConnectionString
    .CommandType = oCommandType
    .CommandText = sCommandText '"[NameTbl/NameConsulta]"
    
    On Error Resume Next
    For Each pmt In .Parameters
        pmt.Value = param(i)
        i = i + 1
    Next
    
End With

With rs
    .CursorLocation = adUseClient
    .Open cmd, , adOpenStatic
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set .ActiveConnection = Nothing
End With

Set f_StaticRecordset = rs

End Function

Sub f_UpdateInsertRecordset(ByVal oCommandType As CommandTypeEnum, ByVal sCommandText As String)
Dim cmd As ADODB.Command

Set cmd = New ADODB.Command

With cmd
    .ActiveConnection = m_ConnectionString
    .CommandType = oCommandType
    .CommandText = sCommandText '"[NameTbl/NameConsulta]"
    .Execute
End With

Set cmd.ActiveConnection = Nothing
Set cmd = Nothing

End Sub

Sub f_Adodc_ConnectionString(param_Adodc As Adodc, param_RecordSource As String)
'Ejecuta un método de un objeto o establece o devuelve una propiedad de un objeto.
'En este caso defino la propiedad ConnectionString y RecordSource del objeto adodc

CallByName param_Adodc, "ConnectionString", VbLet, m_ConnectionString '("FILE NAME=" & UDL_Path & "\" & UDL_Name)
CallByName param_Adodc, "RecordSource", VbLet, param_RecordSource
param_Adodc.Refresh

End Sub

Sub f_Data_DatabaseName(param_Data As Data, param_RecordSource As String)
'define la propiedad databaseName y RecordSource del objeto data

CallByName param_Data, "DatabaseName", VbLet, (DB_Path & "\" & DB_Name)
CallByName param_Data, "RecordSource", VbLet, param_RecordSource
param_Data.Refresh

End Sub

Sub f_Enlaza_ControlData(oData As Object, oDataSource As Object, oRowSource As Object, sDataField As String, sBoundColumn As String, sListField As String)
'inicializa las propiedades de los objetos DataList y DataCombo luego de definir el origen de datos con la funcion f_CargarOrigenDatos propia de cada form
'elimina el problema de tener que definir la propiedad Bountext del objeto data de antemano en el evento load
Dim Result As String

'ya que si el objeto es un data es imposible asignar esa propiedad en tiempo de ejecucion
'por lo tanto se la excluye
If VarType(oDataSource) = vbObject Then
    CallByName oData, "DataSource", VbSet, oDataSource
End If

CallByName oData, "RowSource", VbSet, oRowSource

CallByName oData, "DataField", VbLet, sDataField

CallByName oData, "BoundColumn", VbLet, sBoundColumn

CallByName oData, "ListField", VbLet, sListField

Result = CallByName(oData, "Refresh", VbMethod)

End Sub

Function f_Cant_Registros(tb As Recordset) As Integer
'devuelve al cantidad de registros dentro de un recorset, ya se que es una guazada hacer esto
'pero gracias a microsoft la funcion tb.recordcount no funciona bien por lo que tengo que hacerlo a mano
f_Cant_Registros = 0

If tb.RecordCount > 0 Then

    tb.MoveFirst
    
    While tb.EOF = False
    
        f_Cant_Registros = f_Cant_Registros + 1
        
        tb.MoveNext
    
    Wend

End If

End Function

Sub f_print(CrystalReport1 As CrystalReport, ByVal SelectionFormula As String, Destination As DestinationConstants)
Dim Result As String
'establece el enlace a los datos
CrystalReport1.DataFiles(0) = DB_Path & "\" & DB_Name

'establece el valor de los parametros
CrystalReport1.ParameterFields(0) = "Cli_UserName;" & Cli_UserName & ";true"
CrystalReport1.ParameterFields(1) = "Cli_Direccion;" & Cli_Direccion & ";True"
CrystalReport1.ParameterFields(2) = "Cli_Tel;" & Cli_Tel & ";True"
CrystalReport1.ParameterFields(3) = "Cli_Email;" & Cli_Email & ";True"

'aclare el filtro para imprimir
CrystalReport1.SelectionFormula = SelectionFormula

'CrystalReport1.Destination = crptToWindow
CrystalReport1.Destination = Destination

CrystalReport1.WindowState = crptMaximized
CrystalReport1.PrintReport

End Sub

'Devuelve el Apellido y Nombre del Usuario Logeado
Function f_GetUserLoging(ByVal nUserLoging As Long) As String
Dim tb_NombreProfesional As ADODB.Recordset
Dim strQuery As String

'valida que el usuario logeado corresponda a un profesional
If nUserLoging > 0 Then
    
    'selecciona el profesional con el codigo = nUserLoging
    strQuery = "SELECT (prf_apell & ', ' & prf_nombre) AS NombreProfesional FROM profesionales WHERE prf_codigo = " & nUserLoging
    Set tb_NombreProfesional = f_StaticRecordset(adCmdText, strQuery)
    
    'devuelve el valor obtenido
    f_GetUserLoging = tb_NombreProfesional.Fields("NombreProfesional").Value
    
    Set tb_NombreProfesional = Nothing

Else

    Exit Function

End If

End Function

Sub f_GetPermisos()

If nCodAtencion > 0 Then 'Esta en Modo Atencion/Consulta

    'Habilitar menu modo de ACCESO=========
    MDIForm1.mnu_ModoAcceso.Enabled = False
    MDIForm1.mnu_ConsultaNueva.Enabled = False
    MDIForm1.mnu_ConsultaExistente.Enabled = False
    '======================================

    'definiendo permisos================================================
    MDIForm1.tbToolBar.Buttons("FormulaSintetica").Enabled = True
    MDIForm1.tbToolBar.Buttons("FormulaDesarrollada").Enabled = True
    MDIForm1.tbToolBar.Buttons("GestionPlanAlimentario").Enabled = True
    MDIForm1.tbToolBar.Buttons("FinalizarCosulta").Enabled = True
    
    MDIForm1.mnu_FormulaSintetica.Enabled = True
    MDIForm1.mnu_FormulaDesarollada.Enabled = True
    MDIForm1.mnu_GestionPlanAlimentario.Enabled = True
    MDIForm1.mnu_FinalizarCosulta.Enabled = True
    '===================================================================

Else

    'Habilitar menu modo de ACCESO=========
    MDIForm1.mnu_ModoAcceso.Enabled = True
    MDIForm1.mnu_ConsultaNueva.Enabled = True
    MDIForm1.mnu_ConsultaExistente.Enabled = True
    '======================================
    
    'definiendo permisos================================================
    MDIForm1.tbToolBar.Buttons("FormulaSintetica").Enabled = False
    MDIForm1.tbToolBar.Buttons("FormulaDesarrollada").Enabled = False
    MDIForm1.tbToolBar.Buttons("GestionPlanAlimentario").Enabled = False
    MDIForm1.tbToolBar.Buttons("FinalizarCosulta").Enabled = False
    
    MDIForm1.mnu_FormulaSintetica.Enabled = False
    MDIForm1.mnu_FormulaDesarollada.Enabled = False
    MDIForm1.mnu_GestionPlanAlimentario.Enabled = False
    MDIForm1.mnu_FinalizarCosulta.Enabled = False
    '===================================================================
    
End If

End Sub
