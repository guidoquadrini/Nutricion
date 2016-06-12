Attribute VB_Name = "VarPublic"
Public a, b, c, d, fa, fi

Public dbdiet As Database
Public Lugar As String
Public m_ConnectionString As String

Public CodigoPlato() As String
'Public CodigoAlimento() As String
Public CantFinal
Public CantAux
Public PrevCantAux
Public Suma
Public UnidadPlato
Public CantidadNeta
'Public Kcal
Public Porcion
Public TotalPorcion
Public PlatoAgregado As String
Public unidadAgregado As Long
Public IngredAgregado As Long
Public PorcionAgregado As Integer
Public KeyPlato As String
Public Categoria1

'Public estadoAbm As Integer ' define el estado de un formulario de abm
                            ' 1 = sin cambios; 2 = agregar; 3 = modificar

Dim ArmaArchivo() As String
Dim ArmaArchivo1() As String

'Public Cn As ADODB.Connection 'variable public para origen de datos OLE DB

'Public nLegajo As Integer

'=====================================================================
'Constantes de Cliente
'=====================================================================
Public Const Cli_UserName As String = "Yanina Pulpeiro"
Public Const Cli_Direccion As String = "Italia 5553"
Public Const Cli_Tel As String = "(0341) - 4610798"
Public Const Cli_Email As String = "yaninapulpeiro@hotmail.com"
'=====================================================================
'Constantes de Owner
'=====================================================================
Public Const k_EMail As String = "info@grupoumana.com.ar"
Public Const k_URLwww As String = "http://www.grupoumana.com.ar"
'=====================================================================
'=====================================================================

'======================================
'variables publicas del archivo INI
'[Aplicacion]
Public App_Name As String
Public App_Path As String
'[Owner]
Public Owner_UserName As String
Public Owner_Direcion As String
Public Owner_Tel As String
Public Owner_Email As String
'[Database]
Public DB_Name As String
Public DB_Path As String
Public UDL_Name As String
Public UDL_Path As String
'''[Cliente]
''Public Cli_UserName As String
''Public Cli_Direccion As String
''Public Cli_Tel As String
''Public Cli_Email As String
'=======================================

'=======================================
'Variables Modo Atencion/Cosulta Paciente
'=======================================
Public nCodAtencion As Long 'es = 0 para el caso de que no se acceda en Modo Atencion/Consulta
'=======================================

'=======================================
'Variables Usuario Logeado
'=======================================
Public nUserLoging As Long
Public sUserNivel As String
'=======================================
