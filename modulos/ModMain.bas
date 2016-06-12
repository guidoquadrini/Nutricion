Attribute VB_Name = "ModMain"
Sub main()
'=========================
'=========================
'Variable para manejo de archivo INI
Dim sINIFIle As String
Dim nCount As Integer
Dim i As Integer

sINIFIle = App.Path & "\ini\omnia.ini"

'leer el nombre del archivo ini. Si la clave que se busca no se encuentra devuelve el valor del parametro sDefault
''Cli_UserName = sGetINI(sINIFIle, "Cliente", "Cli_UserName", "?")
''Cli_Direccion = sGetINI(sINIFIle, "Cliente", "Cli_Direccion", "?")
''Cli_Tel = sGetINI(sINIFIle, "Cliente", "Cli_Tel", "?")
''Cli_Email = sGetINI(sINIFIle, "Cliente", "Cli_Email", "?")

DB_Name = sGetINI(sINIFIle, "Database", "DB_Name", "?")
DB_Path = sGetINI(sINIFIle, "Database", "DB_Path", "?")
UDL_Name = sGetINI(sINIFIle, "Database", "UDL_Name", "?")
UDL_Path = sGetINI(sINIFIle, "Database", "UDL_Path", "?")

App_Path = sGetINI(sINIFIle, "Aplicacion", "App_Path", "?")

'If sUserName = "?" Then
'    'No habia usuario...preguntar por él y guardarlo la pproxima vez
'    sUserName = InputBox$("Introduzca su nombre, por favor:")
'    writeINI sINIFIle, "Cliente", "UserName", " " & sUserName
'End If

'=========================
'=========================

'lugar = "c:\dietetica\db1nueva prueba anterior sin replica.mdb"
'Lugar = App.Path & "\database\db1nueva prueba anterior sin replica.mdb"

Lugar = DB_Path & "\" & DB_Name
m_ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_Path & "\" & DB_Name & ";Persist Security Info=False"

Set dbdiet = OpenDatabase(Lugar, False, False, ";pwd=208018")

'============================================================
'CONTROL DE ACCESO
'============================================================
frmLoginApp.Show vbModal

If frmLoginApp.LoginSucceeded Then

    frmSplash.Show
    frmSplash.Refresh
    frmSplash.ProgressBar1.Value = frmSplash.ProgressBar1.Value + 15

    Load MDIForm1
    frmSplash.ProgressBar1.Value = frmSplash.ProgressBar1.Value + 15
    Load frm_formulaDesarrollada
    frmSplash.ProgressBar1.Value = frmSplash.ProgressBar1.Value + 15
    'Load PrincipalFrm
    Load frm_FormulaSintetica
    frmSplash.ProgressBar1.Value = frmSplash.ProgressBar1.Value + 15
    Load frmPacientes
    frmSplash.ProgressBar1.Value = frmSplash.ProgressBar1.Value + 5
    Load frm_Adm_Diet
    
    
    MDIForm1.Show
    frmSplash.SetFocus
    Unload frmSplash
    
    frm_ModoAtencion.Show vbModal

Else
    
    End

End If

'============================================================




'==============================================
'Open "d:\version.txt" For Output As #1
'Dim sfile As String
'sfile = Dir$("D:\version.txt")
'If sfile = "" Then
'If Dir$("c:\gustavo\conexion.txt") = "" Then
'    MsgBox "Inserte el CD de DietCreates"
'    End
'Else
'    MsgBox "el CD de DietCreates es correcto"
'End If

'------------------
'crea origen de datos OLE DB 'Alimentos anterior sin replica.UDL' y asigna a la propriedad 'datasource' la ubicación de la Bd 'c:\dietetica2\db1nueva prueba anterior sin replica.mdb'
'Set Cn = New ADODB.Connection
'Cn.ConnectionString = "FILE NAME=" & App.Path & "\Alimentos anterior sin replica.UDL" '"dsn=biblio"
'Cn.Properties("data source") = App.Path & "\db1nueva prueba anterior sin replica.mdb"
'Cn.Open
'Cn.Open "FILE NAME=" & App.Path & "\Alimentos anterior sin replica.UDL;data source = " & App.Path & "\db1nueva prueba anterior sin replica.mdb"
'Cn.Open "FILE NAME=c:\dietetica2\Alimentos anterior sin replica.UDL;data source = c:\dietetica2\db1nueva prueba anterior sin replica.mdb"
'MsgBox Cn.Properties("data source")
'Cn.Properties("data source") = App.Path & "\db1nueva prueba anterior sin replica.mdb"
'MsgBox Cn.Properties("data source")
'Cn.Close
'Set Cn = Nothing
'-------------------
'==============================================

End Sub
