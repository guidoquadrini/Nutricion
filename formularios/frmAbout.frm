VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H8000000E&
   Caption         =   "Acerca de MiApl"
   ClientHeight    =   3600
   ClientLeft      =   2355
   ClientTop       =   1950
   ClientWidth     =   9150
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "frmAbout.frx":0ECA
   ScaleHeight     =   2484.784
   ScaleMode       =   0  'User
   ScaleWidth      =   8592.324
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   5760
      TabIndex        =   0
      Top             =   2520
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Info. del sistema..."
      Height          =   345
      Left            =   7560
      TabIndex        =   1
      Top             =   2520
      Width           =   1485
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "version 1.0.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A2A2FF&
      Height          =   480
      Left            =   0
      TabIndex        =   12
      Top             =   3216
      Width           =   2970
   End
   Begin VB.Label lblContact 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   6240
      MouseIcon       =   "frmAbout.frx":F3BF
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label lblContact 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   6240
      MouseIcon       =   "frmAbout.frx":F511
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label lbl_Cli_Email 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_Cli_Email"
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label lbl_Cli_Tel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_Cli_Tel"
      Height          =   195
      Left            =   3120
      TabIndex        =   8
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl_Cli_Direccion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_Cli_Direccion"
      Height          =   195
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label lbl_Cli_UserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lbl_Cli_UserName"
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este producto ha sido licenciado a:"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   2490
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción de la aplicación"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   11280
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Título de la aplicación"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   9480
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":F663
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   1560
      TabIndex        =   3
      Top             =   1665
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API Constants
Private Const SW_SHOWNORMAL As Long = 1

' Opciones de seguridad de clave del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos ROOT de clave del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Acerca de " & App.Title
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
    Me.Height = 4005
    Me.Width = 9270
    Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
    Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2

    Me.lbl_Cli_UserName.Caption = Cli_UserName
    Me.lbl_Cli_Direccion.Caption = Cli_Direccion
    Me.lbl_Cli_Tel.Caption = Cli_Tel
    Me.lbl_Cli_Email.Caption = Cli_Email

    lblContact(0).Caption = k_EMail
    lblContact(0).ToolTipText = "Enviar email"
    lblContact(1).Caption = k_URLwww
    lblContact(1).ToolTipText = "Ir al sitio web"
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Intentar obtener ruta de acceso y nombre del programa de Info. del sistema a partir del Registro...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Intentar obtener sólo ruta del programa de Info. del sistema a partir del Registro...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validar la existencia de versión conocida de 32 bits del archivo
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error: no se puede encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error: no se puede encontrar la entrada del Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador de una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de datos de una clave de Registro
    Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de variable de clave de Registro
    '------------------------------------------------------------
    ' Abrir clave de registro bajo KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error de controlador...
    
    tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
    KeyValSize = 1024                                       ' Marcar tamaño de variable
    
    '------------------------------------------------------------
    ' Obtener valor de clave de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtener o crear valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agregar cadena terminada en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Encontrado valor nulo, se va a quitar de la cadena
    Else                                                    ' En WinNT las cadenas no terminan en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se ha encontrado valor nulo, sólo se va a extraer la cadena
    End If
    '------------------------------------------------------------
    ' Determinar tipo de valor de clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Buscar tipos de datos...
    Case REG_SZ                                             ' Tipo de datos String de clave de Registro
        KeyVal = tmpVal                                     ' Copiar valor de cadena
    Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
        For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor carácter a carácter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a cadena
    End Select
    
    GetKeyValue = True                                      ' Se ha devuelto correctamente
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Borrar después de que se produzca un error...
    KeyVal = ""                                             ' Establecer valor a cadena vacía
    GetKeyValue = False                                     ' Fallo de retorno
    rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
End Function

Private Sub Label2_Click()
'frmBrowser.Show
'Unload frmAbout


End Sub

Private Sub lblDescription_Click()
'Dim name1 As String, name2 As String
'Demonstrate use of Write # statement
'Open "PIONEER.TXT" For Output As #1
'Write #1, "ENIAC"
'Write #1, 1946
'Write #1, "ENIAC", 1946; name1 = "Eckert"; name2 = "Mauchly"
'Write #1, 14 * 139, "J.P." & name1, name2, "John"
'Close #1

'GetObject "C:\Archivos de programa\Internet Explorer\IEXPLORE.EXE"
'Shell "C:\Archivos de programa\Internet Explorer\IEXPLORE.EXE, OpenUrl " & App.Path & "c:\gustavo\nutricion.htm", vbMaximizedFocus
'Shell "rundll32.exe, OpenUrl " & App.Path & "c:\gustavo\nutricion.htm", vbMaximizedFocus

End Sub


Private Sub lblContact_Click(Index As Integer)

    Dim sTopic As String
    Dim sFile As String
    Dim sParams As Variant
    Dim sDirectory As Variant

    If Index = 0 Then
        sFile = "mailto:" & lblContact(Index).Caption
    Else
        sFile = lblContact(Index).Caption
    End If
    
    sTopic = "Open"
    sParams = 0&
    sDirectory = 0&
    Call RunShellExecute(sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL)
    
End Sub

Private Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
   
End Sub
