Attribute VB_Name = "mod_ini"
'DECLARATIONS API
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'==========================================================
'Referentes a los lblContact del frmAbout
'==========================================================
Declare Function GetDesktopWindow Lib "user32" () As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function sGetINI(sINIFIle As String, sSection As String, sKey As String, sDefault As String) As String
Dim sTemp As String * 256
Dim nLength As Integer

sTemp = Space(256)
nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFIle)
sGetINI = Left$(sTemp, nLength)

End Function

Public Sub writeINI(sINIFIle As String, sSection As String, sKey As String, sValue As String)
Dim n As Integer
Dim sTemp As String

sTemp = sValue

'reemplazar todos los caracteres CR/LF con espacios
For n = 1 To Len(sValue)
    If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf Then Mid$(sValue, n) = ""
Next n

n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFIle)

End Sub

