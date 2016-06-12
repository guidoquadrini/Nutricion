VERSION 5.00
Begin VB.Form frmLoginApp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3390
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5580
   Icon            =   "frmLoginApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoginApp.frx":08CA
   ScaleHeight     =   2002.923
   ScaleMode       =   0  'User
   ScaleWidth      =   5239.317
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000005&
      Height          =   345
      Left            =   2130
      TabIndex        =   1
      Text            =   "yanina"
      Top             =   1695
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   1335
      TabIndex        =   4
      Top             =   2580
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2940
      TabIndex        =   5
      Top             =   2580
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000005&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2130
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2085
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre de usuario:"
      Height          =   195
      Index           =   0
      Left            =   735
      TabIndex        =   0
      Top             =   1710
      Width           =   1380
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Contraseña:"
      Height          =   195
      Index           =   1
      Left            =   735
      TabIndex        =   2
      Top             =   2100
      Width           =   855
   End
End
Attribute VB_Name = "frmLoginApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Dim tb_Usuarios As ADODB.Recordset


Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    Me.Hide
    
    End
       
End Sub

Private Sub cmdOK_Click()
    'comprobar si la contraseña es correcta
    Dim strQuery As String
    Dim i As Integer
    
    'Get tbl Usuarios ======================================
    strQuery = "Usuarios"
    Set tb_Usuarios = f_StaticRecordset(adCmdTable, strQuery)
    tb_Usuarios.MoveFirst
    '=======================================================
    'para cada registro
    For i = 1 To tb_Usuarios.RecordCount
        
        Dim sUserName, sPassword As String
        sUserName = ""
        sPassword = ""
        
        If Not IsNull(tb_Usuarios.Fields("usr_userName").Value) Then
            sUserName = tb_Usuarios.Fields("usr_userName").Value
        End If
        If Not IsNull(tb_Usuarios.Fields("usr_Password").Value) Then
            sPassword = tb_Usuarios.Fields("usr_Password").Value
        End If
        
        'si el usuario es correcto
        If txtPassword = sPassword And txtUserName = sUserName Then
            'colocar código aquí para pasar al sub
            'que llama si la contraseña es correcta
            'lo más fácil es establecer una variable global
            LoginSucceeded = True
            
            'mantengo el usuario logeado "whois"
            nUserLoging = tb_Usuarios.Fields("usr_codprf").Value
            sUserNivel = tb_Usuarios.Fields("usr_Nivel").Value
            
            Set tb_Usuarios = Nothing
            
            Unload Me
            
            Exit For
            
        End If
        
        tb_Usuarios.MoveNext
        
    Next
    
    If LoginSucceeded = False Then
    
        MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    
    End If
        
    Set tb_Usuarios = Nothing
End Sub


