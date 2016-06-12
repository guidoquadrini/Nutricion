VERSION 5.00
Begin VB.Form frm_Creditos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Especialmente gracias para..."
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   Icon            =   "frm_Creditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_Creditos.frx":0ECA
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   284
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frm_Creditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020    ' (DWORD) dest = source

'Const kCredits = "Alan Harris-Reid|Alessandro Frattini|Alex|Allanon|Andrea Caprioli|Andrew Stead|Arun Kumar|Cooley|Enzo Carfora|Ettore Maronese, il -Mes-|Francesco Bancalà|Fibia FBI|Fred Just|Gert E.R. Drapers|Gianluca Hotz|Gianni Carrozzo|Greg Hines|Humberto Rodrigues|Jan Nielsen|Jason Zheng|Jordan Russel|Karl E. Peterson|Ken Williams|KPDTeam@Allapi.net|Ivano Modenin|Lars Broberg|Lorenzo Benaglia|Luc Wuyts|Luca Bianchi|Lucía Valles|Luigi De Gregori|Marcello Biglioli|Mark Nadig|Mario Lopes|Morten Skille|Narayana Vyas Kondreddi|P. R.|Paco Bueno|Paolo Castagnetti|Paolo Fisco|Paul Tissue|Peter Schmid|Peter Storz|Peter Swaniker|Planet Source Code|Rahul Sharma|Ryan Stone|Robert Vallee|Roberto Gismondi|Ryan Moore|Vincenzo Morgante|William (Bill) Vaughn|Yusuf Incekara"
Const kCredits = "Licenciada en Nutricion Yanina Carla Pulpeiro|quien ha aportado|su gran conocimiento profesional|haciendo posible la creacion|de una herramienta practica|para su desempeño,|a ella la mas sincera GRATITUD|en nombre de|Grupo Umana"
Const kDelta = 1


Private m_sCredit() As String
Private m_lRowYpos As Long
Private m_lRowHeight As Long
Private m_lRegionHeight As Long
Private m_lRowWidth As Long
Private m_iUbound As Integer
Private m_iItemXform As Integer
Private m_iCurrent As Integer
Private m_iTik2Write As Integer

Private Sub Form_Load()

'para centrar el formulario; previamente poner AutoShowChildren = False del form MDI
Me.Height = 3810
Me.Width = 4350
Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2
'---------------------------------------------------------
       
    m_sCredit = Split(kCredits, "|")
    m_iUbound = UBound(m_sCredit)
    
    m_lRowHeight = Me.TextHeight("")
    m_lRowWidth = Me.ScaleWidth
    
    m_iItemXform = ((Me.ScaleHeight \ m_lRowHeight) \ 2) - 1
    m_lRowYpos = Me.ScaleHeight - (2 * m_lRowHeight)
    m_lRegionHeight = Me.ScaleHeight
        
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    TickScroll
End Sub

Private Sub TickScroll()

    Dim lRet As Long
        
    lRet = BitBlt(Me.hDC, 0, -kDelta, m_lRowWidth, m_lRegionHeight, Me.hDC, 0, 0, SRCCOPY)
        
    If m_iTik2Write = 0 Then PrintCurrent
    
    m_iTik2Write = m_iTik2Write + 1
    If m_iTik2Write = 25 Then m_iTik2Write = 0

End Sub
Private Sub PrintCurrent()

    Dim lWidth As Long
    
    If m_iCurrent >= 0 And m_iCurrent <= m_iUbound Then
        lWidth = Me.TextWidth(m_sCredit(m_iCurrent))
        Me.CurrentX = (m_lRowWidth - lWidth) \ 2
        Me.CurrentY = m_lRowYpos
        Me.Print m_sCredit(m_iCurrent)
    End If
    m_iCurrent = m_iCurrent + 1
    If m_iCurrent > m_iUbound Then m_iCurrent = -m_iItemXform

End Sub
