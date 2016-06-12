VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frm_ole_document 
   Caption         =   "frmDocument"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   Icon            =   "frm_ole_document.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   5040
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfText_back 
      Height          =   1995
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frm_ole_document.frx":0ECA
   End
   Begin VB.OLE rtfText 
      Class           =   "Word.Document.8"
      Height          =   1155
      Left            =   120
      OleObjectBlob   =   "frm_ole_document.frx":0F4C
      TabIndex        =   1
      Top             =   480
      Width           =   3075
   End
End
Attribute VB_Name = "frm_ole_document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modificado As Boolean
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If modificado Then
    msg = MsgBox("¿Desea guardar los cambios?", vbYesNoCancel, "Guardar")
    
    If msg = vbYes Then

        'guardar---evento___MDIForm1.mnuFileSave_Click
        
        Dim sFile As String
        If Left$(MDIForm1.ActiveForm.Caption, 8) = "Document" Then
            With MDIForm1.dlgCommonDialog
                .DialogTitle = "Guardar"
                .CancelError = False
                'Pendiente: establecer los indicadores y atributos del control common dialog
                .Filter = "Todos los archivos (*.rtf)|*.rtf"
                .ShowSave
                If Len(.FileName) = 0 Then
                    Exit Sub
                End If
                sFile = .FileName
            End With
            MDIForm1.ActiveForm.rtfText.SaveFile sFile
        Else
            sFile = MDIForm1.ActiveForm.Caption
            MDIForm1.ActiveForm.rtfText.SaveFile sFile
        End If
        '--------------------------
        
        If MDIForm1.CantVentDoc = 1 Then
        
            MDIForm1.tbToolBarDoc.Visible = False
            MDIForm1.tbToolBar.Visible = True
            
            MDIForm1.mnuFileNew.Visible = False
            MDIForm1.mnuFileOpen.Visible = False
            MDIForm1.mnuFileSave.Visible = False
            MDIForm1.mnuFileSaveAs.Visible = False
            MDIForm1.mnuFileBar1.Visible = False
            MDIForm1.mnuFilePageSetup.Visible = False
            MDIForm1.mnuFilePrintPreview.Visible = False
            MDIForm1.mnuFilePrint.Visible = False
            MDIForm1.mnuFileBar2.Visible = False
            MDIForm1.mnuEdit.Visible = False
            
            MDIForm1.mnutexto.Checked = False
            MDIForm1.mnuherramientas.Checked = True
        
            MDIForm1.CantVentDoc = 0
            
            MDIForm1.mnuInforme.Enabled = True
        Else
            
            MDIForm1.CantVentDoc = MDIForm1.CantVentDoc - 1
        
        End If
        
        
    Else
        If msg = vbCancel Then
            'establece si se realiza la accion por defecto (0=salir) o no (1=cancelar "salir")
            Cancel = 1
            
        Else
        
            If MDIForm1.CantVentDoc = 1 Then
            
                MDIForm1.tbToolBarDoc.Visible = False
                MDIForm1.tbToolBar.Visible = True
                
                MDIForm1.mnuFileNew.Visible = False
                MDIForm1.mnuFileOpen.Visible = False
                MDIForm1.mnuFileSave.Visible = False
                MDIForm1.mnuFileSaveAs.Visible = False
                MDIForm1.mnuFileBar1.Visible = False
                MDIForm1.mnuFilePageSetup.Visible = False
                MDIForm1.mnuFilePrintPreview.Visible = False
                MDIForm1.mnuFilePrint.Visible = False
                MDIForm1.mnuFileBar2.Visible = False
                MDIForm1.mnuEdit.Visible = False
                
                MDIForm1.mnutexto.Checked = False
                MDIForm1.mnuherramientas.Checked = True
            
                MDIForm1.CantVentDoc = 0
                
                MDIForm1.mnuInforme.Enabled = True
            Else
                
                MDIForm1.CantVentDoc = MDIForm1.CantVentDoc - 1
            
            End If

        End If
    End If

Else

    If MDIForm1.CantVentDoc = 1 Then
    
        MDIForm1.tbToolBarDoc.Visible = False
        MDIForm1.tbToolBar.Visible = True
        
        MDIForm1.mnuFileNew.Visible = False
        MDIForm1.mnuFileOpen.Visible = False
        MDIForm1.mnuFileSave.Visible = False
        MDIForm1.mnuFileSaveAs.Visible = False
        MDIForm1.mnuFileBar1.Visible = False
        MDIForm1.mnuFilePageSetup.Visible = False
        MDIForm1.mnuFilePrintPreview.Visible = False
        MDIForm1.mnuFilePrint.Visible = False
        MDIForm1.mnuFileBar2.Visible = False
        MDIForm1.mnuEdit.Visible = False
        
        MDIForm1.mnutexto.Checked = False
        MDIForm1.mnuherramientas.Checked = True
    
        MDIForm1.CantVentDoc = 0
        
        MDIForm1.mnuInforme.Enabled = True
    Else
            
        MDIForm1.CantVentDoc = MDIForm1.CantVentDoc - 1
        
    End If
    
End If
                


End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub

'Private Sub rtfText_SelChange()
'    MDIForm1.tbToolBarDoc.Buttons("Negrita").Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
'    MDIForm1.tbToolBarDoc.Buttons("Cursiva").Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
'    MDIForm1.tbToolBarDoc.Buttons("Subrayado").Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
'    MDIForm1.tbToolBarDoc.Buttons("Alinear a la izquierda").Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
'    MDIForm1.tbToolBarDoc.Buttons("Centrar").Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
'    MDIForm1.tbToolBarDoc.Buttons("Alinear a la derecha").Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
'
'modificado = True
'End Sub

Private Sub Form_Load()
    Form_Resize
    
    modificado = False
    
    'rtfText.CreateEmbed ("D:\Dietetica\reportes\Rep_1_2005_2_15_.doc")
    'rtfText.CreateLink ("D:\Dietetica\reportes\Rep_1_2005_2_15_.doc")

End Sub


Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    'rtfText.RightMargin = rtfText.Width - 400
End Sub

