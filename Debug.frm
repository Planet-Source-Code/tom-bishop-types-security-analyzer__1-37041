VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDebug 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Debugger"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7230
   Icon            =   "Debug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Output 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5953
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Debug.frx":57E2
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "HIDDEN"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "Save as..."
      End
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
    
    On Error Resume Next
    Output.Width = Me.ScaleWidth
    Output.Height = Me.ScaleHeight

End Sub

Private Sub Form_Resize()

    Form_Paint

End Sub

Private Sub mnuSave_Click()
    
    Dim fnum As Integer
    
    cd1.DialogTitle = "Save debug text as..."
    cd1.Filter = "Text files (*.txt)|*.txt|"
    cd1.ShowSave
    
    If Len(cd1.FileName) = 0 Then Exit Sub
    
    fnum = FreeFile
    
    Open cd1.FileName For Output As #fnum
        Print #fnum, Output.Text
    Close #fnum
    
    MsgBox "File saved as:" & vbCrLf & cd1.FileName, vbInformation

End Sub

Private Sub Output_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then PopupMenu mnuHidden

End Sub
