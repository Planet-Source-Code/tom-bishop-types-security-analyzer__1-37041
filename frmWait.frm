VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TypeS"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "frmWait.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   485
         Left            =   120
         Picture         =   "frmWait.frx":1703
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   90
         Width           =   485
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   135
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   840
         TabIndex        =   3
         Top             =   75
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " - Unloading bots..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   340
         Width           =   1650
      End
   End
   Begin VB.Timer tmrPercent 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4200
      Top             =   1440
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim i As Long
    
    On Error GoTo Err:
    
    stopSession = True
    Me.Show , frmMain
    DoEvents
    
    tmrPercent.Enabled = True
    pb1.Max = frmMain.lvMain.ListItems.Count
    For i = 1 To frmMain.lvMain.ListItems.Count
        While Not frmMain.lvMain.ListItems(i).ListSubItems(1).Text = "Session stopped."
            pb1.Value = i
            DoEvents
        Wend
    Next i
    
    Unload Me
    
    Exit Sub
    
Err:

    Unload Me

End Sub

Private Sub tmrPercent_Timer()

    Me.Caption = "TypeS - " & Percent(pb1.Value + 1, pb1.Max, pb1.Max) & "%"

End Sub
