VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   240
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":2CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":84EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetup.frx":DCDE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   4440
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ButtonWidth     =   1402
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Bots"
      TabPicture(0)   =   "frmSetup.frx":10260
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkVerbose"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Proxy"
      TabPicture(1)   =   "frmSetup.frx":1027C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "chkProxy"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Sounds"
      TabPicture(2)   =   "frmSetup.frx":10298
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox chkVerbose 
         Caption         =   "Verbose mode"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   5775
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "This will open the debugger window on errors"
            Enabled         =   0   'False
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   3195
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   -74880
         TabIndex        =   21
         Top             =   1440
         Width           =   5775
         Begin VB.CheckBox chkFinished 
            Caption         =   "Play sound when logins found"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   0
            Width           =   2415
         End
         Begin VB.TextBox txtFinished 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   23
            Top             =   360
            Width           =   3375
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   330
            Left            =   4680
            TabIndex        =   22
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            ButtonWidth     =   1640
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Browse"
                  ImageIndex      =   3
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filename:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   5775
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   4680
            TabIndex        =   20
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            ButtonWidth     =   1640
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Browse"
                  ImageIndex      =   3
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.TextBox txtFound 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   19
            Top             =   360
            Width           =   3375
         End
         Begin VB.CheckBox chkFound 
            Caption         =   "Play sound when logins found"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   2415
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filename:"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bot timeout:"
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   5775
         Begin ComCtl2.UpDown UpDown2 
            Height          =   285
            Left            =   4920
            TabIndex        =   14
            Top             =   720
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   10
            BuddyControl    =   "txtBotTimeout"
            BuddyDispid     =   196621
            OrigLeft        =   5400
            OrigTop         =   1080
            OrigRight       =   5640
            OrigBottom      =   1335
            Max             =   60
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtBotTimeout 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4320
            TabIndex        =   13
            Text            =   "10"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sec."
            Height          =   195
            Left            =   5280
            TabIndex        =   15
            Top             =   720
            Width           =   330
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Timeout:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3480
            TabIndex        =   12
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "The more bots you use or the slower your connection is the higher you should set this setting."
            Height          =   495
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   5415
         End
      End
      Begin VB.CheckBox chkProxy 
         Caption         =   "Use proxies"
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   5775
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSetup.frx":102B4
            Enabled         =   0   'False
            Height          =   855
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   5535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Bots:"
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5775
         Begin ComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   5415
            TabIndex        =   5
            Top             =   720
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   50
            BuddyControl    =   "txtBots"
            BuddyDispid     =   196629
            OrigLeft        =   4320
            OrigTop         =   840
            OrigRight       =   4560
            OrigBottom      =   1095
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtBots 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "50"
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bots:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4200
            TabIndex        =   3
            Top             =   720
            Width           =   450
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "This is the amount of Bots that TypeS will use per session. The faster your connection is the higher you should set this number."
            Height          =   495
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   5535
         End
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFinished_Click()

    If chkFinished.Value = 1 Then
        Label10.Enabled = True
        txtFinished.Enabled = True
        Toolbar3.Enabled = True
    Else
        Label10.Enabled = False
        txtFinished.Enabled = False
        Toolbar3.Enabled = False
    End If

End Sub

Private Sub chkFound_Click()

    If chkFound.Value = 1 Then
        Label9.Enabled = True
        txtFound.Enabled = True
        Toolbar2.Enabled = True
    Else
        Label9.Enabled = False
        txtFound.Enabled = False
        Toolbar2.Enabled = False
    End If

End Sub

Private Sub chkProxy_Click()

    If chkProxy.Value = 0 Then
        Label3.Enabled = False
    Else
        Label3.Enabled = True
    End If

End Sub

Private Sub chkVerbose_Click()

    If chkVerbose.Value = 1 Then
        Label11.Enabled = True
    Else
        Label11.Enabled = False
    End If

End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    
    chkVerbose.Value = GetSetting("TYPES", "CONFIG", "VERBOSE")
    txtBots.Text = GetSetting("TYPES", "CONFIG", "BOTS")
    txtBotTimeout.Text = GetSetting("TYPES", "CONFIG", "TIMEOUT")
    UpDown1.Value = txtBots.Text
    UpDown2.Value = txtBotTimeout.Text
    chkProxy.Value = GetSetting("TYPES", "CONFIG", "USEPROXY")
    chkFound.Value = GetSetting("TYPES", "SOUNDS", "FOUND")
    txtFound.Text = GetSetting("TYPES", "SOUNDS", "FOUNDs")
    chkFinished.Value = GetSetting("TYPES", "SOUNDS", "FINISHED")
    txtFinished.Text = GetSetting("TYPES", "SOUNDS", "FINISHEDs")

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Caption
        
        Case "Save"
            Screen.MousePointer = 11
            
            If chkFound.Value = 1 Then
                If Len(txtFound.Text) = 0 Then
                    MsgBox "You need to enter a filename!", vbCritical
                    Screen.MousePointer = 1
                    Exit Sub
                End If
            End If
            
            If chkFinished.Value = 1 Then
                If Len(txtFinished.Text) = 0 Then
                    MsgBox "You need to enter a filename!", vbCritical
                    Screen.MousePointer = 1
                    Exit Sub
                End If
            End If
            
            SaveSetting "TYPES", "CONFIG", "VERBOSE", chkVerbose.Value
            SaveSetting "TYPES", "CONFIG", "BOTS", txtBots.Text
            SaveSetting "TYPES", "CONFIG", "TIMEOUT", txtBotTimeout.Text
            SaveSetting "TYPES", "CONFIG", "USEPROXY", chkProxy.Value
            SaveSetting "TYPES", "SOUNDS", "FOUND", chkFound.Value
            SaveSetting "TYPES", "SOUNDS", "FOUNDs", txtFound.Text
            SaveSetting "TYPES", "SOUNDS", "FINISHED", chkFound.Value
            SaveSetting "TYPES", "SOUNDS", "FINISHEDs", txtFound.Text
            
            If chkProxy.Value = 0 Then
                frmMain.lvProxy.Visible = False
            Else
                frmMain.lvProxy.Visible = True
            End If
            
            Screen.MousePointer = 1
            Unload Me
            
        Case "Close"
            Unload Me
        
    End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    cd1.DialogTitle = "Select sound file..."
    cd1.Filter = "Audio files (*.wav)|*.wav|"
    cd1.ShowOpen
    
    If Len(cd1.FileName) = 0 Then Exit Sub
    
    txtFound.Text = cd1.FileName
    PlayWav cd1.FileName
    
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

    cd1.DialogTitle = "Select sound file..."
    cd1.Filter = "Audio files (*.wav)|*.wav|"
    cd1.ShowOpen
    
    If Len(cd1.FileName) = 0 Then Exit Sub
    
    txtFinished.Text = cd1.FileName
    PlayWav cd1.FileName

End Sub
