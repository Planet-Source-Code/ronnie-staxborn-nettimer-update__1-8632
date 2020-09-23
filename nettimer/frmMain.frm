VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetTimer 1.0"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   2520
      TabIndex        =   11
      Top             =   120
      Width           =   2295
      Begin VB.Label Label3 
         Caption         =   "rompa@hem.passagen.se"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Code added for Hrs, Min and seconds by Ronnie Staxborn"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Data"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   1320
   End
   Begin VB.Frame fraProductInfo 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.Label lblRights 
         BackStyle       =   0  'Transparent
         Caption         =   "All Rights Reserved"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblEMail 
         BackStyle       =   0  'Transparent
         Caption         =   "info@virtual-dev.de"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblProductName 
         BackStyle       =   0  'Transparent
         Caption         =   "NetTimer 1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "© Copyright Fabian Fischer"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblURL 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.virtual-dev.de"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time spend online"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LISTENING - OFFLINE"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblFirstConnection 
      BackStyle       =   0  'Transparent
      Caption         =   "First dial-up connection: XXXXXX"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label lblOnlineMinutes 
      BackStyle       =   0  'Transparent
      Caption         =   "XXX minutes online!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************
'© Copyright 2000 Virtual-Dev, Fabian Fischer, All Rights Reserved.
'**************************************************************************
'Name:              NetTimer 1.0
'Modification:      2000-05-30
'E-Mail:            info@virtual-dev.de
'Internet:          http://www.virtual-dev.de
'**************************************************************************
'You are free to use this code within your own applications, but you are
'expressly forbidden from selling or otherwise distributing this source
'code without prior written consent. This includes both posting free demo
'projects made from this code as well as reproducing the code in text or
'html format.
'**************************************************************************

'**************************************************************************
'Some codes are Added by Ronnie Staxborn (se code below for more info)
'**************************************************************************

Dim mins As Integer
Dim hrs As Integer
Dim Secs As Integer
Dim backup

Private Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long

Private Const RAS95_MaxEntryName = 256
Private Const RAS95_MaxDeviceType = 16
Private Const RAS95_MaxDeviceName = 32

Private Type RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Private Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Private bOnlineNow As Boolean 'Am I online now?
Private lngOnlineSeconds As Long

Private Function IsConnected() As Integer
Dim TRasCon(255) As RASCONN95
Dim lngA As Long
Dim lngB As Long
Dim lngRetVal As Long
Dim TStatus As RASCONNSTATUS95

TRasCon(0).dwSize = 412
lngA = 256 * TRasCon(0).dwSize
lngRetVal = RasEnumConnections(TRasCon(0), lngA, lngB)

If lngB = 0 Then
    IsConnected = 0 'offline
    Exit Function
End If

TStatus.dwSize = 160
lngRetVal = RasGetConnectStatus(TRasCon(0).hRasCon, TStatus)

If TStatus.RasConnState = &H2000 Then
    IsConnected = 2 'online
Else
    IsConnected = 1
End If

End Function

Private Sub cmdReset_Click()
On Error Resume Next
Beep
If MsgBox("Do you really want to reset the settings?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Reset Settings?") = vbYes Then
    DeleteSetting "NetTimer", "General", "FirstConnection"
    DeleteSetting "NetTimer", "General", "OnlineSeconds"
    DeleteSetting "NetTimer", "General"
    DeleteSetting "NetTimer"
    Kill AppPath & "NetTimer.log"
    Form_Load
End If
End Sub

Private Sub Form_Load()

'If NetTimer is already started!
If App.PrevInstance = True Then
    Beep
    MsgBox "NetTimer 1.0 is already running!", vbExclamation, "NetTimer"
    End 'Quit Now
End If

bOnlineNow = False

If GetSetting("NetTimer", "General", "FirstConnection", "0") = "0" Then
    lblFirstConnection.Visible = False
    
    lblOnlineMinutes.Visible = False
Else
    
    lngOnlineSeconds = CLng(GetSetting("NetTimer", "General", "OnlineSeconds", 0))
    
    lblFirstConnection = "First dial-up connection: " & GetSetting("NetTimer", "General", "FirstConnection", Now)
    lblFirstConnection.Visible = True
    
            '-------------------------------------
            '    Added code by Ronnie Staxborn
            '-------------------------------------
           
            backup = lngOnlineSeconds
            If backup > 3600 Then Let hrs = Int(backup / 3600): backup = backup - (3600 * hrs) '3600 is seconds in a minute
            If backup > 60 Then Let mins = Int(backup / 60): backup = backup - (60 * mins)
            Secs = Int(backup)
                        
            lblOnlineMinutes = hrs & ":" & mins & ":" & Secs 'Int(lngOnlineSeconds / 60) & " minutes online!"
            
            '------------------------------------
            '           End added code
            '------------------------------------
            
    lblOnlineMinutes.Visible = True
End If

If Trim(LCase(Command)) = "/hide" Then Me.Hide

tmrTimer.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
            '-------------------------------------
            '    Added code by Ronnie Staxborn
            '-------------------------------------
            Dim intFreeFile As Integer
            intFreeFile = FreeFile
            
            Open AppPath & "NetTimer.log" For Append As #intFreeFile
                Print #intFreeFile, "! " & Now & " SHUTDOWN!"
            Close #intFreeFile
            
            '------------------------------------
            '           End added code
            '------------------------------------
End Sub

Private Sub tmrTimer_Timer()
Dim intFreeFile As Integer

Select Case IsConnected
    Case 0 'offline
        If bOnlineNow = True Then
            
            intFreeFile = FreeFile
            
            Open AppPath & "NetTimer.log" For Append As #intFreeFile
                Print #intFreeFile, "- " & Now & " OFFLINE!"
            Close #intFreeFile
            lblStatus = "LISTENING - OFFLINE"
            
            bOnlineNow = False
            
        End If
        
    Case 2 'online
        If bOnlineNow = False Then
            If lngOnlineSeconds = 0 Then
                SaveSetting "NetTimer", "General", "FirstConnection", Now
                lblFirstConnection = "First dial-up connection: " & GetSetting("NetTimer", "General", "FirstConnection", Now)
                lblFirstConnection.Visible = True
            End If

            intFreeFile = FreeFile
            
            Open AppPath & "NetTimer.log" For Append As #intFreeFile
                Print #intFreeFile, "+ " & Now & " ONLINE!"
            Close #intFreeFile
            
            lblStatus = "LISTENING - ONLINE"
            
            bOnlineNow = True
        Else
            lngOnlineSeconds = lngOnlineSeconds + 1
            SaveSetting "NetTimer", "General", "OnlineSeconds", lngOnlineSeconds
           
            '-------------------------------------
            '    Added code by Ronnie Staxborn
            '-------------------------------------
           
            backup = lngOnlineSeconds
            If backup > 3600 Then Let hrs = Int(backup / 3600): backup = backup - (3600 * hrs) '3600 is seconds in a minute
            If backup > 60 Then Let mins = Int(backup / 60): backup = backup - (60 * mins)
            Secs = Int(backup)
                                
            lblOnlineMinutes = hrs & ":" & mins & ":" & Secs 'Int(lngOnlineSeconds / 60) & " minutes online!"
            
            '------------------------------------
            '           End added code
            '------------------------------------
            
            lblOnlineMinutes.Visible = True
        End If
End Select

End Sub

Function AppPath() As String
AppPath = App.Path
If Right(AppPath, 1) <> "\" Then
    AppPath = AppPath & "\"
End If
End Function
