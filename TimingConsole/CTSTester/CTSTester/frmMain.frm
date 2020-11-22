VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CTS Timing Console"
   ClientHeight    =   8760
   ClientLeft      =   2805
   ClientTop       =   3075
   ClientWidth     =   8715
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   8715
   Begin VB.Frame fraCTS 
      Caption         =   "CTS Console"
      Height          =   6735
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   8535
      Begin VB.CheckBox chkSplits 
         Caption         =   "Splits"
         Height          =   255
         Left            =   6000
         TabIndex        =   49
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkJudging 
         Caption         =   "Judging"
         Height          =   255
         Left            =   4680
         TabIndex        =   48
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chkIndividual 
         Caption         =   "Individual"
         Height          =   255
         Left            =   3240
         TabIndex        =   47
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkBackup 
         Caption         =   "Backup"
         Height          =   255
         Left            =   2040
         TabIndex        =   46
         Top             =   1080
         Width           =   1095
      End
      Begin VB.HScrollBar HScrollRace 
         Height          =   255
         LargeChange     =   10
         Left            =   2040
         Max             =   500
         Min             =   1
         TabIndex        =   38
         Top             =   1440
         Value           =   1
         Width           =   4575
      End
      Begin VB.HScrollBar HScrollEvent 
         Height          =   255
         LargeChange     =   10
         Left            =   2040
         Max             =   255
         Min             =   1
         TabIndex        =   37
         Top             =   1680
         Value           =   1
         Width           =   4575
      End
      Begin VB.HScrollBar HScrollHeat 
         Height          =   255
         Left            =   2040
         Max             =   16
         Min             =   1
         TabIndex        =   36
         Top             =   1920
         Value           =   1
         Width           =   4575
      End
      Begin VB.ComboBox cboSetupSubCommand 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Top             =   720
         Width           =   4575
      End
      Begin VB.CommandButton btnCTSClear 
         Caption         =   "Clear"
         Height          =   285
         Left            =   7440
         TabIndex        =   2
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox txtCTSResponse 
         Height          =   2925
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3120
         Width           =   8055
      End
      Begin VB.CommandButton btnCTSSend 
         Caption         =   "Send"
         Height          =   285
         Left            =   7440
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox cboSubCommand 
         Height          =   315
         Left            =   2040
         TabIndex        =   32
         Top             =   720
         Width           =   4575
      End
      Begin VB.ComboBox cboCommand 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Qualifications:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblRaceT 
         Caption         =   "Race:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblEventT 
         Caption         =   "Event:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblHeatT 
         Caption         =   "Heat:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblRace 
         Caption         =   "= 1"
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
         Left            =   6720
         TabIndex        =   41
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblEvent 
         Caption         =   "= 1"
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
         Left            =   6720
         TabIndex        =   40
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblHeat 
         Caption         =   "= 1"
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
         Left            =   6720
         TabIndex        =   39
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8400
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label5 
         Caption         =   "Sent:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblCTSSend 
         Caption         =   " "
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
         Left            =   960
         TabIndex        =   34
         Top             =   2520
         Width           =   7335
      End
      Begin VB.Label Label4 
         Caption         =   "Response:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblSubComm 
         Caption         =   "Sub Command:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblPrimary 
         Caption         =   "Primary Command:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraDebug 
      Caption         =   "Debugging"
      Height          =   1575
      Left            =   120
      TabIndex        =   25
      Top             =   6840
      Width           =   8535
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   5
         Left            =   3360
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   6
         Left            =   3840
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   7
         Left            =   4320
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   8
         Left            =   4800
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   9
         Left            =   5280
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   10
         Left            =   5760
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtHEX 
         Height          =   285
         Index           =   11
         Left            =   6240
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton btnSend 
         Caption         =   "Send"
         Height          =   285
         Index           =   0
         Left            =   6720
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtTextOut 
         Height          =   285
         Left            =   960
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   720
         Width           =   5655
      End
      Begin VB.CommandButton btnSend 
         Caption         =   "Send"
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtResponse 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   23
         Top             =   1080
         Width           =   5655
      End
      Begin VB.CommandButton btnClear 
         Caption         =   "Clear"
         Height          =   285
         Index           =   0
         Left            =   7440
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton btnClear 
         Caption         =   "Clear"
         Height          =   285
         Index           =   1
         Left            =   7440
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton btnClear 
         Caption         =   "Clear"
         Height          =   285
         Index           =   2
         Left            =   7440
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "HEX:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "TEXT:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Response:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8505
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   450
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2187
            MinWidth        =   2187
            Text            =   "Connected"
            TextSave        =   "Connected"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1270
            MinWidth        =   1270
            Text            =   "com1"
            TextSave        =   "com1"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11809
            MinWidth        =   2258
            Text            =   "9600 8N1"
            TextSave        =   "9600 8N1"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuPort 
         Caption         =   "&Port"
         Begin VB.Menu mnuCom 
            Caption         =   "Com &1"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuCom 
            Caption         =   "Com &2"
            Index           =   2
         End
         Begin VB.Menu mnuCom 
            Caption         =   "Com &3"
            Index           =   3
         End
         Begin VB.Menu mnuCom 
            Caption         =   "Com &4"
            Index           =   4
         End
         Begin VB.Menu mnuCom 
            Caption         =   "Com &5"
            Index           =   5
         End
      End
      Begin VB.Menu mnuSpeed 
         Caption         =   "&Speed"
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "11&0"
            Index           =   0
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "&300"
            Index           =   1
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "&600"
            Index           =   2
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "&1,200"
            Index           =   3
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "&2,400"
            Index           =   4
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "&9,600"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "14,400"
            Index           =   6
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "19,200"
            Index           =   7
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "28,800"
            Index           =   8
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "3&8,400"
            Index           =   9
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "56,000"
            Index           =   10
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "128,000"
            Index           =   11
         End
         Begin VB.Menu mnuSpeedSel 
            Caption         =   "256,000"
            Index           =   12
         End
      End
      Begin VB.Menu mnuDataB 
         Caption         =   "&Data Bits"
         Begin VB.Menu mnuDataBSel 
            Caption         =   "&4"
            Index           =   4
         End
         Begin VB.Menu mnuDataBSel 
            Caption         =   "&5"
            Index           =   5
         End
         Begin VB.Menu mnuDataBSel 
            Caption         =   "&6"
            Index           =   6
         End
         Begin VB.Menu mnuDataBSel 
            Caption         =   "&7"
            Index           =   7
         End
         Begin VB.Menu mnuDataBSel 
            Caption         =   "&8"
            Checked         =   -1  'True
            Index           =   8
         End
      End
      Begin VB.Menu mnuParity 
         Caption         =   "&Parity"
         Begin VB.Menu mnuParitySel 
            Caption         =   "&Even"
            Index           =   0
         End
         Begin VB.Menu mnuParitySel 
            Caption         =   "&Mark"
            Index           =   1
         End
         Begin VB.Menu mnuParitySel 
            Caption         =   "&None"
            Index           =   2
         End
         Begin VB.Menu mnuParitySel 
            Caption         =   "&Odd"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu mnuParitySel 
            Caption         =   "&Space"
            Index           =   4
         End
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop Bits"
         Begin VB.Menu mnuStopSel 
            Caption         =   "&1"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuStopSel 
            Caption         =   "1.&5"
            Index           =   1
         End
         Begin VB.Menu mnuStopSel 
            Caption         =   "&2"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpSel 
         Caption         =   "&Basic Help"
         Index           =   0
      End
      Begin VB.Menu mnuHelpSel 
         Caption         =   "&About"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Basic Colorado Timing Console
'Written by: G. Wood
'
Option Explicit

'Dim SampOb As SampleObject
Dim SerialOb As CommIO
Dim SerialPort As Integer
Dim SerialBaud As String
Dim SerialParity As String
Dim SerialData As String
Dim SerialStop As String

Const MAX_PORTS As Integer = 5

Const bDebug As Boolean = False

Private Sub btnClear_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0:
            For i = 0 To 11
                txtHEX(i) = ""
            Next i
        Case 1:
            txtTextOut = ""
        Case 2:
            txtResponse = ""
    End Select
    
End Sub

Private Sub btnCTSClear_Click()
    txtCTSResponse.Text = ""
End Sub

Private Sub btnCTSSend_Click()
    Dim OutString As String
    Dim OutString2 As String
    Dim Qualifiers As String
         
    Select Case cboCommand.ListIndex
        Case 0 ' Read
            Select Case cboSetupSubCommand.ListIndex
                Case 0 ' d
                    OutString = SerialOb.Hex2ASCII("06005264")
                Case 1 ' e
                    OutString = SerialOb.Hex2ASCII("06005265")
                Case 2 ' f
                    OutString = SerialOb.Hex2ASCII("06005266")
                Case 3 ' g
                    OutString = SerialOb.Hex2ASCII("06005267")
                Case 4 ' h
                    OutString = SerialOb.Hex2ASCII("06005268")
                Case 5 ' i
                    OutString = SerialOb.Hex2ASCII("06005269")
                Case 6 ' j
                    OutString = SerialOb.Hex2ASCII("0600526A")
                Case 7 ' k
                    OutString = SerialOb.Hex2ASCII("0600526B")
                Case 8 ' l
                    OutString = SerialOb.Hex2ASCII("0600526C")
                Case 9 ' m
                    OutString = SerialOb.Hex2ASCII("0600526D")
                Case 10 ' n
                    OutString = SerialOb.Hex2ASCII("0600526E")
                Case 11 ' o
                    OutString = SerialOb.Hex2ASCII("0600526F")
                Case 12 ' p
                    OutString = SerialOb.Hex2ASCII("06005270")
                Case 13 ' q
                    OutString = SerialOb.Hex2ASCII("06005271")
                Case 14 ' r
                    OutString = SerialOb.Hex2ASCII("06005272")
                Case 15 ' s
                    OutString = SerialOb.Hex2ASCII("06005273")
                Case 16 ' t
                    OutString = SerialOb.Hex2ASCII("08005274" & ToHex(HScrollRace.Value) & ToHex(HScrollEvent.Value))
                Case 17 ' u
                    OutString = SerialOb.Hex2ASCII("06005275")
                Case 18 ' v
                    OutString = SerialOb.Hex2ASCII("06005276")
                Case 19 ' w
                    OutString = SerialOb.Hex2ASCII("06005277")
                Case 20 ' x
                    OutString = SerialOb.Hex2ASCII("06005278")
            End Select
        
        Case 1 ' Write
            MsgBox "Write disabled!", vbCritical, App.Title
            Exit Sub
            'OutString = SerialOb.Hex2ASCII("050049")
        
        Case 2 ' Move +
            OutString = SerialOb.Hex2ASCII("06004D2B")
        
        Case 3 ' Move -
            OutString = SerialOb.Hex2ASCII("06004D2D")
        
        Case 4 ' Swim Data
            Qualifiers = "53"
            If chkBackup.Value = 1 Then
                Qualifiers = Qualifiers & "42"
            End If
            If chkIndividual.Value = 1 Then
                Qualifiers = Qualifiers & "49"
            End If
            If chkJudging.Value = 1 Then
                Qualifiers = Qualifiers & "4A"
            End If
            If chkSplits.Value = 1 Then
                Qualifiers = Qualifiers & "53"
            End If
            Select Case cboSubCommand.ListIndex
                Case 1 ' C
                    OutString2 = Qualifiers & "43"
                    OutString = SerialOb.Hex2ASCII(L2H(OutString2, 4) & OutString2) '06005343
                Case 0 ' E#$
                    OutString2 = Qualifiers & "45"
                    OutString = SerialOb.Hex2ASCII(L2H(OutString2, 6) & OutString2 & ToHex(HScrollHeat.Value) & ToHex(HScrollEvent.Value)) '08005345
                Case 2 ' L
                    OutString2 = Qualifiers & "4C"
                    OutString = SerialOb.Hex2ASCII(L2H(OutString2, 4) & OutString2) '0600534C
                Case 3 ' N
                    OutString2 = Qualifiers & "4E"
                    OutString = SerialOb.Hex2ASCII(L2H(OutString2, 3) & OutString2) '0600534E
                Case 4 ' R##
                    OutString2 = Qualifiers & "5200"
                    OutString = SerialOb.Hex2ASCII(L2H(OutString2, 8) & OutString2 & ToHex(HScrollRace.Value)) '0800535200
            End Select
        
        Case 5 ' Time
            OutString = SerialOb.Hex2ASCII("050054") '050054
                    
        Case 6 ' Who are you
            OutString = SerialOb.Hex2ASCII("050057") ' 050057A3FF
    
    End Select
    
    OutString = OutString & SerialOb.GetDIC(OutString)
    
    Dim c As Integer
    Dim i As Integer
    Dim tStr As String
    
    lblCTSSend.Caption = SerialOb.ASCII2Hex(OutString, True)
        
    Send OutString
    
    tStr = Read
    
    If Len(tStr) > 0 Then
        txtCTSResponse.Text = tStr & vbCrLf
    Else
        txtCTSResponse.Text = txtCTSResponse.Text & "No response!" & vbCrLf
    End If
    
End Sub

Private Function L2H(s As String, i As Integer) As String
    L2H = ToHex((Len(s) / 2) + i) & "00"
End Function

Private Function ReverseString(ByVal InputString As String) As String

Dim lLen As Long, lCtr As Long
Dim sChar As String
Dim sAns As String

lLen = Len(InputString)
For lCtr = lLen To 1 Step -1
    sChar = Mid(InputString, lCtr, 1)
    sAns = sAns & sChar
Next

ReverseString = sAns

End Function

Private Function ToHex(c As Integer) As String
    If c = 0 Then
        ToHex = "00"
    Else
        ToHex = Hex(c)
    End If
    If Len(ToHex) < 2 Then
        ToHex = "0" & ToHex
    End If
End Function

Private Sub btnSend_Click(Index As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim Temp As Long
    Dim OutString As String
    
    OutString = ""
    Select Case Index
        Case 0                      ' Send Hex character string
            For i = 0 To 11
                Temp = 0
                If (Len(txtHEX(i)) > 0) Then
                    For j = 1 To Len(Left(txtHEX(i), 2))
                        Temp = Temp * 16 + HexChar(Mid(txtHEX(i), j, 1))
                    Next j
                    OutString = OutString & Chr(Temp)
                End If
            Next i
            Send (OutString)
        Case 1                      ' Send ascii string
          Send txtTextOut.Text
    End Select
    
    Dim lngStatus As Long
    Dim strError As String
    Dim c As Integer

    lngStatus = SerialOb.CommRead(SerialPort, OutString, 512)
    txtResponse.Text = SerialOb.ASCII2Hex(OutString, False)
        
End Sub
  
Private Sub Send(txtin As String)
Dim lngSize As Long
Dim lngStatus As Long
Dim strError As String

    If SerialOb.PortStatus(SerialPort).blnPortOpen = True Then
        ' Write data to serial port.
        lngSize = Len(txtin)
        lngStatus = SerialOb.CommWrite(SerialPort, txtin)
        If lngStatus <> lngSize Then
            lngStatus = SerialOb.CommGetError(strError)
            MsgBox "COM" & CStr(SerialPort) & " Error: " & strError
        End If
    End If
    
End Sub

Private Function Read() As String
Dim lngSize As Long
Dim lngStatus As Long
Dim strError As String
Dim OutString As String
Dim c As Integer
Dim i As Integer

    lngStatus = SerialOb.CommRead(SerialPort, OutString, 512)
    If lngStatus <> -1 Then
        Read = SerialOb.ASCII2Hex(OutString, True)
    Else
        SerialOb.CommGetError Read
    End If
    
End Function

Function HexChar(strData As String) As Integer

    Select Case strData
        Case "0" To "9"
            HexChar = Val(strData) - Val("0")
        Case "a", "A"
            HexChar = 10
        Case "b", "B"
            HexChar = 11
        Case "c", "C"
            HexChar = 12
        Case "d", "D"
            HexChar = 13
        Case "e", "E"
            HexChar = 14
        Case "f", "F"
            HexChar = 15
        Case Else
            HexChar = 0
    End Select

End Function

Private Sub Load_cboCommand()

    cboCommand.AddItem "Read Setup (R)"
    cboCommand.ItemData(cboCommand.NewIndex) = 1
    cboCommand.AddItem "Write Setup (I)"
    cboCommand.ItemData(cboCommand.NewIndex) = 2
    cboCommand.AddItem "Move Meet Pointer (M+)"
    cboCommand.ItemData(cboCommand.NewIndex) = 3
    cboCommand.AddItem "Move Meet Pointer (M-)"
    cboCommand.ItemData(cboCommand.NewIndex) = 4
    cboCommand.AddItem "Swimming Data (S)"
    cboCommand.ItemData(cboCommand.NewIndex) = 5
    cboCommand.AddItem "Time and Date of Meet (T)"
    cboCommand.ItemData(cboCommand.NewIndex) = 6
    cboCommand.AddItem "Who Are You (W)"
    cboCommand.ItemData(cboCommand.NewIndex) = 7
    cboCommand.ListIndex = 0

End Sub

Private Sub Load_cboSubCommand()


    cboSubCommand.AddItem "Send Event and Heat (E#$)"
    cboSubCommand.ItemData(cboSubCommand.NewIndex) = 1
    cboSubCommand.AddItem "Send Current Times (C)"
    cboSubCommand.ItemData(cboSubCommand.NewIndex) = 2
    cboSubCommand.AddItem "Send Last Times (L)"
    cboSubCommand.ItemData(cboSubCommand.NewIndex) = 3
    cboSubCommand.AddItem "Send Next Times (N)"
    cboSubCommand.ItemData(cboSubCommand.NewIndex) = 4
    cboSubCommand.AddItem "Send Race number (R##)"
    cboSubCommand.ItemData(cboSubCommand.NewIndex) = 5
    
    cboSubCommand.Enabled = True
    cboSubCommand.Visible = True
    
    cboSubCommand.ListIndex = 0

End Sub

Private Sub Load_cboSetupSubCommand()

    cboSetupSubCommand.AddItem "Start source (d)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 1
    cboSetupSubCommand.AddItem "Finish/Buttons (e)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 2
    cboSetupSubCommand.AddItem "Hardware (f)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 3
    cboSetupSubCommand.AddItem "Splits (g)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 4
    cboSetupSubCommand.AddItem "Timing (h)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 5
    cboSetupSubCommand.AddItem "Pool (i)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 6
    cboSetupSubCommand.AddItem "Scoreboard Options (j)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 7
    cboSetupSubCommand.AddItem "Scoreboard Definitions (k)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 8
    cboSetupSubCommand.AddItem "Scoreboard One-Line (l)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 9
    cboSetupSubCommand.AddItem "Printer Names (m)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 10
    cboSetupSubCommand.AddItem "Printer Options (n)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 11
    cboSetupSubCommand.AddItem "Printer User Defined Codes (o)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 12
    cboSetupSubCommand.AddItem "Printer Store Printer Format (p)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 13
    cboSetupSubCommand.AddItem "Printer Sponsor's Message (q)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 14
    cboSetupSubCommand.AddItem "Event Sequence Select (r)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 15
    cboSetupSubCommand.AddItem "Event Sequence Change Label (s)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 16
    cboSetupSubCommand.AddItem "Event Sequence Types (t)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 17
    cboSetupSubCommand.AddItem "Event Sequence Change (u)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 18
    cboSetupSubCommand.AddItem "Time of Day (v)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 19
    cboSetupSubCommand.AddItem "Day Date (w)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 20
    cboSetupSubCommand.AddItem "Time Display Mode (x)"
    cboSetupSubCommand.ItemData(cboSetupSubCommand.NewIndex) = 21
      
    cboSetupSubCommand.ListIndex = 0
    
    cboSetupSubCommand.Enabled = True
    cboSetupSubCommand.Visible = True

End Sub

Private Sub cboCommand_Click()
       
    HScrollEvent.Enabled = False
    HScrollRace.Enabled = False
    HScrollHeat.Enabled = False
    chkBackup.Enabled = False
    chkJudging.Enabled = False
    chkSplits.Enabled = False
    chkIndividual.Enabled = False
    
    If cboCommand.ListIndex < 2 Then
        cboSetupSubCommand.Visible = True
        cboSubCommand.Visible = False
    ElseIf cboCommand.ListIndex = 4 Then
        cboSubCommand.Visible = True
        cboSetupSubCommand.Visible = False
        cboSubCommand.Enabled = True
        chkBackup.Enabled = True
        chkJudging.Enabled = True
        chkSplits.Enabled = True
        chkIndividual.Enabled = True
    Else
        cboSubCommand.Visible = True
        cboSubCommand.Enabled = False
        cboSetupSubCommand.Visible = False
    End If
    
    cboSubCommand_Click
    
End Sub


Private Sub cboSetupSubCommand_Click()
    If cboSetupSubCommand.Enabled = True Then
        HScrollEvent.Enabled = False
        HScrollRace.Enabled = False
        HScrollHeat.Enabled = False
        lblRaceT.Caption = "Race:"
        lblEventT.Caption = "Event:"
        lblHeatT.Caption = "Heat:"
        HScrollRace.Max = 500
        HScrollEvent.Max = 255
        HScrollRace.Min = 1
        HScrollEvent.Min = 1
        HScrollRace.Value = 1
        HScrollEvent.Value = 1
        If cboSetupSubCommand.ListIndex = 16 Then
            HScrollRace.Enabled = True
            HScrollEvent.Enabled = True
            HScrollRace.Max = 3
            HScrollEvent.Max = 19
            HScrollRace.Min = 0
            HScrollEvent.Min = 0
            HScrollRace.Value = 0
            HScrollEvent.Value = 0
            lblRaceT.Caption = "Type:"
            lblEventT.Caption = "Seq:"
            lblHeatT.Caption = ""
        End If
    End If

End Sub

Private Sub cboSubCommand_Click()

    If cboSubCommand.Enabled = True Then
        HScrollEvent.Enabled = False
        HScrollRace.Enabled = False
        HScrollHeat.Enabled = False
        If cboSubCommand.ListIndex = 0 Then
            HScrollEvent.Enabled = True
            HScrollHeat.Enabled = True
        ElseIf cboSubCommand.ListIndex = 4 Then
            HScrollRace.Enabled = True
        End If
    End If
End Sub


Private Sub Form_Load()
    Dim i As Integer
    
    Load_cboCommand
    Load_cboSubCommand
    Load_cboSetupSubCommand
    
    'Set SampOb = New SampleObject
    Set SerialOb = New CommIO
    
    SerialPort = 1                  ' Set the default to COM1
    SerialBaud = "9600"
    SerialParity = "O"
    SerialStop = "1"
    SerialData = "8"
    
    For i = 0 To 2                  ' Clear all input/output fields
        btnClear_Click (i)
    Next i
    
    If Not ValidatePort Then
        MsgBox "There are no available comm ports on this computer.", , App.Title
    End If
    
    lblEvent.Caption = "= " & CStr(HScrollEvent.Value)
    lblHeat.Caption = "= " & CStr(HScrollHeat.Value)
    lblRace.Caption = "= " & CStr(HScrollRace.Value)
    HScrollEvent.Enabled = False
    HScrollRace.Enabled = False
    HScrollHeat.Enabled = False
    UpdateStatus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

    For i = 4 To 1 Step -1
        If (SerialOb.PortStatus(i).blnPortOpen = True) Then
            SerialOb.CommClose i
        End If
    Next i
    
    Set SerialOb = Nothing
    'Set SampOb = Nothing
        
End Sub


Private Sub HScrollEvent_Change()
    lblEvent.Caption = "= " & CStr(HScrollEvent.Value)
    DoEvents
End Sub

Private Sub HScrollHeat_Change()
    lblHeat.Caption = "= " & CStr(HScrollHeat.Value)
    DoEvents
End Sub

Private Sub HScrollRace_Change()
    If cboSetupSubCommand.ListIndex = 16 And cboSetupSubCommand.Enabled = True Then
        Select Case HScrollRace.Value
        Case 0
            lblRace.Caption = "= " & CStr(HScrollRace.Value) & " Swimmer Type"
        Case 1
            lblRace.Caption = "= " & CStr(HScrollRace.Value) & " Age Group"
        Case 2
            lblRace.Caption = "= " & CStr(HScrollRace.Value) & " Event Type"
        Case 3
            lblRace.Caption = "= " & CStr(HScrollRace.Value) & " Bracket"
        Case Else
            lblRace.Caption = "= " & CStr(HScrollRace.Value) & " Unknown"
        End Select
    Else
        lblRace.Caption = "= " & CStr(HScrollRace.Value)
    End If
    DoEvents
End Sub

Private Sub mnuCom_Click(Index As Integer)
    Dim i As Integer
    
    On Error Resume Next
   
    If SerialOb.PortStatus(SerialPort).blnPortOpen = True Then
        SerialOb.CommClose SerialPort
    End If
    
    SerialPort = Index
    mnuConnect_Click
    
    For i = 1 To MAX_PORTS
       mnuCom(i).Checked = False
    Next i
    
    mnuCom(Index).Checked = True
        
End Sub

Private Sub mnuConnect_Click()
Dim lngStatus As Long
Dim strError  As String

    On Error Resume Next
    
    If SerialOb.PortStatus(SerialPort).blnPortOpen = True Then
        SerialOb.CommClose SerialPort
        'MsgBox "COM" & CStr(SerialPort) & " not available!"
    Else
        lngStatus = SerialOb.CommOpen(SerialPort, "COM" & CStr(SerialPort), SerialString)
        If lngStatus <> 0 Then
        ' Handle error.
            lngStatus = SerialOb.CommGetError(strError)
            MsgBox "COM" & CStr(SerialPort) & " Error: " & strError
        End If
   End If
    
   UpdateStatus
   
End Sub

Private Function SerialString() As String
    SerialString = "baud=" & SerialBaud & " parity=" & SerialParity & " data=" & SerialData & " stop=" & SerialStop
End Function

Private Sub mnuDataBSel_Click(Index As Integer)
    Dim i As Integer
    Dim lngStatus As Long
    Dim strError  As String
    
    For i = 4 To 8
        If (i = Index) Then
            mnuDataBSel(i).Checked = True
            Select Case Index
                Case 4      ' 4
                    SerialData = "4"
                Case 5      ' 5
                    SerialData = "5"
                Case 6      ' 6
                    SerialData = "6"
                Case 7      ' 7
                    SerialData = "7"
                Case 8      ' 8
                    SerialData = "8"
            End Select
        Else
            mnuDataBSel(i).Checked = False
        End If
    Next i
    
    If SerialOb.PortStatus(SerialPort).blnPortOpen = True Then
        lngStatus = SerialOb.CommSet(SerialPort, SerialString)
        If lngStatus <> 0 Then
            ' Handle error.
            lngStatus = SerialOb.CommGetError(strError)
            MsgBox "COM" & CStr(SerialPort) & " Error: " & strError
        End If
    End If
    
    UpdateStatus

End Sub

Private Sub mnuHelpSel_Click(Index As Integer)
    Select Case Index
        Case 0      ' Basic Help
            MsgBox "Basic CTS Timing Console Communications." _
                           , vbInformation, "Help"
        Case 1      ' About
            MsgBox "Basic CTS Timing Console Communications Version 0.91", , "Help About"
    End Select
End Sub

Private Sub mnuParitySel_Click(Index As Integer)
    Dim i As Integer
    Dim lngStatus As Long
    Dim strError  As String
    
    For i = 0 To 4
        If (i = Index) Then
            mnuParitySel(i).Checked = True
            Select Case Index
                Case 0      ' E
                    SerialParity = "E"
                Case 1      ' M
                    SerialParity = "M"
                Case 2      ' N
                    SerialParity = "N"
                Case 3      ' O
                    SerialParity = "O"
                Case 4      ' S
                    SerialParity = "S"
            End Select
        Else
            mnuParitySel(i).Checked = False
        End If
    Next i
    
    If SerialOb.PortStatus(SerialPort).blnPortOpen = True Then
        lngStatus = SerialOb.CommSet(SerialPort, SerialString)
        If lngStatus <> 0 Then
            ' Handle error.
            lngStatus = SerialOb.CommGetError(strError)
            MsgBox "COM" & CStr(SerialPort) & " Error: " & strError
        End If
    End If
    
    UpdateStatus
    
End Sub

Private Sub mnuSpeedSel_Click(Index As Integer)
   Dim i As Integer
   Dim CurPortOpen As Boolean
   Dim lngStatus As Long
   Dim strError  As String
    
   For i = 0 To 12
        If (i = Index) Then
            mnuSpeedSel(i).Checked = True
            Select Case Index
                Case 0      ' 110
                    SerialBaud = "110"
                Case 1      ' 300
                    SerialBaud = "300"
                Case 2      ' 600
                    SerialBaud = "600"
                Case 3      ' 1200
                    SerialBaud = "1200"
                Case 4      ' 2400
                    SerialBaud = "2400"
                Case 5      ' 9600
                    SerialBaud = "9600"
                Case 6      ' 14400
                    SerialBaud = "14400"
                Case 7      ' 19200
                    SerialBaud = "19200"
                Case 8      ' 28800
                    SerialBaud = "28800"
                Case 9      ' 38400
                    SerialBaud = "38400"
                Case 10      ' 56000
                    SerialBaud = "56000"
                Case 11      ' 128000
                    SerialBaud = "128000"
                Case 12      ' 256000
                    SerialBaud = "256000"
            End Select
        Else
            mnuSpeedSel(i).Checked = False
        End If
    Next i
    
    If SerialOb.PortStatus(SerialPort).blnPortOpen = True Then
        lngStatus = SerialOb.CommSet(SerialPort, SerialString)
        If lngStatus <> 0 Then
            ' Handle error.
            lngStatus = SerialOb.CommGetError(strError)
            MsgBox "COM" & CStr(SerialPort) & " Error: " & strError
        End If
    End If
    
    UpdateStatus
    
End Sub

Private Sub mnuStopSel_Click(Index As Integer)
    Dim i As Integer
    Dim lngStatus As Long
    Dim strError  As String
    
    For i = 0 To 2
        If (i = Index) Then
            mnuStopSel(i).Checked = True
            Select Case Index
                Case 0      ' 1
                    SerialStop = "1"
                Case 1      ' 1.5
                    SerialStop = "1.5"
                Case 2      ' 2
                    SerialStop = "2"
            End Select
        Else
            mnuStopSel(i).Checked = False
        End If
    Next i
    
    If SerialOb.PortStatus(SerialPort).blnPortOpen = True Then
        lngStatus = SerialOb.CommSet(SerialPort, SerialString)
        If lngStatus <> 0 Then
            ' Handle error.
            lngStatus = SerialOb.CommGetError(strError)
            MsgBox "COM" & CStr(SerialPort) & " Error: " & strError
        End If
    End If
    UpdateStatus
End Sub

Private Sub UpdateStatus()

    If SerialOb.PortStatus(SerialPort).blnPortOpen = True Then
        StatusBar1.Panels(1).Text = "Connected"
        mnuConnect.Caption = "Dis&connect"
        btnSend(0).Enabled = True
        btnSend(1).Enabled = True
        btnCTSSend.Enabled = True
    Else
        StatusBar1.Panels(1).Text = "Disconnected"
        mnuConnect.Caption = "&Connect"
        btnSend(0).Enabled = False
        btnSend(1).Enabled = False
        btnCTSSend.Enabled = bDebug
    End If
    StatusBar1.Panels(2).Text = "COM" & SerialPort
    StatusBar1.Panels(3).Text = SerialString

End Sub
Private Function ValidatePort() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim intPortID As Integer ' Ex. 1, 2, 3, 4 for COM1 - COM4
    Dim lngStatus As Long
    Dim strError  As String
       
    On Error Resume Next
    ValidatePort = False
        
    For i = 1 To MAX_PORTS
       mnuCom(i).Checked = False
    Next i
    
    For i = MAX_PORTS To 1 Step -1
        ' Initialize Communications
        intPortID = i
        lngStatus = SerialOb.CommOpen(intPortID, "COM" & CStr(intPortID), SerialString)
    
        If lngStatus <> 0 Then
        ' Handle error.
            'lngStatus = SerialOb.CommGetError(strError)
            'MsgBox "COM" & CStr(i) & " Error: " & strError
            mnuCom(i).Enabled = False
        Else
            ValidatePort = True
            j = i
            SerialOb.CommClose intPortID
        End If
                
    Next i
        
    If j > 0 Then
        SerialPort = j
        mnuCom(j).Checked = True
    End If
    
End Function

Private Sub txtHEX_GotFocus(Index As Integer)
    txtHEX(Index).SelStart = 0
    txtHEX(Index).SelLength = Len(txtHEX(Index))
End Sub

Private Sub txtHEX_LostFocus(Index As Integer)
    Dim Temp As String
    
    Temp = txtHEX(Index)
    Select Case Len(txtHEX(Index))
        Case 0
            txtHEX(Index) = ""
        Case 1
            txtHEX(Index) = "0" & LegalHex(Mid(txtHEX(Index), 1, 1))
        Case Else
            txtHEX(Index) = LegalHex(Mid(txtHEX(Index), 1, 1)) & _
                            LegalHex(Mid(txtHEX(Index), 2, 1))
    End Select
    If (Len(txtHEX(Index)) > 0) Then
        Do While (Len(txtHEX(Index)) < 2)
            txtHEX(Index) = "0" & txtHEX(Index)
        Loop
    End If
End Sub

Private Function LegalHex(c As String) As String
    c = UCase(c)
    Select Case c
        Case "0" To "9", "A" To "F"
            LegalHex = c
        Case Else
            LegalHex = ""
    End Select
End Function

Private Sub txtTextOut_GotFocus()
    txtTextOut.SelStart = 0
    txtTextOut.SelLength = Len(txtTextOut)
End Sub

