VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "never ogre special edition"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleMode       =   0  'User
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "II"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Text            =   "C:"
      Top             =   0
      Width           =   5055
   End
   Begin VB.FileListBox File1 
      Height          =   2970
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   4880
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8599
      _cy             =   1058
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function FolderExists(sFullPath As String) As Boolean
Dim myFSO As Object
Set myFSO = CreateObject("Scripting.FileSystemObject")
FolderExists = myFSO.FolderExists(sFullPath)
End Function
Private Function play()
Me.Caption = "Loading"
Dim name As String
Dim path As String
name = File1.FileName
path = (Text1.Text & "\" & name)
If Dir(path) <> "" Then
WindowsMediaPlayer1.URL = path
WindowsMediaPlayer1.Controls.play
Else
File1.Refresh
Me.Caption = "never ogre special edition"
End If
End Function

Private Sub Check1_Click()

End Sub

Private Sub Command1_Click()
play
End Sub

Private Sub Command3_Click()
WindowsMediaPlayer1.Controls.pause
End Sub

Private Sub Command2_Click()
WindowsMediaPlayer1.Controls.play
End Sub

Private Sub File1_DblClick()
play
End Sub


Private Sub File1_KeyPress(KeyAscii As Integer)
If (KeyAscii = vbKeyReturn) Then
    play
ElseIf (KeyAscii = vbKeySpace) Then
    If WindowsMediaPlayer1.playState = wmppsPlaying Then
    WindowsMediaPlayer1.Controls.pause
    ElseIf WindowsMediaPlayer1.playState = wmppsPaused Then
    WindowsMediaPlayer1.Controls.play
    End If
End If

End Sub

Private Sub Form_Load()
Command2.Visible = False
Command3.Visible = False

End Sub

Private Sub Text1_Change()

If FolderExists(Text1.Text) Then
File1.path = Text1.Text
End If
End Sub

Private Sub WindowsMediaPlayer1_PlayStateChange(ByVal NewState As Long)
Command3.Visible = False
Command2.Visible = False
Me.Caption = Replace(WindowsMediaPlayer1.URL, Text1.Text & "\", "")
If WindowsMediaPlayer1.playState = wmppsPaused Then
Command2.Visible = True
ElseIf WindowsMediaPlayer1.playState = wmppsPlaying Then
Command3.Visible = True
Else
Me.Caption = "never ogre special edition"
End If
End Sub
