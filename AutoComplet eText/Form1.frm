VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Complet Text Demo"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   465
      Left            =   1545
      TabIndex        =   1
      Top             =   1545
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   945
      Width           =   4515
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   3735
      TabIndex        =   4
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "A Sample of Auto Complete Text, Like IE address bar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   30
      TabIndex        =   3
      Top             =   15
      Width           =   4470
   End
   Begin VB.Label lblNote 
      Caption         =   "P.C. Work with IE 5.0 or above."
      Height          =   210
      Left            =   60
      TabIndex        =   2
      Top             =   2445
      Width           =   3105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DllVersionInfo
   cbSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformID As Long
End Type

Private Declare Function DllGetVersion _
   Lib "Shlwapi.dll" _
  (dwVersion As DllVersionInfo) As Long

Private Declare Function SHAutoComplete _
   Lib "Shlwapi.dll" _
  (ByVal hwndEdit As Long, _
   ByVal dwFlags As Long) As Long

Private Const SHACF_DEFAULT  As Long = &H0

Private Sub MakeAutoComplete(ByRef TextB As TextBox)
    Call SHAutoComplete(TextB.hWnd, SHACF_DEFAULT)
End Sub

Private Sub Form_Load()
  If IEVersion >= 5 Then
    Call MakeAutoComplete(Text1)
  Else
    MsgBox "You can't see the effect. Sorry!!" & vbCrLf & "You need to Install IE ver 5.0 or above"
  End If
End Sub

Public Function IEVersion() As Long

    Dim VersionInfo As DllVersionInfo
    VersionInfo.cbSize = Len(VersionInfo)
    
    Call DllGetVersion(VersionInfo)
    
    IEVersion = VersionInfo.dwMajorVersion

End Function

Public Function IEVersionString()

    Dim VersionInfo As DllVersionInfo
    VersionInfo.cbSize = Len(VersionInfo)
    
    Call DllGetVersion(VersionInfo)
    
    IEVersionString = "Internet Explorer " & _
        VersionInfo.dwMajorVersion & "." & _
        VersionInfo.dwMinorVersion & "." & _
        VersionInfo.dwBuildNumber

End Function

Private Sub Command2_Click()
    End
End Sub

Private Sub lblInfo_Click()
    MsgBox IEVersionString & vbCrLf & "Please feel free to write your Comments/Suggestions. Thnx!" & vbCrLf & "-Deepakk_2k@yahoo.com"
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo.FontUnderline = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo.FontUnderline = False
End Sub
