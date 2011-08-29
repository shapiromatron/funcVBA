VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "z_Tweet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long

'REQUIRED INPUTS (AND EXAMPLES IN COMMENTS)
Private pTweetDir As String          '"C:\Program Files\tweet\AJS"
Private pTweetEXE As String          '"tweet.exe"
Private pTweetTimer As String        'Now()
Private pTweetFrequency As String    '"00:30:00"

Private Sub Class_Initialize()
    pTweetTimer = Now()
End Sub

'INITIALIZE CLASS VARIABLES
Public Property Get TweetDir() As String
    TweetDir = pTweetDir
End Property
Public Property Let TweetDir(Value As String)
    pTweetDir = Value
End Property
Public Property Get TweetEXE() As String
    TweetEXE = pTweetEXE
End Property
Public Property Let TweetEXE(Value As String)
    pTweetEXE = Value
End Property
Public Property Get TweetFrequency() As String
    TweetFrequency = pTweetFrequency
End Property
Public Property Let TweetFrequency(Value As String)
    pTweetFrequency = Value
End Property

' Send a tweet
Public Function SendTweet(TweetString As String) As Boolean
    Dim ProcessID As Integer
    On Error GoTo IsError
    ChDir pTweetDir
    ProcessID = Shell(pTweetEXE & " " & AddQuotes(Left(TweetString, 140)), vbHide)
    SendTweet = True
    Exit Function
IsError:
    SendTweet = False
End Function

' SendTweetifTimer
Public Sub SendTweetAtTweetFreq(TweetString As String)
    If TimeValue(Now()) > TimeValue(pTweetTimer) + TimeValue(pTweetFrequency) Then
        SendTweet TweetString
        pTweetTimer = Now()
    End If
End Sub
