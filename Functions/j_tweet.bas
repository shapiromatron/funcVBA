Option Explicit
Global TweetFeed As Tweet

Private Sub MainText()
    
    'setup in beginning
    If Range("TweetOn") = True Then
        Set TweetFeed = New Tweet
        SetupTweet
    End If
    
    'error messaging
    On Error GoTo SendTweet
    
    For i = 1 To 100
        
        'call successive looping tweet
        If Range("TweetOn") = True Then
            TweetFeed.SendTweetAtTweetFreq "Model running succesfully at time " & Now()
        End If
        
    Next i
    
    'send completion tweet
    TweetFeed.SendTweet "Model complete!"
    
    Exit Sub
SentTweet:
    TweetFeed.Tweet "Error!"
End Sub

Private Sub SetupTweet()
    Set TweetFeed = New Tweet
    TweetFeed.TweetDir = MakeDirString(Range("TweetDir"))     ' "C:\Program Files\tweet"
    TweetFeed.TweetEXE = Range("TweetEXE")                          ' "tweet.exe"
    TweetFeed.TweetFrequency = Range("TweetFrequency")              ' "00:00:15"
End Sub

