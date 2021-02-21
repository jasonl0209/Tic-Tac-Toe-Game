'Author: Jason
'Date: January 8th, 2019
'Purpose: Applying all previously learned VB principles into a fun game
Option Explicit 
'GLOBAL VARIABLES
Dim intUser1 As Integer
Dim intUser2 As Integer
Dim intUser3 As Integer
Dim intComp1 As Integer
Dim intComp2 As Integer
Dim intComp3 As Integer
Dim intTurnCounter As Integer
Dim intUserScore As Integer
Dim intCompScore As Integer

Private Sub cmdBack1_Click()
picSettings.Visible = False
End Sub

Private Sub cmdBack2_Click()
picStartGame.Visible = True
End Sub

Private Sub cmdStartGame_Click()
picStartGame.Visible = False
ResetGame
End Sub

Private Sub cmdSettings_Click()
picSettings.Visible = True
End Sub

Private Sub Form_Activate()
picStartGame.Visible = True
Randomize
Initialize
imgUser1.Picture = LoadPicture(App.Path & "\GreenCircle.jpg")
imgUser2.Picture = LoadPicture(App.Path & "\GreenCircle.jpg")
imgUser3.Picture = LoadPicture(App.Path & "\GreenCircle.jpg")
imgComp1.Picture = LoadPicture(App.Path & "\RedCircle.jpg")
imgComp2.Picture = LoadPicture(App.Path & "\RedCircle.jpg")
imgComp3.Picture = LoadPicture(App.Path & "\RedCircle.jpg")
intUserScore = 0
intCompScore = 0
picUserScore.Print intUserScore
picCompScore.Print intCompScore
End Sub
Private Sub cmdReset_Click()
ResetGame
Initialize
intUserScore = 0
intCompScore = 0
picUserScore.Cls
picCompScore.Cls
picUserScore.Print intUserScore
picCompScore.Print intCompScore
End Sub

Private Sub cmdStartCompFirst_Click()
cmdStartCompFirst.Enabled = False
intComp1 = 5
imgComp1.Left = 208
imgComp1.Top = 200
intTurnCounter = intTurnCounter + 1
imgUser1.Enabled = True
cmdReset.Enabled = True
picCompScore.Print intCompScore
picUserScore.Print intUserScore
End Sub

Private Sub cmdStartUserFirst_Click()
imgUser1.Enabled = True
cmdReset.Enabled = True
picCompScore.Print intCompScore
picUserScore.Print intUserScore
End Sub

Private Sub optCompGoesFirst_Click()
cmdStartCompFirst.Enabled = True
cmdStartUserFirst.Visible = False
frmWhoGoesFirst.Enabled = False
End Sub

Private Sub optSoundOff_Click()
wmp1.URL = ""
End Sub

Private Sub optSoundOn_Click()
wmp1.URL = "D:\ieb.mp3"
End Sub

Private Sub optUserGoesFirst_Click()
cmdStartUserFirst.Enabled = True
cmdStartCompFirst.Visible = False
frmWhoGoesFirst.Enabled = False
End Sub

Private Sub picBackground_DragDrop(Source As Control, X As Single, Y As Single)

intTurnCounter = intTurnCounter + 1

Dim intFunction As Integer
intFunction = 0

If X > 0 And X < 180 Then
   If Y > 0 And Y < 180 Then
        Source.Left = 24
        Source.Top = 24
    ElseIf Y > 180 And Y < 360 Then
        Source.Left = 24
        Source.Top = 200
    Else
        Source.Left = 24
        Source.Top = 376
    End If
ElseIf X > 180 And X < 360 Then
    If Y > 0 And Y < 180 Then
        Source.Left = 208
        Source.Top = 24
    ElseIf Y > 180 And Y < 360 Then
        Source.Left = 208
        Source.Top = 200
    Else
        Source.Left = 208
        Source.Top = 376
    End If
ElseIf X > 360 And X < 540 Then
    If Y > 0 And Y < 180 Then
        Source.Left = 392
        Source.Top = 24
    ElseIf Y > 180 And Y < 360 Then
        Source.Left = 392
        Source.Top = 200
    Else
        Source.Left = 392
        Source.Top = 376
    End If
End If


If Source = imgUser1 Then
    intUser1 = ConvertXY(X, Y)
ElseIf Source = imgUser2 Then
    intUser2 = ConvertXY(X, Y)
ElseIf Source = imgUser3 Then
    intUser3 = ConvertXY(X, Y)
End If


'FIRST TURN CODE

If intTurnCounter = 1 Then
    If intUser1 = 5 Then
        intComp1 = 8
        imgComp1.Left = 24
        imgComp1.Top = 24
    Else
        intComp1 = 5
        imgComp1.Left = 208
        imgComp1.Top = 200
    End If

'SECOND TURN CODE

ElseIf intTurnCounter = 2 And optUserGoesFirst = True Then

'Basic Defensive Code
    
    If Conditions(15 - intUser1 - intUser2, intUser1, intComp1, intUser2, intUser3, intComp3) Then
        intFunction = MoveImage(15 - intUser1 - intUser2, intComp2, imgComp2)
        
    
'Special Conditional Double Trap Code
 
    ElseIf intUser1 = 5 And intUser2 = 2 Then
        intComp2 = 6
        imgComp2.Left = 392
        imgComp2.Top = 24

'Randomized, Non-Strategic Offensive Code

    Else
        Dim intRandom As Integer
        intRandom = 0
        Do
            intRandom = Int(9 * Rnd) + 1
        Loop While intRandom = intUser1 Or intRandom = intComp1 Or intRandom = intUser2
        intFunction = MoveImage(intRandom, intComp2, imgComp2)
    End If
    
ElseIf intTurnCounter = 2 And optCompGoesFirst = True Then

    If intUser1 = 8 Then
        intComp2 = 2
        imgComp2.Left = 392
        imgComp2.Top = 376
    ElseIf intUser1 = 2 Then
        intComp2 = 8
        imgComp2.Left = 24
        imgComp2.Top = 24
    ElseIf intUser1 = 6 Then
        intComp2 = 4
        imgComp2.Left = 24
        imgComp2.Top = 376
    ElseIf intUser1 = 4 Then
        intComp2 = 6
        imgComp2.Left = 392
        imgComp2.Top = 24
    Else
        intComp2 = 2
        imgComp2.Left = 392
        imgComp2.Top = 376
    End If
        
'TURN 3 CODE
        
ElseIf intTurnCounter = 3 Then

'Offensive Code for Turn 3

    If Conditions(15 - intComp1 - intComp2, intUser1, intComp1, intUser2, intComp2, intUser3) Then
        intFunction = MoveImage(15 - intComp1 - intComp2, intComp3, imgComp3)
                 
'Basic Defensive Code for Turn 3

    ElseIf Conditions(15 - intUser1 - intUser2, intComp1, intUser1, intComp2, intUser2, intComp3) Then
        intFunction = MoveImage(15 - intUser1 - intUser2, intComp3, imgComp3)

    ElseIf Conditions(15 - intUser2 - intUser3, intComp1, intUser2, intComp2, intUser3, intComp3) Then
        intFunction = MoveImage(15 - intUser2 - intUser3, intComp3, imgComp3)

    ElseIf Conditions(15 - intUser1 - intUser3, intUser1, intComp1, intComp2, intUser3, intComp3) Then
        intFunction = MoveImage(15 - intUser1 - intUser3, intComp3, imgComp3)
            
    'ElseIf optCompGoesFirst = True And
        
'Randomize Code for Turn 3

    Else
        Dim intRandom2 As Integer
        intRandom2 = 0
        Do
            intRandom2 = Int(9 * Rnd) + 1
        Loop While intRandom2 = intUser1 Or intRandom2 = intComp1 Or intRandom2 = intUser2 Or intRandom2 = intComp2 Or intRandom2 = intUser3
        intFunction = MoveImage(intRandom2, intComp3, imgComp3)
    End If

    
'TURN>3 CODE

Else

'Offensive Code for Turn>3

    If Conditions(15 - intComp1 - intComp2, intUser1, intComp1, intUser2, intComp2, intUser3) Then
        intFunction = MoveImage(15 - intComp1 - intComp2, intComp3, imgComp3)
        Picture1.Print "trump"
        Picture1.Print intUser1
        Picture1.Print intUser2
        Picture1.Print intUser3
        Picture1.Print intComp1
        Picture1.Print intComp2
        Picture1.Print intComp3

    ElseIf Conditions(15 - intComp1 - intComp3, intUser1, intComp1, intUser2, intUser3, intComp3) Then
        intFunction = MoveImage(15 - intComp1 - intComp3, intComp2, imgComp2)
        Picture1.Print "hillary"
        Picture1.Print intUser1
        Picture1.Print intUser2
        Picture1.Print intUser3
        Picture1.Print intComp1
        Picture1.Print intComp2
        Picture1.Print intComp3

    ElseIf Conditions(15 - intComp2 - intComp3, intUser1, intUser2, intComp2, intUser3, intComp3) Then
        intFunction = MoveImage(15 - intComp2 - intComp3, intComp1, imgComp1)
        Picture1.Print "Mueller"
        Picture1.Print intUser1
        Picture1.Print intUser2
        Picture1.Print intUser3
        Picture1.Print intComp1
        Picture1.Print intComp2
        Picture1.Print intComp3
        
        


    ElseIf intUser1 + intUser2 + intComp1 = 15 And intUser1 + intUser3 + intComp2 = 15 Then
        intFunction = EmbedFunction(imgComp3, intComp3)
        
ElseIf intUser1 + intUser2 + intComp1 = 15 And intUser1 + intUser3 + intComp3 = 15 Then
        intFunction = EmbedFunction(imgComp2, intComp2)
        
ElseIf intUser1 + intUser2 + intComp1 = 15 And intUser2 + intUser3 + intComp2 = 15 Then
        intFunction = EmbedFunction(imgComp3, intComp3)
        
ElseIf intUser1 + intUser2 + intComp1 = 15 And intUser2 + intUser3 + intComp3 = 15 Then
        intFunction = EmbedFunction(imgComp2, intComp2)
        

ElseIf intUser1 + intUser2 + intComp2 = 15 And intUser1 + intUser3 + intComp1 = 15 Then
        intFunction = EmbedFunction(imgComp3, intComp3)
        
ElseIf intUser1 + intUser2 + intComp2 = 15 And intUser1 + intUser3 + intComp3 = 15 Then
        intFunction = EmbedFunction(imgComp1, intComp1)
        
ElseIf intUser1 + intUser2 + intComp2 = 15 And intUser2 + intUser3 + intComp1 = 15 Then
        intFunction = EmbedFunction(imgComp3, intComp3)
        
ElseIf intUser1 + intUser2 + intComp2 = 15 And intUser2 + intUser3 + intComp3 = 15 Then
        intFunction = EmbedFunction(imgComp1, intComp1)
        

ElseIf intUser1 + intUser2 + intComp3 = 15 And intUser1 + intUser3 + intComp2 = 15 Then
        intFunction = EmbedFunction(imgComp1, intComp1)
        
ElseIf intUser1 + intUser2 + intComp3 = 15 And intUser1 + intUser3 + intComp1 = 15 Then
        intFunction = EmbedFunction(imgComp2, intComp2)
        
ElseIf intUser1 + intUser2 + intComp3 = 15 And intUser2 + intUser3 + intComp2 = 15 Then
        intFunction = EmbedFunction(imgComp1, intComp1)
        
ElseIf intUser1 + intUser2 + intComp3 = 15 And intUser2 + intUser3 + intComp1 = 15 Then
        intFunction = EmbedFunction(imgComp2, intComp2)
        

ElseIf intUser1 + intUser3 + intComp1 = 15 And intUser2 + intUser3 + intComp2 = 15 Then
        intFunction = EmbedFunction(imgComp3, intComp3)
        
ElseIf intUser1 + intUser3 + intComp1 = 15 And intUser2 + intUser3 + intComp3 = 15 Then
        intFunction = EmbedFunction(imgComp2, intComp2)
        

 ElseIf intUser1 + intUser3 + intComp2 = 15 And intUser2 + intUser3 + intComp1 = 15 Then
        intFunction = EmbedFunction(imgComp3, intComp3)
        
ElseIf intUser1 + intUser3 + intComp2 = 15 And intUser2 + intUser3 + intComp3 = 15 Then
        intFunction = EmbedFunction(imgComp1, intComp1)
        

ElseIf intUser1 + intUser3 + intComp3 = 15 And intUser2 + intUser3 + intComp1 = 15 Then
        intFunction = EmbedFunction(imgComp2, intComp2)
        
ElseIf intUser1 + intUser3 + intComp3 = 15 And intUser2 + intUser3 + intComp2 = 15 Then
        intFunction = EmbedFunction(imgComp1, intComp1)
        
        
ElseIf intUser1 + intUser2 + intComp1 = 15 Then
        intFunction = EmbedFunction(imgComp3, intComp3)
        
ElseIf intUser1 + intUser2 + intComp2 = 15 Then
        intFunction = EmbedFunction(imgComp1, intComp1)
        
ElseIf intUser1 + intUser2 + intComp3 = 15 Then
        intFunction = EmbedFunction(imgComp2, intComp2)
        
        
ElseIf intUser1 + intUser3 + intComp1 = 15 Then
        intFunction = EmbedFunction(imgComp3, intComp3)
        
ElseIf intUser1 + intUser3 + intComp2 = 15 Then
        intFunction = EmbedFunction(imgComp1, intComp1)
        
ElseIf intUser1 + intUser3 + intComp3 = 15 Then
        intFunction = EmbedFunction(imgComp2, intComp2)
        
        
ElseIf intUser2 + intUser3 + intComp1 = 15 Then
        intFunction = EmbedFunction(imgComp3, intComp3)
        
ElseIf intUser2 + intUser3 + intComp2 = 15 Then
        intFunction = EmbedFunction(imgComp1, intComp1)
        
ElseIf intUser2 + intUser3 + intComp3 = 15 Then
        intFunction = EmbedFunction(imgComp2, intComp2)
        

        


'Basic Defensive Code for Turn>3

    ElseIf Conditions(15 - intUser1 - intUser2, intUser1, intComp1, intUser2, intComp2, intComp3) Then
        intFunction = MoveImage(15 - intUser1 - intUser2, intComp3, imgComp3)
        Picture1.Print "defensive 1"

    ElseIf Conditions(15 - intUser2 - intUser3, intComp1, intUser2, intComp2, intUser3, intComp3) Then
        intFunction = MoveImage(15 - intUser2 - intUser3, intComp3, imgComp3)
        Picture1.Print "defensive 2"

    ElseIf Conditions(15 - intUser1 - intUser3, intUser1, intComp1, intComp2, intUser3, intComp3) Then
        intFunction = MoveImage(15 - intUser1 - intUser3, intComp3, imgComp3)
        Picture1.Print "defensive 3"
        Picture1.Print intComp3
        
        
'Randomize Code for TC > 3
    
    Else
       Dim intRandom3 As Integer
       intRandom3 = 0
       Do
           intRandom3 = Int(9 * Rnd) + 1
       Loop While intRandom3 = intUser1 Or intRandom3 = intComp1 Or intRandom3 = intUser2 Or intRandom3 = intComp2 Or intRandom3 = intUser3 Or intRandom3 = intComp3
       intFunction = MoveImage(intRandom3, intComp3, imgComp3)
       Picture1.Print "random"
    End If
End If

If optUserGoesFirst = True Then
    If intTurnCounter = 1 Then
        imgUser1.Enabled = False
        imgUser2.Enabled = True
    ElseIf intTurnCounter = 2 Then
        imgUser2.Enabled = False
        imgUser3.Enabled = True
    Else
        imgUser1.Enabled = True
        imgUser2.Enabled = True
    End If
End If

If optCompGoesFirst = True Then
    If intTurnCounter = 2 Then
        imgUser2.Enabled = True
    ElseIf intTurnCounter = 3 Then
        imgUser3.Enabled = True
    Else
        imgUser1.Enabled = True
        imgUser2.Enabled = True
    End If
End If

If intUser1 + intUser2 + intUser3 = 15 Then
    intUserScore = intUserScore + 1
    picUserScore.Cls
    picUserScore.Print intUserScore
    MsgBox ("Congraluations! You won!")
    ResetGame
    Initialize
ElseIf intComp1 + intComp2 + intComp3 = 15 Then
    intCompScore = intCompScore + 1
    picCompScore.Cls
    picCompScore.Print intCompScore
    MsgBox ("Unfortunately, the computer won this time. Better Luck Next Time!")
    ResetGame
    Initialize
End If
End Sub

Public Function ConvertXY(X As Single, Y As Single)
Dim intMagicNumber As Integer
intMagicNumber = 0
If X > 0 And X < 180 Then
    If Y > 0 And Y < 180 Then
       intMagicNumber = 8
    ElseIf Y > 180 And Y < 360 Then
        intMagicNumber = 3
    ElseIf Y > 360 And Y < 540 Then
        intMagicNumber = 4
    End If
ElseIf X > 180 And X < 360 Then
    If Y > 0 And Y < 180 Then
        intMagicNumber = 1
    ElseIf Y > 180 And Y < 360 Then
        intMagicNumber = 5
    ElseIf Y > 360 And Y < 540 Then
        intMagicNumber = 9
    End If
ElseIf X > 360 And X < 540 Then
    If Y > 0 And Y < 180 Then
        intMagicNumber = 6
    ElseIf Y > 180 And Y < 360 Then
        intMagicNumber = 7
    ElseIf Y > 360 And Y < 540 Then
        intMagicNumber = 2
    End If
End If
ConvertXY = intMagicNumber
End Function

Public Function MoveImage(X As Integer, intX As Integer, Image As Image)
If X = 8 Then
    intX = 8
    Image.Left = 24
    Image.Top = 24
ElseIf X = 3 Then
    intX = 3
    Image.Left = 24
    Image.Top = 200
ElseIf X = 4 Then
    intX = 4
    Image.Left = 24
    Image.Top = 376
ElseIf X = 1 Then
    intX = 1
    Image.Left = 208
    Image.Top = 24
ElseIf X = 5 Then
    intX = 5
    Image.Left = 208
    Image.Top = 200
ElseIf X = 9 Then
    intX = 9
    Image.Left = 208
    Image.Top = 376
ElseIf X = 6 Then
    intX = 6
    Image.Left = 392
    Image.Top = 24
ElseIf X = 7 Then
    intX = 7
    Image.Left = 392
    Image.Top = 200
Else
    intX = 2
    Image.Left = 392
    Image.Top = 376
End If
End Function

Public Sub ResetGame()
imgUser1.Left = 616
imgUser1.Top = 56
imgUser2.Left = 616
imgUser2.Top = 224
imgUser3.Left = 616
imgUser3.Top = 392
imgComp1.Left = 808
imgComp1.Top = 56
imgComp2.Left = 808
imgComp2.Top = 224
imgComp3.Left = 808
imgComp3.Top = 392
imgUser1.Enabled = False
imgUser2.Enabled = False
imgUser3.Enabled = False
frmWhoGoesFirst.Enabled = True
cmdStartCompFirst.Enabled = False
cmdStartCompFirst.Visible = True
cmdStartUserFirst.Enabled = False
cmdStartUserFirst.Visible = True
optUserGoesFirst.Value = False
optCompGoesFirst.Value = False
End Sub

Public Sub Initialize()
intUser1 = -91
intUser2 = -92
intUser3 = -93
intComp1 = -94
intComp2 = -95
intComp3 = -96
intTurnCounter = 0
End Sub

Public Function Conditions(A As Integer, B As Integer, C As Integer, D As Integer, E As Integer, F As Integer)
Conditions = A >= 1 And A <= 9 And A <> B And A <> C And A <> D And A <> E And A <> F
End Function

'Sub PlayBackgroundSoundFile()
'If optSoundOn = True Then
    'My.Computer.Audio.Play("H:\ABC.mp3", AudioPlayMode.WaitToComplete)
'End If
'End Sub

Private Sub wmp1_OpenStateChange(ByVal NewState As Long)

End Sub


Public Function GenerateRandom(imgCompX As Image, intCompX As Integer)
'Input imgComp1, imgComp2, or imgComp3 into imgCompX and intComp1, intComp2, or intComp3 into intCompX, X must correspond

Dim intRandomX As Integer
intRandomX = 0
Do
    intRandomX = Int(9 * Rnd) + 1
Loop While intRandomX = intUser1 Or intRandomX = intComp1 Or intRandomX = intUser2 Or intRandomX = intComp2 Or intRandomX = intUser3 Or intRandomX = intComp3

GenerateRandom = MoveImage(intRandomX, intCompX, imgCompX)

End Function

Public Function EmbedFunction(imgCompA As Image, intCompA As Integer)
Dim intFunction2 As Integer
intFunction2 = 0
If Conditions(15 - intUser1 - intUser2, intUser1, intComp1, intUser2, intComp2, intComp3) Then
    intFunction2 = MoveImage(15 - intUser1 - intUser2, intCompA, imgCompA)

ElseIf Conditions(15 - intUser2 - intUser3, intComp1, intUser2, intComp2, intUser3, intComp3) Then
    intFunction2 = MoveImage(15 - intUser2 - intUser3, intCompA, imgCompA)

ElseIf Conditions(15 - intUser1 - intUser3, intUser1, intComp1, intComp2, intUser3, intComp3) Then
    intFunction2 = MoveImage(15 - intUser1 - intUser3, intCompA, imgCompA)

Else
    intFunction2 = GenerateRandom(imgCompA, intCompA)
End If

EmbedFunction = intFunction2


End Function


