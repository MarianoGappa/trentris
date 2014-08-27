Attribute VB_Name = "Module2"
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Type rank
    nombre As String * 15
    puntaje As Long
End Type
Public Const MULTX = 25, MULTY = 25, STARTX = 0, STARTY = 0
Public map(15, 21) As Integer
Public objs(6, 4, 4, 4) As Integer
Public obj(4, 4, 4) As Integer
Public ranks(5) As rank, higherlevel As Integer
Public levelup As Boolean
Public soundEnabled As Boolean
Public objx As Integer
Public objy As Integer
Public objp As Integer
Public objc As Integer
Public num As Integer
Public nextb As Integer
Public nextc As Integer
Public nextp As Integer
Public auxchoto As Boolean
Public Const SND_ASYNC = &H1

Sub Main()
    ChDir App.Path
    Randomize Timer
    InitObjs
    LoadRankings
    soundEnabled = True
    frmMain.Show
    setPremios (higherlevel)
End Sub
Public Sub PlaySoundX(filename As String)

' If sound is enabled and filename exists,
' play the specified sound.

filename = App.Path & "\" & filename

If HZW_FileExists(filename) And soundEnabled Then
    PlaySound filename, CLng(0), SND_ASYNC
End If

End Sub
Function HZW_FileExists(ByVal WhatFile As String) As Boolean
    HZW_FileExists = (0 < Len(Trim$(Dir$(WhatFile))))
End Function
Sub LoadRankings()
    If HZW_FileExists("ranks.dat") Then
        Dim i As Integer
        Open "ranks.dat" For Input As #1
        Input #1, higherlevel
        For i = 1 To 5
            Input #1, ranks(i).nombre
            Input #1, ranks(i).puntaje
        Next
        Close #1
    Else
            higherlevel = 1
            ranks(1).nombre = "JAMES BLUNT"
            ranks(1).puntaje = 5000
            ranks(2).nombre = "NENE DE AIRBAG"
            ranks(2).puntaje = 4000
            ranks(3).nombre = "PATO DE TEXITEZ"
            ranks(3).puntaje = 3000
            ranks(4).nombre = "SHAKIRA"
            ranks(4).puntaje = 2000
            ranks(5).nombre = "ROBBIE WILLIAMS"
            ranks(5).puntaje = 1000
    End If
End Sub
Sub SaveRankings()
        Dim i As Integer
        Open "ranks.dat" For Output As #1
        Write #1, higherlevel
        For i = 1 To 5
            Write #1, ranks(i).nombre
            Write #1, ranks(i).puntaje
        Next
        Close #1
End Sub
Sub ModifyRankings(nombre As String, puntaje As Long)
    Dim i As Integer, p As Integer
    i = 6
    While i > 1 And puntaje > ranks(i - 1).puntaje
        i = i - 1
    Wend
    p = i
    If i < 6 Then
        While i > p
            ranks(i) = ranks(i - 1)
            i = i - 1
        Wend
        ranks(p).nombre = nombre
        ranks(p).puntaje = puntaje
    End If
    SaveRankings
End Sub
Function levelSpeed(num As Integer)
    Select Case num
    Case 1
        levelSpeed = 800
    Case 2
        levelSpeed = 700
    Case 3
        levelSpeed = 600
    Case 4
        levelSpeed = 500
    Case 5
        levelSpeed = 400
    Case 6
        levelSpeed = 300
    Case 7
        levelSpeed = 200
    Case 8
        levelSpeed = 100
    Case 9
        levelSpeed = 50
    Case Else
        levelSpeed = 25
    End Select
End Function
Sub startUp()
    levelup = False
    lastlevel = higherlevel
    frmMain.Level.Caption = 1
    frmMain.score.Caption = Str(0)
    frmMain.tiempo.Interval = levelSpeed(frmMain.Level.Caption)
    auxchoto = False
    nextc = Int(Rnd * 15) + 1
    nextp = Int(Rnd * 4) + 1
    nextb = Int(Rnd * 6) + 1
    chooseObj
    For yi = 1 To 20
        For xi = 1 To 15
            map(xi, yi) = 0
        Next
    Next
    For i = 1 To 15
        map(i, 21) = 1
    Next
    refreshMap
    frmMain.tiempo.Enabled = True
End Sub
Sub setPremios(num As Integer)
    If num >= 2 Then frmMain.mnuP1.Enabled = True
    If num >= 3 Then frmMain.mnuP2.Enabled = True
    If num >= 4 Then frmMain.mnuP3.Enabled = True
    If num >= 5 Then frmMain.mnuP4.Enabled = True
    If num >= 6 Then frmMain.mnuP5.Enabled = True
    If num >= 7 Then frmMain.mnuP6.Enabled = True
    If num >= 8 Then frmMain.mnuP7.Enabled = True
    If num >= 9 Then frmMain.mnuP8.Enabled = True
    If num >= 10 Then frmMain.mnuP9.Enabled = True
End Sub
Sub chooseObj()
        Dim pi As Integer, xi As Integer, yi As Integer
        objy = 1
        objx = 8
        num = nextb
        objc = nextc
        objp = nextp
        nextc = Int(Rnd * 15) + 1
        nextp = Int(Rnd * 4) + 1
        nextb = Int(Rnd * 6) + 1
        showNext
        For pi = 1 To 4
                For xi = 1 To 4
                        For yi = 1 To 4
                                obj(pi, xi, yi) = objs(num, pi, xi, yi)
                        Next
                Next
        Next
End Sub

Sub killLine(y As Integer)
    If y > 1 Then
        For i = 0 To y - 2
            For x = 1 To 15
                map(x, y - i) = map(x, y - i - 1)
            Next
        Next
    End If
    For x = 1 To 15
        map(x, 1) = 0
    Next
End Sub
Sub checkLines()
    Dim xi As Integer, yi As Integer, linesCount As Integer, isLine As Boolean
    Dim lastlevel As Integer

    linesCount = 0
    For yi = 1 To 20
        isLine = True
        For xi = 1 To 15
            If map(xi, yi) = 0 Then isLine = False
        Next
        If isLine Then
            linesCount = linesCount + 1
            killLine yi
        End If
    Next
    Select Case linesCount
    Case 1
        frmMain.score.Caption = frmMain.score.Caption + 10
        PlaySoundX "1line.wav"
    Case 2
        frmMain.score.Caption = frmMain.score.Caption + 25
        PlaySoundX "2line.wav"
    Case 3
        frmMain.score.Caption = frmMain.score.Caption + 50
        PlaySoundX "3line.wav"
    Case 4
        frmMain.score.Caption = frmMain.score.Caption + 100
        PlaySoundX "4line.wav"
    End Select
    
    lastlevel = Val(frmMain.Level.Caption)
    While Val(frmMain.score.Caption) >= Val(frmMain.Level.Caption) * 300
        frmMain.Level.Caption = frmMain.Level.Caption + 1
    Wend
    
    ' Actualizar máximo nivel
    If levelup = False And Val(frmMain.Level.Caption) > higherlevel Then levelup = True
    
    If linesCount > 0 Then
        refreshMap
    End If
End Sub

Sub InitObjs()

' Block
objs(1, 1, 1, 1) = 1: objs(1, 1, 2, 1) = 1: objs(1, 1, 3, 1) = 0: objs(1, 1, 4, 1) = 0
objs(1, 1, 1, 2) = 1: objs(1, 1, 2, 2) = 1: objs(1, 1, 3, 2) = 0: objs(1, 1, 4, 2) = 0
objs(1, 1, 1, 3) = 0: objs(1, 1, 2, 3) = 0: objs(1, 1, 3, 3) = 0: objs(1, 1, 4, 3) = 0
objs(1, 1, 1, 4) = 0: objs(1, 1, 2, 4) = 0: objs(1, 1, 3, 4) = 0: objs(1, 1, 4, 4) = 0

objs(1, 2, 1, 1) = 1: objs(1, 2, 2, 1) = 1: objs(1, 2, 3, 1) = 0: objs(1, 2, 4, 1) = 0
objs(1, 2, 1, 2) = 1: objs(1, 2, 2, 2) = 1: objs(1, 2, 3, 2) = 0: objs(1, 2, 4, 2) = 0
objs(1, 2, 1, 3) = 0: objs(1, 2, 2, 3) = 0: objs(1, 2, 3, 3) = 0: objs(1, 2, 4, 3) = 0
objs(1, 2, 1, 4) = 0: objs(1, 2, 2, 4) = 0: objs(1, 2, 3, 4) = 0: objs(1, 2, 4, 4) = 0

objs(1, 3, 1, 1) = 1: objs(1, 3, 2, 1) = 1: objs(1, 3, 3, 1) = 0: objs(1, 3, 4, 1) = 0
objs(1, 3, 1, 2) = 1: objs(1, 3, 2, 2) = 1: objs(1, 3, 3, 2) = 0: objs(1, 3, 4, 2) = 0
objs(1, 3, 1, 3) = 0: objs(1, 3, 2, 3) = 0: objs(1, 3, 3, 3) = 0: objs(1, 3, 4, 3) = 0
objs(1, 3, 1, 4) = 0: objs(1, 3, 2, 4) = 0: objs(1, 3, 3, 4) = 0: objs(1, 3, 4, 4) = 0

objs(1, 4, 1, 1) = 1: objs(1, 4, 2, 1) = 1: objs(1, 4, 3, 1) = 0: objs(1, 4, 4, 1) = 0
objs(1, 4, 1, 2) = 1: objs(1, 4, 2, 2) = 1: objs(1, 4, 3, 2) = 0: objs(1, 4, 4, 2) = 0
objs(1, 4, 1, 3) = 0: objs(1, 4, 2, 3) = 0: objs(1, 4, 3, 3) = 0: objs(1, 4, 4, 3) = 0
objs(1, 4, 1, 4) = 0: objs(1, 4, 2, 4) = 0: objs(1, 4, 3, 4) = 0: objs(1, 4, 4, 4) = 0
                                                                         

'Straight
objs(2, 1, 1, 1) = 1: objs(2, 1, 2, 1) = 1: objs(2, 1, 3, 1) = 1: objs(2, 1, 4, 1) = 1
objs(2, 1, 1, 2) = 0: objs(2, 1, 2, 2) = 0: objs(2, 1, 3, 2) = 0: objs(2, 1, 4, 2) = 0
objs(2, 1, 1, 3) = 0: objs(2, 1, 2, 3) = 0: objs(2, 1, 3, 3) = 0: objs(2, 1, 4, 3) = 0
objs(2, 1, 1, 4) = 0: objs(2, 1, 2, 4) = 0: objs(2, 1, 3, 4) = 0: objs(2, 1, 4, 4) = 0

objs(2, 2, 1, 1) = 0: objs(2, 2, 2, 1) = 1: objs(2, 2, 3, 1) = 0: objs(2, 2, 4, 1) = 0
objs(2, 2, 1, 2) = 0: objs(2, 2, 2, 2) = 1: objs(2, 2, 3, 2) = 0: objs(2, 2, 4, 2) = 0
objs(2, 2, 1, 3) = 0: objs(2, 2, 2, 3) = 1: objs(2, 2, 3, 3) = 0: objs(2, 2, 4, 3) = 0
objs(2, 2, 1, 4) = 0: objs(2, 2, 2, 4) = 1: objs(2, 2, 3, 4) = 0: objs(2, 2, 4, 4) = 0

objs(2, 3, 1, 1) = 1: objs(2, 3, 2, 1) = 1: objs(2, 3, 3, 1) = 1: objs(2, 3, 4, 1) = 1
objs(2, 3, 1, 2) = 0: objs(2, 3, 2, 2) = 0: objs(2, 3, 3, 2) = 0: objs(2, 3, 4, 2) = 0
objs(2, 3, 1, 3) = 0: objs(2, 3, 2, 3) = 0: objs(2, 3, 3, 3) = 0: objs(2, 3, 4, 3) = 0
objs(2, 3, 1, 4) = 0: objs(2, 3, 2, 4) = 0: objs(2, 3, 3, 4) = 0: objs(2, 3, 4, 4) = 0

objs(2, 4, 1, 1) = 0: objs(2, 4, 2, 1) = 1: objs(2, 4, 3, 1) = 0: objs(2, 4, 4, 1) = 0
objs(2, 4, 1, 2) = 0: objs(2, 4, 2, 2) = 1: objs(2, 4, 3, 2) = 0: objs(2, 4, 4, 2) = 0
objs(2, 4, 1, 3) = 0: objs(2, 4, 2, 3) = 1: objs(2, 4, 3, 3) = 0: objs(2, 4, 4, 3) = 0
objs(2, 4, 1, 4) = 0: objs(2, 4, 2, 4) = 1: objs(2, 4, 3, 4) = 0: objs(2, 4, 4, 4) = 0

                           
'Worse
objs(3, 1, 1, 1) = 1: objs(3, 1, 2, 1) = 1: objs(3, 1, 3, 1) = 0: objs(3, 1, 4, 1) = 0
objs(3, 1, 1, 2) = 0: objs(3, 1, 2, 2) = 1: objs(3, 1, 3, 2) = 1: objs(3, 1, 4, 2) = 0
objs(3, 1, 1, 3) = 0: objs(3, 1, 2, 3) = 0: objs(3, 1, 3, 3) = 0: objs(3, 1, 4, 3) = 0
objs(3, 1, 1, 4) = 0: objs(3, 1, 2, 4) = 0: objs(3, 1, 3, 4) = 0: objs(3, 1, 4, 4) = 0

objs(3, 2, 1, 1) = 0: objs(3, 2, 2, 1) = 1: objs(3, 2, 3, 1) = 0: objs(3, 2, 4, 1) = 0
objs(3, 2, 1, 2) = 1: objs(3, 2, 2, 2) = 1: objs(3, 2, 3, 2) = 0: objs(3, 2, 4, 2) = 0
objs(3, 2, 1, 3) = 1: objs(3, 2, 2, 3) = 0: objs(3, 2, 3, 3) = 0: objs(3, 2, 4, 3) = 0
objs(3, 2, 1, 4) = 0: objs(3, 2, 2, 4) = 0: objs(3, 2, 3, 4) = 0: objs(3, 2, 4, 4) = 0

objs(3, 3, 1, 1) = 1: objs(3, 3, 2, 1) = 1: objs(3, 3, 3, 1) = 0: objs(3, 3, 4, 1) = 0
objs(3, 3, 1, 2) = 0: objs(3, 3, 2, 2) = 1: objs(3, 3, 3, 2) = 1: objs(3, 3, 4, 2) = 0
objs(3, 3, 1, 3) = 0: objs(3, 3, 2, 3) = 0: objs(3, 3, 3, 3) = 0: objs(3, 3, 4, 3) = 0
objs(3, 3, 1, 4) = 0: objs(3, 3, 2, 4) = 0: objs(3, 3, 3, 4) = 0: objs(3, 3, 4, 4) = 0

objs(3, 4, 1, 1) = 0: objs(3, 4, 2, 1) = 1: objs(3, 4, 3, 1) = 0: objs(3, 4, 4, 1) = 0
objs(3, 4, 1, 2) = 1: objs(3, 4, 2, 2) = 1: objs(3, 4, 3, 2) = 0: objs(3, 4, 4, 2) = 0
objs(3, 4, 1, 3) = 1: objs(3, 4, 2, 3) = 0: objs(3, 4, 3, 3) = 0: objs(3, 4, 4, 3) = 0
objs(3, 4, 1, 4) = 0: objs(3, 4, 2, 4) = 0: objs(3, 4, 3, 4) = 0: objs(3, 4, 4, 4) = 0


'Boat
objs(4, 1, 1, 1) = 1: objs(4, 1, 2, 1) = 1: objs(4, 1, 3, 1) = 1: objs(4, 1, 4, 1) = 0
objs(4, 1, 1, 2) = 0: objs(4, 1, 2, 2) = 1: objs(4, 1, 3, 2) = 0: objs(4, 1, 4, 2) = 0
objs(4, 1, 1, 3) = 0: objs(4, 1, 2, 3) = 0: objs(4, 1, 3, 3) = 0: objs(4, 1, 4, 3) = 0
objs(4, 1, 1, 4) = 0: objs(4, 1, 2, 4) = 0: objs(4, 1, 3, 4) = 0: objs(4, 1, 4, 4) = 0
                                                  
objs(4, 2, 1, 1) = 0: objs(4, 2, 2, 1) = 1: objs(4, 2, 3, 1) = 0: objs(4, 2, 4, 1) = 0
objs(4, 2, 1, 2) = 0: objs(4, 2, 2, 2) = 1: objs(4, 2, 3, 2) = 1: objs(4, 2, 4, 2) = 0
objs(4, 2, 1, 3) = 0: objs(4, 2, 2, 3) = 1: objs(4, 2, 3, 3) = 0: objs(4, 2, 4, 3) = 0
objs(4, 2, 1, 4) = 0: objs(4, 2, 2, 4) = 0: objs(4, 2, 3, 4) = 0: objs(4, 2, 4, 4) = 0
 
objs(4, 3, 1, 1) = 0: objs(4, 3, 2, 1) = 1: objs(4, 3, 3, 1) = 0: objs(4, 3, 4, 1) = 0
objs(4, 3, 1, 2) = 1: objs(4, 3, 2, 2) = 1: objs(4, 3, 3, 2) = 1: objs(4, 3, 4, 2) = 0
objs(4, 3, 1, 3) = 0: objs(4, 3, 2, 3) = 0: objs(4, 3, 3, 3) = 0: objs(4, 3, 4, 3) = 0
objs(4, 3, 1, 4) = 0: objs(4, 3, 2, 4) = 0: objs(4, 3, 3, 4) = 0: objs(4, 3, 4, 4) = 0
 
objs(4, 4, 1, 1) = 0: objs(4, 4, 2, 1) = 0: objs(4, 4, 3, 1) = 1: objs(4, 4, 4, 1) = 0
objs(4, 4, 1, 2) = 0: objs(4, 4, 2, 2) = 1: objs(4, 4, 3, 2) = 1: objs(4, 4, 4, 2) = 0
objs(4, 4, 1, 3) = 0: objs(4, 4, 2, 3) = 0: objs(4, 4, 3, 3) = 1: objs(4, 4, 4, 3) = 0
objs(4, 4, 1, 4) = 0: objs(4, 4, 2, 4) = 0: objs(4, 4, 3, 4) = 0: objs(4, 4, 4, 4) = 0


'Hook
objs(5, 1, 1, 1) = 1: objs(5, 1, 2, 1) = 1: objs(5, 1, 3, 1) = 1: objs(5, 1, 4, 1) = 0
objs(5, 1, 1, 2) = 0: objs(5, 1, 2, 2) = 0: objs(5, 1, 3, 2) = 1: objs(5, 1, 4, 2) = 0
objs(5, 1, 1, 3) = 0: objs(5, 1, 2, 3) = 0: objs(5, 1, 3, 3) = 0: objs(5, 1, 4, 3) = 0
objs(5, 1, 1, 4) = 0: objs(5, 1, 2, 4) = 0: objs(5, 1, 3, 4) = 0: objs(5, 1, 4, 4) = 0

objs(5, 2, 1, 1) = 1: objs(5, 2, 2, 1) = 1: objs(5, 2, 3, 1) = 0: objs(5, 2, 4, 1) = 0
objs(5, 2, 1, 2) = 1: objs(5, 2, 2, 2) = 0: objs(5, 2, 3, 2) = 0: objs(5, 2, 4, 2) = 0
objs(5, 2, 1, 3) = 1: objs(5, 2, 2, 3) = 0: objs(5, 2, 3, 3) = 0: objs(5, 2, 4, 3) = 0
objs(5, 2, 1, 4) = 0: objs(5, 2, 2, 4) = 0: objs(5, 2, 3, 4) = 0: objs(5, 2, 4, 4) = 0

objs(5, 3, 1, 1) = 1: objs(5, 3, 2, 1) = 0: objs(5, 3, 3, 1) = 0: objs(5, 3, 4, 1) = 0
objs(5, 3, 1, 2) = 1: objs(5, 3, 2, 2) = 1: objs(5, 3, 3, 2) = 1: objs(5, 3, 4, 2) = 0
objs(5, 3, 1, 3) = 0: objs(5, 3, 2, 3) = 0: objs(5, 3, 3, 3) = 0: objs(5, 3, 4, 3) = 0
objs(5, 3, 1, 4) = 0: objs(5, 3, 2, 4) = 0: objs(5, 3, 3, 4) = 0: objs(5, 3, 4, 4) = 0

objs(5, 4, 1, 1) = 0: objs(5, 4, 2, 1) = 1: objs(5, 4, 3, 1) = 0: objs(5, 4, 4, 1) = 0
objs(5, 4, 1, 2) = 0: objs(5, 4, 2, 2) = 1: objs(5, 4, 3, 2) = 0: objs(5, 4, 4, 2) = 0
objs(5, 4, 1, 3) = 1: objs(5, 4, 2, 3) = 1: objs(5, 4, 3, 3) = 0: objs(5, 4, 4, 3) = 0
objs(5, 4, 1, 4) = 0: objs(5, 4, 2, 4) = 0: objs(5, 4, 3, 4) = 0: objs(5, 4, 4, 4) = 0

'Hook 2
objs(6, 1, 1, 1) = 0: objs(6, 1, 2, 1) = 0: objs(6, 1, 3, 1) = 1: objs(6, 1, 4, 1) = 0
objs(6, 1, 1, 2) = 1: objs(6, 1, 2, 2) = 1: objs(6, 1, 3, 2) = 1: objs(6, 1, 4, 2) = 0
objs(6, 1, 1, 3) = 0: objs(6, 1, 2, 3) = 0: objs(6, 1, 3, 3) = 0: objs(6, 1, 4, 3) = 0
objs(6, 1, 1, 4) = 0: objs(6, 1, 2, 4) = 0: objs(6, 1, 3, 4) = 0: objs(6, 1, 4, 4) = 0

objs(6, 2, 1, 1) = 1: objs(6, 2, 2, 1) = 1: objs(6, 2, 3, 1) = 0: objs(6, 2, 4, 1) = 0
objs(6, 2, 1, 2) = 0: objs(6, 2, 2, 2) = 1: objs(6, 2, 3, 2) = 0: objs(6, 2, 4, 2) = 0
objs(6, 2, 1, 3) = 0: objs(6, 2, 2, 3) = 1: objs(6, 2, 3, 3) = 0: objs(6, 2, 4, 3) = 0
objs(6, 2, 1, 4) = 0: objs(6, 2, 2, 4) = 0: objs(6, 2, 3, 4) = 0: objs(6, 2, 4, 4) = 0

objs(6, 3, 1, 1) = 1: objs(6, 3, 2, 1) = 1: objs(6, 3, 3, 1) = 1: objs(6, 3, 4, 1) = 0
objs(6, 3, 1, 2) = 1: objs(6, 3, 2, 2) = 0: objs(6, 3, 3, 2) = 0: objs(6, 3, 4, 2) = 0
objs(6, 3, 1, 3) = 0: objs(6, 3, 2, 3) = 0: objs(6, 3, 3, 3) = 0: objs(6, 3, 4, 3) = 0
objs(6, 3, 1, 4) = 0: objs(6, 3, 2, 4) = 0: objs(6, 3, 3, 4) = 0: objs(6, 3, 4, 4) = 0

objs(6, 4, 1, 1) = 1: objs(6, 4, 2, 1) = 0: objs(6, 4, 3, 1) = 0: objs(6, 4, 4, 1) = 0
objs(6, 4, 1, 2) = 1: objs(6, 4, 2, 2) = 0: objs(6, 4, 3, 2) = 0: objs(6, 4, 4, 2) = 0
objs(6, 4, 1, 3) = 1: objs(6, 4, 2, 3) = 1: objs(6, 4, 3, 3) = 0: objs(6, 4, 4, 3) = 0
objs(6, 4, 1, 4) = 0: objs(6, 4, 2, 4) = 0: objs(6, 4, 3, 4) = 0: objs(6, 4, 4, 4) = 0

End Sub

Function isHit(x As Integer, y As Integer, p As Integer)
        isHit = False
        Dim xi As Integer, yi As Integer
        For xi = 1 To 4
                For yi = 1 To 4
                        If okLimit(x + xi - 1, y + yi - 1) = False And obj(p, xi, yi) <> 0 Then isHit = True
                        If okLimit(x + xi - 1, y + yi - 1) = True And obj(p, xi, yi) <> 0 Then
                                If map(x + xi - 1, y + yi - 1) <> 0 Then
                                        isHit = True
                                End If
                        End If
                Next
        Next
End Function

Sub markMap()
        For xi = 0 To 3
                For yi = 0 To 3
                        If obj(objp, xi + 1, yi + 1) <> 0 Then
                                map(objx + xi, objy + yi) = objc
                        End If
                Next
        Next
End Sub
Sub refreshMap()
    Dim xi As Integer, yi As Integer
    For yi = 1 To 20
        For xi = 1 To 15
            frmMain.Line (xi * MULTX - MULTX + STARTX, yi * MULTY - MULTY + STARTY)-Step(MULTX - 1, MULTY - 1), QBColor(map(xi, yi)), BF
        Next
    Next
End Sub

Function nextOf(num As Integer)
        If num < 4 Then nextOf = num + 1 Else nextOf = 1
End Function

Function okLimit(x As Integer, y As Integer)
        okLimit = False
        If x >= 1 And x <= 15 And y >= 1 And y <= 20 Then okLimit = True
End Function

Sub showObj(x As Integer, y As Integer, action As Integer)
        Dim xi As Integer, yi As Integer
        If action = 1 Then action = objc Else action = 0
        For xi = 0 To 3
                For yi = 0 To 3
                        If obj(objp, xi + 1, yi + 1) <> 0 And okLimit(x + xi, y + yi) Then
                                frmMain.Line ((x + xi) * MULTX - MULTX + STARTX, (y + yi) * MULTY - MULTY + STARTY)-Step(MULTX - 1, MULTY - 1), QBColor(action), BF
                        End If
                Next
        Next
End Sub
Sub showNext()
        Dim xi As Integer, yi As Integer
        frmMain.Line (408, 376)-Step(14 * 4 + 1, 14 * 4 + 1), 0, BF
        For xi = 1 To 4
                For yi = 1 To 4
                    If objs(nextb, nextp, xi, yi) <> 0 Then
                        frmMain.Line (xi * 14 - 14 + 408, yi * 14 - 14 + 376)-Step(14, 14), QBColor(nextc), BF
                    End If
                Next
        Next
End Sub

