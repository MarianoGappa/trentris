VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trentris"
   ClientHeight    =   7500
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tiempo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   379
      X2              =   379
      Y1              =   0
      Y2              =   504
   End
   Begin VB.Label score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label level 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Viene:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5685
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu mnumnuNew 
      Caption         =   "Nuevo Juego"
   End
   Begin VB.Menu mnuPausar 
      Caption         =   "Pausar"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnumnuRanks 
      Caption         =   "Ver Ranking"
   End
   Begin VB.Menu mnumnuAbout 
      Caption         =   "Autor"
   End
   Begin VB.Menu mnuSound 
      Caption         =   "Sacar Sonido"
   End
   Begin VB.Menu mnuPremios 
      Caption         =   "Ver Premios"
      Begin VB.Menu mnuP1 
         Caption         =   "#1"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuP2 
         Caption         =   "#2"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuP3 
         Caption         =   "#3"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuP4 
         Caption         =   "#4"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuP5 
         Caption         =   "#5"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuP6 
         Caption         =   "#6"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuP7 
         Caption         =   "#7"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuP8 
         Caption         =   "#8"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuP9 
         Caption         =   "#9"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnumnuSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then End
    If tiempo.Enabled Then
        Select Case KeyCode
        Case vbKeyLeft
                If isHit(objx - 1, objy, objp) = False Then
                        showObj objx, objy, 0
                        objx = objx - 1
                        showObj objx, objy, 1
                End If
        Case vbKeyRight
                If isHit(objx + 1, objy, objp) = False Then
                        showObj objx, objy, 0
                        objx = objx + 1
                        showObj objx, objy, 1
                End If
        Case vbKeyUp
                If isHit(objx, objy, nextOf(objp)) = False Then
                        showObj objx, objy, 0
                        objp = nextOf(objp)
                        showObj objx, objy, 1
                End If
        Case vbKeyDown
                If Not auxchoto Then
                    tiempo.Enabled = False
                    tiempo.Interval = 50
                    tiempo.Enabled = True
                    auxchoto = True
                End If
        Case 32
                While isHit(objx, objy + 1, objp) = False
                    showObj objx, objy, 0
                    objy = objy + 1
                    showObj objx, objy, 1
                Wend
        End Select
    End If


End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If tiempo.Enabled Then
        Select Case KeyCode
        Case vbKeyDown
              tiempo.Interval = levelSpeed(level.Caption)
              auxchoto = False
        End Select
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    'close all sub forms
    ModifyRankings "LIBRA", Val(score.Caption)
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
End Sub

Private Sub GO_Click()
    tiempo.Enabled = True
    
End Sub

Private Sub mnumnuAbout_Click()
    frmAbout.Show
    
End Sub

Private Sub mnumnuNew_Click()
    ModifyRankings "LIBRA", Val(score.Caption)
    startUp
    mnuPausar.Enabled = True
    mnuPausar.Caption = "Pausar"
End Sub

Private Sub mnuPausar_Click()
    If mnuPausar.Caption = "Pausar" Then
        mnuPausar.Caption = "Seguir"
        tiempo.Enabled = False
    Else
        mnuPausar.Caption = "Pausar"
        tiempo.Enabled = True
    End If
End Sub
Private Sub mnuP1_Click()
    frmPremio.Show
    frmPremio.exePremio (0)
End Sub

Private Sub mnuP2_Click()
    frmPremio.Show
    frmPremio.exePremio (1)
End Sub

Private Sub mnuP3_Click()
    frmPremio.Show
    frmPremio.exePremio (2)
End Sub

Private Sub mnuP4_Click()
    frmPremio.Show
    frmPremio.exePremio (3)
End Sub

Private Sub mnuP5_Click()
    frmPremio.Show
    frmPremio.exePremio (4)
End Sub

Private Sub mnuP6_Click()
    frmPremio.Show
    frmPremio.exePremio (5)
End Sub

Private Sub mnuP7_Click()
    frmPremio.Show
    frmPremio.exePremio (6)
End Sub

Private Sub mnuP8_Click()
    frmPremio.Show
    frmPremio.exePremio (7)
End Sub

Private Sub mnuP9_Click()
    frmPremio.Show
    frmPremio.exePremio (8)
End Sub

Private Sub mnumnuRanks_Click()
    frmRanking.Show
End Sub

Private Sub mnumnuSalir_Click()
End
End Sub

Private Sub mnuSound_Click()
    If soundEnabled Then
        soundEnabled = False
        frmMain.mnuSound.Caption = "Poner Sonido"
    Else
        soundEnabled = True
        frmMain.mnuSound.Caption = "Sacar Sonido"
    End If

End Sub

Private Sub tiempo_Timer()
        Dim text As String
        If isHit(objx, objy + 1, objp) = False Then
                showObj objx, objy, 0
                objy = objy + 1
                showObj objx, objy, 1
        Else
                markMap
                checkLines
                chooseObj
                ' Pierde porque la siguiente pieza no entra
                If isHit(objx, objy, objp) Then
                    ' Si superó su récord
                    If levelup Then
                        higherlevel = higherlevel + 1
                        PlaySoundX "win.wav"
                        text = "Se te adjudicó el premio #" + Str(higherlevel - 1) + "!"
                        MsgBox text, , "Premio"
                    Else
                        PlaySoundX "mujaja.wav"
                    End If
                    ModifyRankings "LIBRA", Val(score.Caption)
                    setPremios (higherlevel)
                    startUp
                    tiempo.Enabled = False
                    frmMain.Caption = "Trentris"
                    mnuPausar.Enabled = False
                    ' Adjudicar premio
                End If
        End If

End Sub
