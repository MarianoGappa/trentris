VERSION 5.00
Begin VB.Form frmPremio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Premio"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox gif 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4545
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   2160
         Top             =   2040
      End
   End
End
Attribute VB_Name = "frmPremio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cambio As Integer, indice As Integer
Dim gifs(8, 3) As String

Private Sub Form_Load()
    cambio = 0
    loadPremios
End Sub
Sub loadPremios()
    gifs(0, 0) = "1.wav"
    gifs(0, 1) = "11.jpg"
    gifs(0, 2) = "12.jpg"
    gifs(0, 3) = "13.jpg"
    
    gifs(1, 0) = "2.wav"
    gifs(1, 1) = "21.jpg"
    gifs(1, 2) = "22.jpg"
    gifs(1, 3) = "23.jpg"
    
    gifs(2, 0) = "3.wav"
    gifs(2, 1) = "31.jpg"
    gifs(2, 2) = "32.jpg"
    gifs(2, 3) = "33.jpg"
    
    gifs(3, 0) = "4.wav"
    gifs(3, 1) = "41.jpg"
    gifs(3, 2) = "42.jpg"
    gifs(3, 3) = "43.jpg"
    
    gifs(4, 0) = "5.wav"
    gifs(4, 1) = "51.jpg"
    gifs(4, 2) = "52.jpg"
    gifs(4, 3) = "53.jpg"
    
    gifs(5, 0) = "6.wav"
    gifs(5, 1) = "61.jpg"
    gifs(5, 2) = "62.jpg"
    gifs(5, 3) = "63.jpg"
    
    gifs(6, 0) = "7.wav"
    gifs(6, 1) = "71.jpg"
    gifs(6, 2) = "72.jpg"
    gifs(6, 3) = "73.jpg"
    
    gifs(7, 0) = "8.wav"
    gifs(7, 1) = "81.jpg"
    gifs(7, 2) = "82.jpg"
    gifs(7, 3) = "83.jpg"
    
    gifs(8, 0) = "9.wav"
    gifs(8, 1) = "91.jpg"
    gifs(8, 2) = "92.jpg"
    gifs(8, 3) = "93.jpg"
End Sub
Sub exePremio(num As Integer)
    indice = num
    cambio = 1
    gif.Picture = LoadPicture(gifs(indice, cambio + 1))
    PlaySoundX gifs(indice, 0)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If cambio = 3 Then cambio = 0
    gif.Picture = LoadPicture(gifs(indice, cambio + 1))
    cambio = cambio + 1
End Sub
