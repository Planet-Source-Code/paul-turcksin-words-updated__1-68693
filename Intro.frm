VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00800000&
   Caption         =   "Introduction"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6240
   Icon            =   "Intro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Holy sh**, I don't want to be part of this nonsense!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   6015
   End
   Begin VB.CommandButton cmdGiveMeMore 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   6015
   End
   Begin VB.Label lblIntro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Educational fun!
'
' Inspired by "//Dil Se Desi// Only great minds can read this"
' Thanks to Sunny Sing for sending me this stuff.
'
' Icon:
'    Author: Icons-Land (Taras Berezyuk)
'    Homepage: http://www.icons-land.com/
'
'    This is a Vista icon and I used Robert Rayment's Tiny GFX  to make it "normal" ;)

' Paul Turcksin, May 2007
'
' Version 1.1 implements a suggestion made by GriGri to avoid adjacent same letters.
'   The solution implemented is far from ideal and will not always yield the intended
'   improvement. The adjacent same letter problem is best visible with long words,
'   and transformed long words are usually ard to read.

Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdGiveMeMore_Click()
   frmMain.Show
   Unload Me
End Sub

Private Sub Form_Load()
   lblIntro = "The phaonmneal pweor of the hmuan mnid, aoccdrnig to a rscheearch at " & _
                  "Cmabrigde Uinervtisy, it dseno't mtaetr in waht oerdr the ltteres in " & _
                  "a wrod are, the olny iproamtnt tihng is taht the frsit and lsat ltteer " & _
                  "be in the rghit pclae. The rset can be a taotl mses and you can sitll " & _
                  "raed it whotuit a pboerlm. Tihs is bcuseae the huamn mnid deos not raed " & _
                  "ervey lteter by istlef, but the wrod as a wlohe. Azanmig huh? yaeh and " & _
                  "I awlyas tghuhot slpeling was ipmorantt!" & vbCrLf & vbCrLf & _
                   "Cna yuo raed tihs?" & vbCrLf & vbCrLf & _
              "Olny 55 plepoe out of 100 can."
              
   cmdGiveMeMore.Caption = "Yes I cloud and wnat to try mroe of tihs by mleysf!"

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmIntro = Nothing
End Sub
