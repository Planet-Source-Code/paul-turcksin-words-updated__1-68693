VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00800000&
   Caption         =   "Can you read this?"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6495
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox fTransformedText 
      BackColor       =   &H00C0FFFF&
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Main.frx":0CCA
      Top             =   2400
      Width           =   6255
   End
   Begin VB.CommandButton cmdTransform 
      Caption         =   "Transform"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox fOriginalText 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Main.frx":0CD0
      Top             =   480
      Width           =   6255
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Transformed text"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter text "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arSubStrings() As String
Private arTransformedSubStrings() As String

Private Sub cmdTransform_Click()

   Dim i As Long
   Dim l As Long
   Dim ws As String

' the textbox must contain some text, at least 16 characters (why 16? why not?)
   If Len(fOriginalText.Text) < 16 Then
      MsgBox "The original text must at least be 16 characters long."
      Exit Sub
   End If
   
' as the Split function doesn't like "special" characters, we replace hem by spaces.
   For i = 1 To Len(fOriginalText.Text)
      Select Case Mid$(fOriginalText.Text, i, 1)
         Case "0" To "9", "a" To "z", "A" To "Z"
            ws = ws & Mid$(fOriginalText.Text, i, 1)
         Case Else
            ws = ws & " "
      End Select
   Next i
      
' split text into individual words
   arSubStrings = Split(ws)
   
' perform transformtion on each word in the array and store in  new array
   ReDim arTransformedSubStrings(UBound(arSubStrings))
   For i = 0 To UBound(arSubStrings)
      arTransformedSubStrings(i) = fncTransformWord(arSubStrings(i))
   Next i
   
' finally replace the words in the original by the transformed ones
   ws = fOriginalText
   l = 1
   For i = 0 To UBound(arSubStrings)
      l = InStr(l, fOriginalText.Text, arSubStrings(i))
      Mid$(ws, l, Len(arSubStrings(i))) = arTransformedSubStrings(i)
   Next i
   fTransformedText = ws

End Sub

Private Sub Form_Load()

   fOriginalText = "The quick brown fox jumped over the lazy dog." & vbCrLf & _
                   "------" & vbCrLf & _
                   "This should be fun. " & vbCrLf & _
                   "You can overwrite this text or paste some phrases you like to see transformed." & vbCrLf & _
                   "If you particularly like the transformation you can select it and copy it to the clipboard."

   cmdTransform_Click
End Sub

'=========================================================================================
'
'                                  LOCAL PROCEDURES
'_________________________________________________________________________________________

Private Function fncTransformWord(sWord As String) As String
' scramble letters of the input except the first and last one
   Dim i As Long
   Dim lLengthWord As Long
   Dim arLetters() As String * 1
   Dim arScramble() As String * 1
   Dim r As Integer
   Dim Char As String * 1
   
' length of this word
   lLengthWord = Len(sWord)
   
' if the length is less than 4, there is not much to transform
   If lLengthWord < 3 Then
      Exit Function
   End If
   
' dim array with all letters of the word and an array for the leters to scramble
   ReDim arLetters(lLengthWord - 1)   ' minus 1 because its zero based
   ReDim arScramble(lLengthWord - 3)  ' minus 2 (first/last) and minus 1 (zero based)
   
' and put letters in array
   For i = 0 To lLengthWord - 1
      arLetters(i) = Mid$(sWord, i + 1, 1)
   Next i
   
' make a copy of the letters to scramble
   For i = 1 To lLengthWord - 2
      arScramble(i - 1) = arLetters(i)
   Next i
   
'shuffle (scramble) this array.
   
   Randomize Timer
' Transpose i with a randomly chosen number <=i. We have to preserve the position
   For i = lLengthWord - 3 To 1 Step -1
      r = Int(Rnd * i)
      Char = arScramble(r)
      arScramble(r) = arScramble(i)
      arScramble(i) = Char
   Next i
   
' put shuffled letters into arraywith original
   For i = 0 To lLengthWord - 3
      arLetters(i + 1) = arScramble(i)
  Next i
   
' check if we have identical adjecent letters like "oo" or "ll". If so exhange with previous
' or next. Only do this check if more than 2 letters in transformation.
   If UBound(arScramble) > 2 Then
      For i = 0 To UBound(arLetters) - 1
         If arLetters(i) = arLetters(i + 1) Then
            If i = 0 Then   ' first letter cannot move, so we swap letters 2 and 3
               Char = arLetters(1)
               arLetters(1) = arLetters(2)
               arLetters(2) = Char
            ElseIf i = lLengthWord - 2 Then   ' last letter cannot move
               Char = arLetters(i - 1)
               arLetters(i - 1) = arLetters(i - 2)
               arLetters(i - 2) = Char
            Else                   ' somewhere in between
               Char = arLetters(i)
               arLetters(i) = arLetters(i - 1)
               arLetters(i - 1) = Char
            End If
         End If
      Next i
   End If

' copy result to output
   For i = 0 To lLengthWord - 1
   fncTransformWord = fncTransformWord & arLetters(i)
   Next i

End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frmMain = Nothing
End Sub
