VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Just a crap thing to Test"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "'"
      Top             =   4440
      Width           =   5895
   End
   Begin VB.TextBox Text2 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4800
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Skip"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I just make this only to strip out the comment or something that
'I don't need in my work with out going to select them and delete.
'Because of using string scan are too slow, so this method is fast and easy.
'I use an array to hold the very string but this is not a good idea coz
'It's uses much memory in large text file.
'If you know how to skip any unneeded character without using a bunch of array
'Please tell me or send me a source code or link or your source code.

'-------------------------->>> CRIPTEDEM@YAHOO.COM <<<----------------------------

Private Sub Command1_Click()
'I use vb comment and remove them
'If you want to test the other place them in the text box
'Note that this is a case sencetive

Dim TheString() As String      'Place where the string is hold
Dim lCnt, Cnt, lPos As Integer

'Firts split the string with Chr$(13) & Chr$(10))
'It will be added latter on

TheString = Split(Text1.Text, Chr$(13) & Chr$(10))

For lCnt = 0 To UBound(TheString)
    
    If Mid$(TheString(lCnt), 1, 1) = Text3.Text Then 'Find in first character occorence in string
    'Do nothing
    'This will skip the array which has the string "'"
    Else
    'Find comment in the array string if any
    'This will remove all string after the unwanted string
        For Cnt = 1 To Len(TheString(lCnt))
            lPos = InStr(Cnt, TheString(lCnt), Text3.Text)
                If lPos <> 0 Then
                    Text2.Text = Text2.Text & Mid$(TheString(lCnt), 1, lPos - Cnt) & Chr$(13) & Chr$(10)
                    Exit For
                Else
                    'If no unwanted string in string array then add them
                    Text2.Text = Text2.Text & TheString(lCnt) & Chr$(13) & Chr$(10)
                    Exit For
                End If
        Next
    End If

Next
End Sub
