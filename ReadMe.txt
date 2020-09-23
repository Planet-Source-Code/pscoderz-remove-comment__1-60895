Author : LMF	  "Lazy M***** F*****"
Date   : May 2005

'Just a simple method to strip off unwanted character like comment.
'This method is very fast but uses large memory in a larger text file.

Example :

'If you want to skip a comment eg VB comment " ' " just put it in the search string with no spaces.

'============================================================================================================================
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
    'This will skip the string which has the string "'"
    Else
    'Find comment in the string if any
    'This will remove all string after the search string
        For Cnt = 1 To Len(TheString(lCnt))
            lPos = InStr(Cnt, TheString(lCnt), Text3.Text)
                If lPos <> 0 Then
                    Text2.Text = Text2.Text & Mid$(TheString(lCnt), 1, lPos - Cnt) & Chr$(13) & Chr$(10)
                    Exit For
                Else
                    'If no string search then add them
                    Text2.Text = Text2.Text & TheString(lCnt) & Chr$(13) & Chr$(10)
                    Exit For
                End If
        Next
    End If

Next
End Sub

'============================================================================================================================
'The result you can see it in the project

'I think this method sucks but can be useful though...
'Any suggestions and comment please email me at this address 

-------------------------------->>> CRIPTEDEM@YAHOO.COM <<<--------------------------------

'People who appreciate my work are always welcome. 
'To those who don't can just piss me off.
'Anyways i'm just a new comer in Visual Basic so there is so much to learn.
'Happy coding...
