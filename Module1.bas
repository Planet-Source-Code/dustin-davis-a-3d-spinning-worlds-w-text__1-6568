Attribute VB_Name = "Module1"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'This is for screen saver purposes, its not needed'
'''''''''''''''''''''''''''''''''''''''''''''''''''

Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal x%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer



Sub Main()

Dim strCmdLine As String

strCmdLine = Left(Command, 2)


    If strCmdLine = "/p" Then
        Form1.Show
    ElseIf strCmdLine = "/c" Then
        Form3.Show
    Else
        Form1.Show
    End If

End Sub

