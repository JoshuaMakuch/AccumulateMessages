'Joshua Makuch
'RCET0265
'Fall 2023
'Accumulate Messages
'https://github.com/JoshuaMakuch/AccumulateMessages.git

Option Strict On
Option Explicit On

Imports System

Module MessageProgram
    Sub Main(args As String())
        'uncomment to test interactively
        'Test.Manual()
        Test.Auto()
    End Sub

    Function UserMessages(ByVal newMessage As String, ByVal clear As Boolean) As String

        Static userMessagesStored As String 'Creates a variable the persists after every function call

        'Determines whether or not the new message is not just a blank line and amends accordingly or clears based on the clear variable
        If clear Then
            userMessagesStored = ""
        ElseIf newMessage = "" Then
            Return userMessagesStored
        Else
            userMessagesStored = userMessagesStored & newMessage & vbCrLf
        End If

        Return userMessagesStored

    End Function


End Module
