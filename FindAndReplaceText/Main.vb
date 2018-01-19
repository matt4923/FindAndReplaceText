Imports System
Imports System.IO

Public Class Main

    Public Shared Sub Main()
        Dim myInfo = ProcessCommandLine()
        If myInfo Is Nothing Then Return

        Dim fileText As String = My.Computer.FileSystem.ReadAllText(myInfo.GetFullFromPath)
        fileText = Replace(fileText, myInfo.FindText, myInfo.ReplaceText)
        My.Computer.FileSystem.WriteAllText(myInfo.GetFullToPath, fileText, False)
    End Sub

    Private Shared Function ProcessCommandLine() As MyCommandInfo
        Dim myInfo As New MyCommandInfo

        Dim allCommand = My.Application.CommandLineArgs
        For Each item In allCommand
            If item = Args.Help OrElse item = Args.AltHelp OrElse item = Args.AnotherHelp Then
                MsgBox(Args.ShowHelp, MsgBoxStyle.Information, "Help")
                Return Nothing
            Else
                Dim posEqual = InStr(item, "=")
                If posEqual = 0 Then
                    Throw New ApplicationException("Invalid format.  There must be an = after each argument title.")
                End If
                Dim name = Left(item, posEqual - 1)
                Dim data = Mid(item, posEqual + 1)

                Select Case name
                    Case Args.Path
                        myInfo.Path = data
                    Case Args.FromFile
                        myInfo.FromFile = data
                    Case Args.ToFile
                        myInfo.ToFile = data
                    Case Args.FindText
                        myInfo.FindText = data
                    Case Args.ReplaceText
                        myInfo.ReplaceText = data
                    Case Else
                        Throw New ApplicationException("Invalid argument name, " & name)
                End Select
            End If
        Next

        If String.IsNullOrEmpty(myInfo.Path) Then
            myInfo.Path = Application.ExecutablePath
        End If

        If String.IsNullOrEmpty(myInfo.FindText) Then
            Throw New ApplicationException(Args.FindText & " parameter is not provided.")
        End If

        If String.IsNullOrEmpty(myInfo.FromFile) Then
            Throw New ApplicationException(Args.FromFile & " parameter is not provided.")
        End If

        If String.IsNullOrEmpty(myInfo.ToFile) Then
            Throw New ApplicationException(Args.ToFile & " parameter is not provided.")
        End If

        Return myInfo
    End Function

End Class

Public Class Args
    Public Const Path As String = "PATH"
    Public Const FromFile As String = "FROMFILE"
    Public Const ToFile As String = "TOFILE"
    Public Const FindText As String = "FIND"
    Public Const ReplaceText As String = "REPLACE"
    Public Const Help As String = "?"
    Public Const AltHelp As String = "HELP"
    Public Const AnotherHelp As String = "/?"

    Public Shared Function ShowHelp() As String
        Return String.Concat("The following command line elements are supported: ", ControlChars.CrLf, _
                Path, ControlChars.CrLf, FromFile, ControlChars.CrLf, ToFile, ControlChars.CrLf, FindText, ControlChars.CrLf, ReplaceText, ControlChars.CrLf, Help,
                ControlChars.CrLf, AltHelp, ControlChars.CrLf, AnotherHelp, ControlChars.CrLf, ControlChars.CrLf, "Example:  ", _
                """PATH=C:\PROGRAM DATA\"" ", "FROMFILE=BeforeReplace.txt ", "TOFILE=AfterReplace.txt", " FIND=[6~", " REPLACE=", _
                ControlChars.CrLf, ControlChars.CrLf, "In the example above, the replace parameter is not needed because the text to replace is empty.",
                ControlChars.CrLf, ControlChars.CrLf, "Things to watch out for:  ", ControlChars.CrLf, _
                "If there is a space in the data for an element, that element must be surrounded by "" characters.", ControlChars.CrLf, _
                "If there is a "" character in the data for an element, it must be replaced by two "" characters.")
    End Function

End Class

Public Class MyCommandInfo
    Public Property Path As String
    Public Property FromFile As String
    Public Property ToFile As String
    Public Property FindText As String
    Public Property ReplaceText As String = String.Empty

    Public Function GetFullFromPath() As String
        Return System.IO.Path.Combine(Path, FromFile)
    End Function

    Public Function GetFullToPath() As String
        Return System.IO.Path.Combine(Path, ToFile)
    End Function

End Class
