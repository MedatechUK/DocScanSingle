Imports System.Reflection

Module Module1
    Dim WithEvents cApp As New ConsoleApp.CA
    Enum myRunMode As Integer
        file = 0
        url = 1
        Config = 2
        enviro = 3
    End Enum
    Dim FILENAME As String = Nothing
    Sub Main()
        'ScanControl.write_error("its dead jim testy testy test", 1)
        '/e "010123 \\hannibal\emerge\test testtest ttyyttyy"
        Try
            With cApp
                .GetArgs(Command)
                Select Case .RunMode
                    Case myRunMode.Config
                        Dim g As New UserSettings(True)
                        g.ShowDialog()
                        .Quit = True

                    Case myRunMode.file
                        .doWelcome(Assembly.GetExecutingAssembly())

                        Dim LineOfData() As String = FILENAME.Split(" ")
                        ScanControl.scannall( _
                            New ScanControl.ScanDocument( _
                                LineOfData(3), _
                                LineOfData(4), _
                                LineOfData(2), _
                                LineOfData(1) _
                            ) _
                        )

                End Select

            End With

        Catch ex As Exception
            Console.WriteLine(ex.ToString)

        End Try

    End Sub

    Private Sub cApp_Switch(ByVal StrVal As String, ByRef State As String, ByRef Valid As Boolean) Handles cApp.Switch
        Try
            With cApp
                Select Case StrVal
                    Case "runmode"
                        State = "rm"
                    Case "config"
                        .RunMode = myRunMode.Config
                        State = Nothing
                    Case "p", "prn", "printer"
                        State = "p"
                    Case "f", "file"
                        State = "f"
                    Case "u", "url"
                        State = "u"
                    Case "e"
                        State = "e"
                    Case Else
                        Valid = False
                End Select
            End With
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Sub

    Private Sub cApp_SwitchVar(ByVal State As String, ByVal StrVal As String, ByRef Valid As Boolean) Handles cApp.SwitchVar
        Try
            With cApp
                Select Case State
                    Case "f"
                        .RunMode = myRunMode.file
                        FILENAME = StrVal

                    Case "config"
                        .RunMode = myRunMode.Config                        

                    Case Else
                        Valid = False

                End Select
            End With
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try
    End Sub

End Module
