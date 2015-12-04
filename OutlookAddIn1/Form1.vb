Public Class Form1

    Private Sub Click_to_Call_Click(sender As Object, e As EventArgs) Handles Click_to_Dial.Click

        Try

            Dim Phone As String = TextBox1.Text

            If Phone IsNot Nothing And Phone <> "" Then
                Phone = RemoveWhitespace(Phone)
                If Phone.Length < 4 Then
                    MsgBox("Неверная длина номера телефона")
                    Return
                ElseIf Phone.Length > 4 And Phone.Length < 10 Then
                    MsgBox("Неверная длина номера телефона")
                    Return
                ElseIf Phone.Length = 10 Then
                    Phone = My.Settings.local_numbers & Phone
                ElseIf Phone.Length = 11 Then
                    Phone = Phone.Remove(0, 1)
                    Phone = My.Settings.local_numbers & Phone
                ElseIf Phone.Length >= 12 And Phone.Chars(0) = "+" And Phone.Chars(1) = "7" Then
                    Phone = Phone.Remove(0, 2)
                    Phone = My.Settings.local_numbers & Phone
                ElseIf Phone.Length >= 12 And Phone.Chars(0) = "+" And Phone.Chars(1) <> "7" Then
                    Phone = Phone.Remove(0, 1)
                    Phone = My.Settings.foreign_numbers & Phone
                End If


                Dim server As String = My.Settings.Server
                Dim port As String = My.Settings.Port

                Dim URL = "https://" & server & ":" & port & "/webdialer/Webdialer?destination=" & Phone

                Dim wshShell = CreateObject("WScript.Shell")
                Dim path As String = wshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

                Dim RetCode = Shell(path & "\Internet Explorer\IEXPLORE.EXE " & URL, vbNormalFocus)

                Me.Close()
            Else
                MsgBox("Введите номер телефона")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Click_to_Exit_Click(sender As Object, e As EventArgs) Handles Click_to_Exit.Click
        Me.Close()
    End Sub

    Private Function RemoveWhitespace(fullString As String) As String
        Return New String(fullString.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray())
    End Function

    Private Sub TextBox1_Return(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress

        If Asc(e.KeyChar) = 13 Then
            Click_to_Call_Click(AcceptButton, e)
        End If


    End Sub

    Private Sub LoadWindowPosition()
        Dim ptLocation As System.Drawing.Point = My.Settings.WindowLocation

        If (ptLocation.X = -1) And (ptLocation.Y = -1) Then
            Return
        End If

        Dim bLocationVisible As Boolean = False
        For Each S As Windows.Forms.Screen In Windows.Forms.Screen.AllScreens
            If S.Bounds.Contains(ptLocation) Then
                bLocationVisible = True
            End If
        Next

        If Not bLocationVisible Then
            Return
        End If

        Me.StartPosition = Windows.Forms.FormStartPosition.Manual
        Me.Location = ptLocation
        Me.Size = My.Settings.WindowSize

    End Sub

    Private Sub Form1_Load(sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        LoadWindowPosition()

        Dim iData As Windows.Forms.IDataObject = Windows.Forms.Clipboard.GetDataObject()
        'Check to see if the data is in a text format
        If iData.GetDataPresent(Windows.Forms.DataFormats.Text) Then
            'If it's text, then paste it into the textbox
            TextBox1.SelectedText = CType(iData.GetData(Windows.Forms.DataFormats.Text), String)
        End If

    End Sub

    Private Sub Form1_Exit(sender As Object, ByVal e As EventArgs) Handles MyBase.Closing
        My.Settings.WindowLocation = Me.Location
        My.Settings.WindowSize = Me.Size


        My.Settings.Save()
    End Sub

End Class