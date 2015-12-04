'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

<Runtime.InteropServices.ComVisible(True)> _
    Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI



    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("OutlookAddIn1.Ribbon1.xml")
    End Function



    'Обработка ленты

    Public Sub On_click_to_dial(ByVal control As Office.IRibbonControl)
        Dim f As New Form1()
        f.ShowDialog()
    End Sub




    'Обработка контекстных меню

    Public Sub OnClickToCall_W(
     ByVal control As Office.IRibbonControl)

        Try
            Dim card As Office.IMsoContactCard =
                TryCast(control.Context, Office.IMsoContactCard)

            If card IsNot Nothing Then

                Dim Phone As String = GetWorkPhone(card)
                If Phone IsNot Nothing Then
                    Phone = RemoveWhitespace(Phone)

                    Dim server As String = My.Settings.Server
                    Dim port As String = My.Settings.Port

                    Dim URL = "https://" & server & ":" & port & "/webdialer/Webdialer?destination=" & Phone

                    'Try
                    'System.Diagnostics.Process.Start(URL)
                    'Catch ex As Exception
                    'MsgBox(ex.Message)
                    'End Try
                    Dim wshShell = CreateObject("WScript.Shell")
                    Dim path As String = wshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

                    Dim RetCode = Shell(path & "\Internet Explorer\IEXPLORE.EXE " & URL, vbNormalFocus)
                Else
                    MsgBox("Рабочий телефон для данного контакта не определен")
                End If

            Else
                MsgBox("Карточка контакта недоступна")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub OnClickToCall_M(
      ByVal control As Office.IRibbonControl)

        Try
            Dim card As Office.IMsoContactCard =
                TryCast(control.Context, Office.IMsoContactCard)

            If card IsNot Nothing Then
                Dim Phone As String = GetMobilePhone(card)


                If Phone IsNot Nothing Then

                    Phone = RemoveWhitespace(Phone)

                    If Phone.Chars(0) = "+" Then
                        Phone = Phone.Remove(0, 2)
                    Else
                        Phone = Phone.Remove(0, 1)
                    End If
                    Phone = My.Settings.local_numbers & Phone

                    Dim server As String = My.Settings.Server
                    Dim port As String = My.Settings.Port

                    Dim URL = "https://" & server & ":" & port & "/webdialer/Webdialer?destination=" & Phone

                    Dim wshShell = CreateObject("WScript.Shell")
                    Dim path As String = wshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

                    Dim RetCode = Shell(path & "\Internet Explorer\IEXPLORE.EXE " & URL, vbNormalFocus)
                Else
                    MsgBox("Мобильный телефон для данного контакта не определен")
                End If

            Else
                MsgBox("Карточка контакта недоступна")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub OnClickToCall_W_Contact(
      ByVal control As Office.IRibbonControl)
        Try
            Dim selected As Outlook.Selection = TryCast(control.Context, Outlook.Selection)
            Dim x As System.Collections.IEnumerator = selected.GetEnumerator
            x.MoveNext()
            Dim card As Outlook.ContactItem = TryCast(x.Current, Outlook.ContactItem)

            If card IsNot Nothing Then
                Dim Phone As String = card.BusinessTelephoneNumber
                If Phone IsNot Nothing Then
                    Phone = RemoveWhitespace(Phone)

                    Dim server As String = My.Settings.Server
                    Dim port As String = My.Settings.Port

                    Dim URL = "https://" & server & ":" & port & "/webdialer/Webdialer?destination=" & Phone

                    Dim wshShell = CreateObject("WScript.Shell")
                    Dim path As String = wshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

                    Dim RetCode = Shell(path & "\Internet Explorer\IEXPLORE.EXE " & URL, vbNormalFocus)
                Else
                    MsgBox("Мобильный телефон для данного контакта не определен")
                End If

            Else
                MsgBox("Карточка контакта недоступна")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub OnClickToCall_M_Contact(
      ByVal control As Office.IRibbonControl)
        Try
            Dim selected As Outlook.Selection = TryCast(control.Context, Outlook.Selection)
            Dim x As System.Collections.IEnumerator = selected.GetEnumerator
            x.MoveNext()
            Dim card As Outlook.ContactItem = TryCast(x.Current, Outlook.ContactItem)

            If card IsNot Nothing Then
                Dim Phone As String = card.MobileTelephoneNumber
                If Phone IsNot Nothing Then
                    Phone = RemoveWhitespace(Phone)
                    If Phone.Chars(0) = "+" Then
                        Phone = Phone.Remove(0, 2)
                    Else
                        Phone = Phone.Remove(0, 1)
                    End If
                    Phone = My.Settings.local_numbers & Phone


                    Dim server As String = My.Settings.Server
                    Dim port As String = My.Settings.Port

                    Dim URL = "https://" & server & ":" & port & "/webdialer/Webdialer?destination=" & Phone

                    Dim wshShell = CreateObject("WScript.Shell")
                    Dim path As String = wshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

                    Dim RetCode = Shell(path & "\Internet Explorer\IEXPLORE.EXE " & URL, vbNormalFocus)
                Else
                    MsgBox("Мобильный телефон для данного контакта не определен")
                End If

            Else
                MsgBox("Карточка контакта недоступна")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Function GetWorkPhone(ByVal card As Office.IMsoContactCard) As String
        If card.AddressType =
            Office.MsoContactCardAddressType.msoContactCardAddressTypeOutlook Then
            Dim host As Outlook.Application = Globals.ThisAddIn.Application
            Dim ae As Outlook.AddressEntry =
                host.Session.GetAddressEntryFromID(card.Address)

            If (ae.AddressEntryUserType =
              Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry _
              OrElse ae.AddressEntryUserType = _
              Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) Then
                Dim ex As Outlook.ExchangeUser = _
                  ae.GetExchangeUser()
                Dim cu As Outlook.AddressEntry = _
                    host.Session.CurrentUser.AddressEntry.GetExchangeUser()
                If cu.Address <> ex.Address Then
                    Return ex.BusinessTelephoneNumber
                Else
                    Throw New Exception("Вы пытаетесь позвонить себе")
                End If
            Else
                Throw New Exception("Некорретный пользователь")
            End If
        Else
            Throw New Exception("Некорретный пользователь")
        End If
    End Function

    Private Function GetMobilePhone(ByVal card As Office.IMsoContactCard) As String
        If card.AddressType =
            Office.MsoContactCardAddressType.msoContactCardAddressTypeOutlook Then
            Dim host As Outlook.Application = Globals.ThisAddIn.Application
            Dim ae As Outlook.AddressEntry =
                host.Session.GetAddressEntryFromID(card.Address)

            If (ae.AddressEntryUserType =
             Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry _
             OrElse ae.AddressEntryUserType = _
             Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) Then
                Dim ex As Outlook.ExchangeUser = _
                  ae.GetExchangeUser()
                Dim cu As Outlook.AddressEntry = _
                    host.Session.CurrentUser.AddressEntry.GetExchangeUser()
                Return ex.MobileTelephoneNumber
            Else
                Throw New Exception("Некорретный пользователь")
            End If
        Else
            Throw New Exception("Некорретный пользователь")
        End If
    End Function

    Private Function RemoveWhitespace(fullString As String) As String
        Return New String(fullString.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray())
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1.
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub



#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
