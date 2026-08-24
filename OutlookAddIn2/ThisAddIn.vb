Public Class ThisAddIn

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

    End Sub

    Private Sub Application_NewMail() Handles Application.NewMail
        Dim objItem As Outlook.MailItem
        'Dim objMailItem As Outlook.MailItem
        Dim strEmailAddress As String
        Dim strSubject As String
        Dim strBody As String
        Dim olFolderInbox As Outlook.MAPIFolder

        olFolderInbox = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)

        ' Get the item that likely triggered the event.   
        objItem = olFolderInbox.Items(0)

        strSubject = objItem.Subject & " From: " & objItem.SenderName
        strBody = objItem.Body
        strEmailAddress = objItem.SenderEmailAddress

        Console.WriteLine(strEmailAddress & ": " & strSubject)

        MsgBox(strEmailAddress & ": " & strSubject)
        
        'objMailItem = Nothing
    End Sub
End Class
