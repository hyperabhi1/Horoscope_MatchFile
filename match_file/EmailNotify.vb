Imports System.Net
Imports System.Net.Mail

Module EmailNotify
    Public ConsoleLogs = ""
    Sub SendEmail(ByVal Subject As String, ByVal Time As DateTime, ByVal flag As String)
        Dim mail As New MailMessage
        mail.To.Add("abhishek.singh@enukesoftware.com")
        'mail.To.Add("vcdubai@gmail.com")
        mail.From = New MailAddress("iamemailsender@gmail.com")
        mail.Subject = Subject
        mail.IsBodyHtml = True
        If flag = "start" Then
            mail.Body += "<br><h1>Start Time: </h1>" + Time + "<br>"
        End If
        If flag = "end" Then
            mail.Body += "<br><h1>Total Time: </h1>" + DateTime.Now.Subtract(Time).ToString() + "<br>"
            mail.Body += "<br><h3>Console Logs</h3><br>"
            mail.Body += ConsoleLogs
        End If
        Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(mail.Body, Nothing, "text/html")
        mail.AlternateViews.Add(htmlView)
        Dim smtp As New SmtpClient
        smtp.Host = "smtp.gmail.com"
        smtp.Credentials = New NetworkCredential("iamemailsender@gmail.com", "dakiyadaaklaya")
        smtp.EnableSsl = True
        Dim IsMailSent As Boolean = True
        smtp.Send(mail)
    End Sub
End Module
