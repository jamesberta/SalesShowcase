<%
'Sends an email
Dim oMailer
Set oMailer = Server.CreateObject("Persits.MailSender")
    With oMailer
        .FromName = Request.Form("AreaName")
        .MailFrom = Request.Form("AreaEmail")
        .AddAddress Request.Form("to")
        .Host = "10.1.1.36"
        .Subject    = Request.Form("Subject")
        .CharSet    = 1
        .Body = Request.Form("AreaComments")
    End With
oMailer.Send
Set oMailer = nothing
%>
