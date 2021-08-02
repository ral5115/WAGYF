Imports Microsoft.VisualBasic
Imports System.Net.Mail
Imports System.Net

Public Class clsCorreo

    Public Function EnviarEmail(ByVal pNombreRemitente As String, _
                                ByVal pEmailRemitente As String, _
                                ByVal pEmailDestinatario As String, _
                                ByVal pAsunto As String, _
                                ByVal pMensaje As String, _
                                ByRef pError As String) As String


        Dim smtp As New SmtpClient(My.Settings.smtp)
        smtp.Port = My.Settings.MailPuerto
        smtp.EnableSsl = False
        smtp.Credentials = New NetworkCredential(My.Settings.MailUser, My.Settings.MailPwd)

        smtp.DeliveryMethod = SmtpDeliveryMethod.Network
        smtp.Timeout = 30000

        Dim fromAddress As New MailAddress(My.Settings.MailUser, "GTIntegration Giros y Finanzas")
        Dim subject As String = pAsunto
        Dim body As String = " <br> Mensaje: " & pMensaje

        Dim message As New MailMessage()
        message.To.Add(pEmailDestinatario)
        message.From = fromAddress
        message.Subject = subject
        message.Body = body
        message.IsBodyHtml = True

        Try
            smtp.Send(message)
            pError = "ok"
            Return pError
        Catch ex As Exception
            pError = "ERROR: " & ex.Message
            Return pError
        End Try

    End Function

End Class
