Imports System.Text
Imports System.IO

Public Class frmSincronizacion
    Dim objTarea As New clsTarea
    Dim objDatos As New clsDatos

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub frmSincronizacion_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Try

            Dim resultadoCarga As Short = 1

            'Copia de datos de la tabla TRANSACTION_CV a las tablas habiles y backup
            'objDatos.SP_GT_COPIAR_TRANSACTION_CV()
            objDatos.SP_GT_COPIAR_PYE()
            objDatos.SP_GT_TBL_GENERICA()

            'MsgBox("sps ejecutados")



            'Validacion Dia Habil o No Habil
            Dim Tipodias As String = objDatos.sp_GT_VALIDARDIA()

            If Tipodias = "habil" Then
                Sincronizacion(resultadoCarga)
            Else
                EnviarNotificacion_Movimientos_Tablas("No se ejecuto la tarea, Dia No Habil : " & Date.Now.AddDays(-1).ToString)
            End If

            'objTarea.LogFechaFin()
            'If objTarea.EnviarNotificaciones Then
            '    If ValidarErrores() Thenmsg
            '        EnviarNotificacion()
            '    End If
            'End If

        Catch ex As Exception
            Me.Close()
            'EnviarNotificacion_Movimientos_Tablas(ex.Message)
        Finally
            Me.Close()
        End Try

        Me.Close()

    End Sub

    Private Sub Sincronizacion(ByRef ResultadoCarga As Short)

        Dim dsOrigen As DataSet
        objTarea.BaseDatos = "SQL"
        Dim Plano As String
        Dim Resultado As String = ""
        Dim XMLCreacionTercero As String = ""
        Dim idTareaValido As Boolean = True
        Dim PlanoGenerado As String


        'Id de la tarea enviado como parametro
        Dim args As String() = Environment.GetCommandLineArgs()

        If args.Length = 2 Then
            If IsNumeric(args(1)) Then
                objTarea.Tarea = args(1)
                objTarea.LogInicio()
            Else
                objTarea.LogRecuperacionDatosOrigen(0)
                objTarea.LogMensajesError("El argumento enviado debe ser numérico, verifique la configuración de la tarea programada y configure como argumento el Id de la tarea")
                Exit Sub
            End If
        Else
            objTarea.LogRecuperacionDatosOrigen(0)
            objTarea.LogMensajesError("No se envió el Id de la tarea como argumento al programa, verifique la configuración de la tarea programada y configure como argumento el Id de la tarea")
            Exit Sub
        End If

        'Obtener los datos fuentes del cliente para la tarea X
        Try
            objTarea.LogFechaInicioRecuperacionDatosOrigen()
            dsOrigen = objTarea.DatosOrigen(idTareaValido)
            If idTareaValido Then
                objTarea.LogFechaFinRecuperacionDatosOrigen()
                objTarea.LogRecuperacionDatosOrigen(1)
            Else
                objTarea.LogRecuperacionDatosOrigen(0)
                objTarea.LogMensajesError("No se encontró el Id de tarea: " & objTarea.Tarea & " en la base de datos, verifique la configuración de la tarea programada de Windows")
                Exit Sub
            End If
        Catch ex As Exception
            objTarea.LogRecuperacionDatosOrigen(0)
            objTarea.LogMensajesError(ex.Message)
            Exit Sub
        End Try

        If dsOrigen.Tables.Count >= 1 Then
            If dsOrigen.Tables(0).Rows.Count >= 1 Then

                'Invocar Web Service de GT para Obtener el plano
                Try
                    objTarea.LogFechaInicioGeneracionPlano()
                    Dim objGenericTransfer As New wsGenerarPlano.wsGenerarPlano
                    PlanoGenerado = objTarea.RutaGeneracionPlano

                    'Verificacion de informacion de encabezado
                    If dsOrigen.Tables.Count > 0 Then
                        If dsOrigen.Tables(0).Rows.Count > 0 Then

                            Plano = objGenericTransfer.GenerarPlanoXML(objTarea.IdDocumento,
                                                            objTarea.NombreDocumento,
                                                            2, "GyF", "gt", "gt",
                                                            dsOrigen, PlanoGenerado)

                        End If
                    End If


                    'If Resultado <> "Se genero el plano correctamente" Then
                    '    objTarea.LogGeneracionDePlano(0)
                    '    objTarea.LogMensajesError(Resultado)
                    '    Exit Sub
                    'Else
                    objTarea.LogFechaFinGeneracionPlano()
                    objTarea.LogGeneracionDePlano(1)
                    'End If



                Catch ex As Exception
                    objTarea.LogGeneracionDePlano(0)
                    objTarea.LogMensajesError(ex.Message)
                    Exit Sub
                End Try

                'Invocar el Web Service De Siesa
                If objTarea.InvocarWebService Then
                    Try
                        XMLCreacionTercero &= "<?xml version='1.0' encoding='utf-8'?>"
                        XMLCreacionTercero &= "<Importar>"
                        XMLCreacionTercero &= "<NombreConexion>" & objTarea.NombreConexionWsSiesa & "</NombreConexion>"
                        XMLCreacionTercero &= "<IdCia>" & objTarea.CiaWsSiesa & "</IdCia>"
                        XMLCreacionTercero &= "<Usuario>" & objTarea.UsuarioWsSiesa & "</Usuario>"
                        XMLCreacionTercero &= "<Clave>" & objTarea.ClaveWsSiesa & "</Clave> "
                        XMLCreacionTercero &= Plano
                        XMLCreacionTercero &= "</Importar>"
                        'AlmacenarXML(PlanoGenerado.ToString.Replace(".txt", ".xml"), XMLCreacionTercero)

                        Dim objUnoEE As New wsUnoEE.WSUNOEE
                        'Se define un tiempo de espera de 15 minutos: 15 * 60 * 1000
                        objUnoEE.Timeout = 900000

                        Dim dsResultado As DataSet
                        Dim UnoEEResultado As Short
                        objTarea.LogFechaInicioWebServiceSiesa()
                        dsResultado = objUnoEE.ImportarXML(XMLCreacionTercero, UnoEEResultado)
                        ResultadoCarga = UnoEEResultado

                        If UnoEEResultado = 0 Then
                            objTarea.LogWebServiceSiesa(1)
                            objTarea.LogFechaFinWebServiceSiesa()
                            objDatos.SP_GT_ACTUALIZAR_TRANSACTION_CV_HABILES("1")
                            EnviarNotificacion_Movimientos_Tablas("Se ejecuto la tarea de carga al sistema Siesa correctamente")
                        Else
                            objDatos.SP_GT_ACTUALIZAR_TRANSACTION_CV_HABILES("0")
                            Dim ErrorSiesa As String = ""
                            For Each FilaError As System.Data.DataRow In dsResultado.Tables(0).Rows
                                ErrorSiesa &= "Fila: " & FilaError.Item(0).ToString & " Tipo de Registro:" & FilaError.Item(1).ToString & " SubTipo de resgistro:" & FilaError.Item(2).ToString & " Version:" & FilaError.Item(3).ToString & " Nivel" & FilaError.Item(4).ToString & " Valor" & FilaError.Item(5).ToString & " Detalle:" & FilaError.Item(6).ToString & vbCrLf
                            Next
                            objTarea.LogMensajesError(ErrorSiesa)
                            objTarea.LogWebServiceSiesa(0)
                            EnviarNotificacion()
                            Exit Sub
                        End If
                    Catch ex As Exception
                        objTarea.LogWebServiceSiesa(0)
                        objTarea.LogMensajesError(ex.Message)
                        EnviarNotificacion()
                        Exit Sub
                    End Try

                Else
                    'objDatos.SP_GT_ACTUALIZAR_TRANSACTION_CV_HABILES("1")
                    objDatos.SP_GT_MARCAR_PROCESOTERMINADO()
                End If

            Else
                ResultadoCarga = "-1"
                EnviarNotificacion_Movimientos_Tablas("Al ejecutar la tarea no se encontraron datos para enviar a Siesa")
            End If
        Else
            ResultadoCarga = "-1"
            EnviarNotificacion_Movimientos_Tablas("Al ejecutar la tarea no se encontraron datos para enviar a Siesa")
        End If

        'objDatos.SP_GT_MARCAR_PROCESOTERMINADO()
        'objTarea.LogEjecucionCompleta()
    End Sub

    Private Sub AlmacenarXML(ByVal path As String, ByVal XML As String)
        Dim fs As FileStream = File.Create(path)
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(XML)
        fs.Write(info, 0, info.Length)
        fs.Close()
    End Sub

    ''' <summary>
    ''' Valida si ocurrieron errores al importar los datos a Siesa Enterprise
    ''' </summary>
    ''' <returns> Valor booleano: True -> Ocurrieron errores False -> No ocurrieron errores</returns>
    ''' <remarks></remarks>
    Private Function ValidarErrores() As Boolean
        Dim dtResultado As DataTable = objTarea.ResultadoOperacion()

        If dtResultado.Rows(0).Item("Ejecución completa").ToString = "No" Then
            Return True
        ElseIf dtResultado.Rows(0).Item("Mensaje de error").ToString <> "" Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub EnviarNotificacion_Movimientos_Tablas(ByVal mensajeError As String)

        Dim strMensaje As New StringBuilder
        Dim strAsunto As String

        strAsunto = "GTIntegration Ejecucion de Tarea:  "

        strMensaje.AppendLine("<style type=""text/css"">")
        strMensaje.AppendLine("</style>")
        strMensaje.AppendLine("<BR/><BR/><table style=""color: #666; font-family: 'font-family: 'Georgia'', sans-serif;background-image:url('http://www.generictransfer.com/imagenes/main_bg.png')"" border=""0"" cellpadding=""0"" cellspacing=""0"" ")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td style=""border-bottom-style: solid; border-bottom-width: medium; border-bottom-color: #2AA0D0"" bgcolor=""White"">")
        strMensaje.AppendLine("<img alt="""" src=""http://www.generictransfer.com/imagenes/logo_ge.png"" />")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td width=""700px"">")
        strMensaje.AppendLine("<br />")
        strMensaje.AppendLine("Generic Transfer Integration")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td style=""font-size: 18px;font-style: normal;font-weight: normal;font-variant: normal;text-decoration: none;color: #4799CC"">")
        strMensaje.AppendLine("<BR/>Información de la tarea " & mensajeError & " <BR/>")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("&nbsp;")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<br />")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<br />")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")
        strMensaje.AppendLine("<tr>")

        strMensaje.AppendLine("<br />")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")

        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")
        strMensaje.AppendLine("<tr>")

        strMensaje.AppendLine("<br />")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td align=""center"">")
        strMensaje.AppendLine("<strong><span class=""style2"">")
        strMensaje.AppendLine("<BR/><BR/>Interfaces y Soluciones S.A.S<br /> <BR/><BR/>")
        strMensaje.AppendLine("</span></strong>")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")
        strMensaje.AppendLine("</table>")

        Dim objMail As New clsCorreo
        objMail.EnviarEmail("Generic Transfer Integration GYF", My.Settings.MailUser, objTarea.Destinatarios, strAsunto, strMensaje.ToString, "")
    End Sub

    Private Sub EnviarNotificacion()
        Dim strMensaje As New StringBuilder
        Dim dtResultado As DataTable = objTarea.ResultadoOperacion()
        Dim strAsunto As String

        If dtResultado.Rows(0).Item("Mensaje de error").ToString = "" Then
            strAsunto = "GT Integration, Tarea: " & dtResultado.Rows(0).Item("Nombre").ToString & ", Completa: " & dtResultado.Rows(0).Item("Ejecución completa").ToString
        Else
            strAsunto = "GT Integration, Tarea: " & dtResultado.Rows(0).Item("Nombre").ToString & ", Completa: " & dtResultado.Rows(0).Item("Ejecución completa").ToString & ", con errores"
        End If

        strMensaje.AppendLine("<style type=""text/css"">")
        strMensaje.AppendLine("</style>")
        strMensaje.AppendLine("<BR/><BR/><table style=""color: #666; font-family: 'font-family: 'Georgia'', sans-serif;background-image:url('http://www.generictransfer.com/imagenes/main_bg.png')"" border=""0"" cellpadding=""0"" cellspacing=""0"" ")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td style=""border-bottom-style: solid; border-bottom-width: medium; border-bottom-color: #2AA0D0"" bgcolor=""White"">")
        strMensaje.AppendLine("<img alt="""" src=""http://www.generictransfer.com/imagenes/logo_ge.png"" />")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td width=""700px"">")
        strMensaje.AppendLine("<br />")
        strMensaje.AppendLine("Generic Transfer Integration")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td style=""font-size: 18px;font-style: normal;font-weight: normal;font-variant: normal;text-decoration: none;color: #4799CC"">")
        strMensaje.AppendLine("<BR/>Información de la tarea " & dtResultado.Rows(0).Item("Id Tarea").ToString & " - " & dtResultado.Rows(0).Item("Nombre").ToString & " <BR/>")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("&nbsp;")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        If dtResultado.Rows(0).Item("Mensaje de error").ToString = "" Then
            strMensaje.AppendLine("Ejecución completa:" & dtResultado.Rows(0).Item("Ejecución completa").ToString)
        Else
            strMensaje.AppendLine("Ejecución completa:" & dtResultado.Rows(0).Item("Ejecución completa").ToString & ", con errores")
        End If
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("Mensajes de Siesa Enterprise:" & dtResultado.Rows(0).Item("Mensaje de error").ToString)
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<br />")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("Inicio:" & dtResultado.Rows(0).Item("Fecha Inicio").ToString)
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("Fin:" & dtResultado.Rows(0).Item("Fecha Fin").ToString)
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<br />")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("Inicio recuperación de datos Origen:" & dtResultado.Rows(0).Item("Fecha inicio de recuperacion de datos Origen").ToString)
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("Fin recuperación de datos Origen:" & dtResultado.Rows(0).Item("Fecha fin de recuperacion de datos Origen").ToString)
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("Recuperación de datos origen:" & dtResultado.Rows(0).Item("Recuperacion de datos origen").ToString)
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")
        strMensaje.AppendLine("<tr>")

        strMensaje.AppendLine("<br />")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("Inicio de generación de plano:" & dtResultado.Rows(0).Item("Fecha inicio de generacion de plano").ToString)
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("Fin de generación de plano:" & dtResultado.Rows(0).Item("Fecha fin de generacion de plano").ToString)
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td>")
        strMensaje.AppendLine("Generación de plano:" & dtResultado.Rows(0).Item("Generacion de plano").ToString)
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")
        strMensaje.AppendLine("<tr>")


        If dtResultado.Rows(0).Item("InvocarWebService").ToString = "True" Then
            strMensaje.AppendLine("<br />")

            strMensaje.AppendLine("<tr>")
            strMensaje.AppendLine("<td>")
            strMensaje.AppendLine("Inicio Web Service:" & dtResultado.Rows(0).Item("Fecha inicio Web Service").ToString)
            strMensaje.AppendLine("</td>")
            strMensaje.AppendLine("</tr>")

            strMensaje.AppendLine("<tr>")
            strMensaje.AppendLine("<td>")
            strMensaje.AppendLine("Fin Web Service:" & dtResultado.Rows(0).Item("Fecha fin Web Service").ToString)
            strMensaje.AppendLine("</td>")
            strMensaje.AppendLine("</tr>")

            strMensaje.AppendLine("<tr>")
            strMensaje.AppendLine("<td>")
            strMensaje.AppendLine("Web Service de Siesa:" & dtResultado.Rows(0).Item("Consumir el Web Service de Siesa").ToString)
            strMensaje.AppendLine("</td>")
            strMensaje.AppendLine("</tr>")
        End If

        strMensaje.AppendLine("<br />")

        strMensaje.AppendLine("<tr>")
        strMensaje.AppendLine("<td align=""center"">")
        strMensaje.AppendLine("<strong><span class=""style2"">")
        strMensaje.AppendLine("<BR/><BR/>Interfaces y Soluciones S.A.S<br /> <BR/><BR/>")
        strMensaje.AppendLine("</span></strong>")
        strMensaje.AppendLine("</td>")
        strMensaje.AppendLine("</tr>")
        strMensaje.AppendLine("</table>")

        Dim objMail As New clsCorreo
        objMail.EnviarEmail("GTIntegration Giros y Finanzas", My.Settings.MailUser, objTarea.Destinatarios, strAsunto, strMensaje.ToString, "")
    End Sub

    Private Sub AlmacenarLog(ByVal Mensaje As String)
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("C:\Users\interfaces\Documents\Log.txt", True)
        file.WriteLine(Mensaje)
        file.Close()
    End Sub

End Class