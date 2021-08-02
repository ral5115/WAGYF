Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Net

Public Class clsDatos

    Public Sub SP_GT_COPIAR_TRANSACTION_CV()

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter
        Dim ds As New DataSet

        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "SP_GT_COPIAR_TRANSACTION_CV"

        Try

            sqlAdaptador.SelectCommand = sqlComando
            sqlAdaptador.SelectCommand.CommandTimeout = 99999
            sqlConexion.Open()
            sqlComando.ExecuteNonQuery()

        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Sub

    Public Sub sp_GT_LIMPIAR_TRANSACTION_CV_HABILES()

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter
        Dim ds As New DataSet

        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "sp_GT_LIMPIAR_TRANSACTION_CV_HABILES"

        Try

            sqlAdaptador.SelectCommand = sqlComando
            sqlAdaptador.SelectCommand.CommandTimeout = 99999
            sqlConexion.Open()
            sqlComando.ExecuteNonQuery()

        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Sub

    Public Function sp_GT_VALIDARDIA() As String

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter
        Dim ds As New DataSet

        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "sp_GT_VALIDARDIA"


        Dim dtfecha As Date = Date.Now.AddDays(-1)
        Dim strfecha As String = dtfecha.ToString("yyyyMMdd")

        sqlComando.Parameters.AddWithValue("@fecha", strfecha)

        Try

            sqlAdaptador.SelectCommand = sqlComando
            sqlAdaptador.Fill(ds)

            If ds.Tables.Count > 0 Then
                If ds.Tables(0).Rows.Count > 0 Then

                    If ds.Tables(0).Rows(0).Item("tipoDia") = 1 Then
                        Return "no habil"
                    Else
                        Return "habil"
                    End If

                Else
                    Return "La fecha: " & strfecha & " no existe"
                End If
            Else
                Return "La fecha: " & strfecha & " no existe"
            End If

        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Function

    Public Sub SP_GT_ACTUALIZAR_TRANSACTION_CV_HABILES(ByVal ID_TRANSACTION As String)

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter

        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "SP_GT_ACTUALIZAR_TRANSACTION_CV_HABILES"

        sqlComando.Parameters.AddWithValue("@ID_TRANSACTION", ID_TRANSACTION)

        Try

            sqlAdaptador.SelectCommand = sqlComando
            sqlComando.Connection.Open()
            sqlComando.ExecuteNonQuery()

        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Sub

    Public Sub SP_GT_COPIAR_PYE()

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter
        Dim ds As New DataSet

        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "SP_GT_COPIAR_PYE"

        Try

            sqlAdaptador.SelectCommand = sqlComando
            sqlAdaptador.SelectCommand.CommandTimeout = 99999
            sqlConexion.Open()
            sqlComando.ExecuteNonQuery()

        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Sub

    Public Sub SP_GT_TBL_GENERICA()

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter
        Dim ds As New DataSet

        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "SP_GT_TBL_GENERICA"

        Try

            sqlAdaptador.SelectCommand = sqlComando
            sqlAdaptador.SelectCommand.CommandTimeout = 99999
            sqlConexion.Open()
            sqlComando.ExecuteNonQuery()

        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Sub

    Public Sub SP_GT_MARCAR_PROCESOTERMINADO()

        Dim sqlConexion As New SqlConnection(My.Settings.strConexionGT)
        Dim sqlComando As SqlCommand = New SqlCommand
        Dim sqlAdaptador As SqlDataAdapter = New SqlDataAdapter
        Dim ds As New DataSet

        sqlComando.Connection = sqlConexion
        sqlComando.CommandType = CommandType.StoredProcedure
        sqlComando.CommandText = "SP_GT_MARCAR_PROCESOTERMINADO"

        Try

            sqlAdaptador.SelectCommand = sqlComando
            sqlAdaptador.SelectCommand.CommandTimeout = 99999
            sqlConexion.Open()
            sqlComando.ExecuteNonQuery()

        Catch ex As Exception
            Throw ex
        Finally
            sqlComando.Parameters.Clear()
            sqlComando.Connection.Close()
        End Try

    End Sub

End Class
