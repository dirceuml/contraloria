Module FuncionesVarias

    Function EjecutaProcesoCarga(ByVal package_name As String, ByRef log As String, _
                                 Optional ByVal parametro1 As String = "", Optional ByVal parametro2 As String = "", Optional ByVal parametro3 As String = "", Optional ByVal parametro4 As String = "")

        Dim app As Microsoft.SqlServer.Dts.Runtime.Application
        Dim package As Microsoft.SqlServer.Dts.Runtime.Package
        Dim results As Microsoft.SqlServer.Dts.Runtime.DTSExecResult
        Dim path As String
        Dim flag As Boolean

        flag = True
        path = "E:\ATV\ATVContraloriaIS\bin\"

        Try
            app = New Microsoft.SqlServer.Dts.Runtime.Application
            package = app.LoadPackage(path & package_name, Nothing)
            If package_name = "Facturacion.dtsx" Then
                package.Variables("IdPeriodoCarga").Value = parametro1
            ElseIf package_name = "Rating.dtsx" Then
                package.Variables("IdPeriodo").Value = parametro1
            ElseIf package_name = "Ventas.dtsx" Then
                package.Variables("FechaIni").Value = parametro1
                package.Variables("FechaFin").Value = parametro2
            ElseIf package_name = "LibroInventario.dtsx" Then
                package.Variables("IdPeriodo").Value = parametro1
                package.Variables("Fecha").Value = parametro2
            End If

            results = package.Execute()

            If results = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure Then
                Dim dtserror As Microsoft.SqlServer.Dts.Runtime.DtsError
                For Each dtserror In package.Errors
                    log = log & package_name & ": " & dtserror.Description & "<br>"
                Next
                flag = False
            End If
        Catch ex As Exception
            log = log & package_name & ": " & ex.Message & "<br>"
            flag = False
        End Try

        Return flag
    End Function

    Sub GrabaLog(ByVal IdUsuario As Integer, ByVal Logs As String)
        Dim sql As String
        Dim cn As System.Data.SqlClient.SqlConnection
        Dim cmd As System.Data.SqlClient.SqlCommand

        cn = New System.Data.SqlClient.SqlConnection()
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn2").ConnectionString
        cn.Open()

        sql = "insert into Logs(IdUsuario, Logs, Fecha) values (@IdUsuario, @Logs, getdate())"
        cmd = New System.Data.SqlClient.SqlCommand(sql, cn)
        cmd.Parameters.AddWithValue("@IdUsuario", IdUsuario)
        cmd.Parameters.AddWithValue("@Logs", Logs)
        cmd.ExecuteNonQuery()

        cn.Close()
    End Sub

    Sub GeneraArchivoDescarga(ByVal Response As System.Web.HttpResponse, ByVal RutaArchivo As String, ByVal NombreArchivo As String)
        Dim fileStream As System.IO.FileStream
        fileStream = New System.IO.FileStream(RutaArchivo, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        Dim bytes() As Byte
        ReDim bytes(fileStream.Length)
        fileStream.Read(bytes, 0, fileStream.Length)
        fileStream.Close()

        Response.Buffer = True
        Response.Clear()
        Response.ContentType = "application/octet-stream;"
        'This header is for saving it as an Attachment and popup window should display to to offer save as or open a PDF file 
        Response.AddHeader("Content-Disposition", "attachment; filename=" + NombreArchivo)
        'Response.AddHeader("content-disposition", "inline; filename=" + Archivo);
        Response.BinaryWrite(bytes)
        Response.Flush()
        Response.End()
    End Sub

    Sub DescargaExcel(ByVal Response As System.Web.HttpResponse, ByVal grid As System.Web.UI.WebControls.GridView, ByVal NombreArchivo As String)
        Dim sw As New System.IO.StringWriter()
        Dim htw As New HtmlTextWriter(sw)
        Dim frm As New System.Web.UI.HtmlControls.HtmlForm()

        grid.Parent.Controls.Add(frm)
        frm.Attributes("runat") = "server"
        frm.Controls.Add(grid)
        frm.RenderControl(htw)

        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/vnd.ms-excel"
        Response.AddHeader("Content-Disposition", "attachment;filename=" & NombreArchivo)
        Response.Charset = "UTF-8"
        Response.ContentEncoding = System.Text.Encoding.[Default]
        Response.Write(sw.ToString())
        Response.[End]()
    End Sub
End Module
