'Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration
Imports System.Data.SqlClient

Public Class Funciones
    Public Const iREGxPAG As Integer = 20

    Public Shared Function Hoy() As String
        Dim sDia As String = Date.Today.Day.ToString()
        Dim sMes As String = Date.Today.Month.ToString()
        If sDia.Length = 1 Then
            sDia = "0" + sDia
        End If
        If sMes.Length = 1 Then
            sMes = "0" + sMes
        End If
        Return Date.Today.Year.ToString() + sMes + sDia
    End Function

    Public Shared Sub EjecutarSentencia(ByVal sSql As String)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        Try
            cn.Open()
            cmd = New OleDbCommand(sSql, cn)
            cmd.ExecuteNonQuery()
        Catch ee As SqlException
            'lblMessage.Text = ee.Message
        Finally
            cmd.Dispose()
            cn.Close()
            cn.Dispose()
        End Try
    End Sub

    Public Shared Function ExtraerValor(ByVal sSql As String, ByVal sCampo As String) As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As New DataSet
        Dim rdr As OleDbDataReader
        Dim sValor As String

        sValor = ""
        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        Try
            cn.Open()
            cmd = New OleDbCommand(sSql, cn)
            rdr = cmd.ExecuteReader
            If rdr.Read() Then
                sValor = rdr(sCampo).ToString()
            End If
            rdr.Close()
        Catch ee As SqlException
            sValor = ee.Message
        Finally
            cmd.Dispose()
            cn.Close()
            cn.Dispose()
        End Try
        Return sValor
    End Function

    Public Shared Function FechaSQL(ByVal sFecha As String) As String
        '01/01/2011  11:00
        Dim sDia As String = sFecha.Substring(0, 2)
        Dim sMes As String = sFecha.Substring(3, 2)
        Dim sAnio As String = sFecha.Substring(6, 4)

        If sAnio = "----" Then
            Return "0"
        Else
            Return sAnio + sMes + sDia
        End If
        'return 20110101
    End Function

    Public Shared Function QuitaComa(ByVal sValor As String) As String
        Return sValor.Replace(",", "")
    End Function

    Public Shared Function ValidarImporte(ByVal sValor As String) As String
        Dim sImporte As String
        sImporte = QuitaComa(sValor)
        If sImporte = "" Then
            sImporte = 0
        End If
        Return sImporte
    End Function

    Public Shared Function FormatoDinero(ByVal sValor As String) As String
        Dim dCantidad As Decimal
        dCantidad = Decimal.Parse(sValor)
        Return String.Format("{0:c}", dCantidad)
    End Function

    Public Shared Sub IngresarMovimiento2(ByVal sIdMovimiento As String, ByVal sCodCuentaFC2 As String, ByVal sCodArea2 As String)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As New DataSet
        Dim rdr As OleDbDataReader
        Dim sSql, sIdPeriodo, sCodCuentaBanco, sNroVoucher, sNroItem, sCodCuentaFC, sCodArea, sObservaciones As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()
        sSql = "SELECT IdPeriodo, CodCuentaBanco, NroVoucher, NroItem, CodCuentaFC, CodArea, Observaciones FROM Movimiento Where idMovimiento = " + sIdMovimiento
        cmd = New OleDbCommand(sSql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() Then
            sIdPeriodo = rdr("IdPeriodo").ToString()
            sCodCuentaBanco = rdr("CodCuentaBanco").ToString()
            sNroVoucher = rdr("NroVoucher").ToString()
            sNroItem = rdr("NroItem").ToString()
            sCodCuentaFC = rdr("CodCuentaFC").ToString()
            sCodArea = rdr("CodArea").ToString()
            sObservaciones = rdr("Observaciones").ToString()
            'ELIMINANDO EL REGISTRO EN CASO EXISTA
            sSql = " Delete From Movimiento2 Where IdPeriodo = " + sIdPeriodo + " "
            sSql += "AND CodCuentaBanco = " + sCodCuentaBanco + " "
            sSql += "AND NroVoucher = " + sNroVoucher + " "
            sSql += "AND NroItem = " + sNroItem + " "
            sSql += "AND CodCuentaFC = " + sCodCuentaFC + " "
            sSql += "AND Observaciones = '" + sObservaciones + "' "
            EjecutarSentencia(sSql)
            If sCodArea2 <> "-1" Or sCodCuentaFC2 <> "-1" Then
                If sCodArea2 = "-1" Then
                    sCodArea2 = sCodArea
                End If
                If sCodCuentaFC2 = "-1" Then
                    sCodCuentaFC2 = sCodCuentaFC
                End If
                'INSERTANDO EL REGISTRO EN CASO EXISTA
                sSql = " Insert into Movimiento2 (IdPeriodo, CodCuentaBanco, NroVoucher, "
                sSql += "NroItem, CodCuentaFC, Observaciones, CodCuentaFC2, CodArea2) "
                sSql += "Values (" + sIdPeriodo + ", " + sCodCuentaBanco + ", " + sNroVoucher + ", "
                sSql += "" + sNroItem + ", " + sCodCuentaFC + ", '" + sObservaciones + "', " + sCodCuentaFC2 + ", " + sCodArea2 + ") "
                EjecutarSentencia(sSql)
            End If
        End If
        rdr.Close()
        cn.Close()
    End Sub

    Public Shared Sub IngresarMovimientoN(ByVal sIdMovimiento As String, ByVal sCodCuentaFCN As String, ByVal sCodAreaN As String)
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim dtaset As New DataSet
        Dim rdr As OleDbDataReader
        Dim sSql, sIdPeriodo, sCodCuentaBanco, sNroVoucher, sNroItem, sCodCuentaFC, sCodArea, sObservaciones As String

        cn.ConnectionString = ConfigurationManager.ConnectionStrings("cn").ConnectionString
        cn.Open()
        sSql = "SELECT IdPeriodo, CodCuentaBanco, NroVoucher, NroItem, CodCuentaFC, CodArea, Observaciones FROM Movimiento Where idMovimiento = " + sIdMovimiento
        cmd = New OleDbCommand(sSql, cn)
        rdr = cmd.ExecuteReader
        If rdr.Read() Then
            sIdPeriodo = rdr("IdPeriodo").ToString()
            sCodCuentaBanco = rdr("CodCuentaBanco").ToString()
            sNroVoucher = rdr("NroVoucher").ToString()
            sNroItem = rdr("NroItem").ToString()
            sCodCuentaFC = rdr("CodCuentaFC").ToString()
            sCodArea = rdr("CodArea").ToString()
            sObservaciones = rdr("Observaciones").ToString()
            'ELIMINANDO EL REGISTRO EN CASO EXISTA
            sSql = " Delete From MovimientoN Where IdPeriodo = " + sIdPeriodo + " "
            sSql += "AND CodCuentaBanco = " + sCodCuentaBanco + " "
            sSql += "AND NroVoucher = " + sNroVoucher + " "
            sSql += "AND NroItem = " + sNroItem + " "
            sSql += "AND CodCuentaFC = " + sCodCuentaFC + " "
            sSql += "AND Observaciones = '" + sObservaciones + "' "
            EjecutarSentencia(sSql)
            If sCodAreaN <> "-1" Or sCodCuentaFCN <> "-1" Then
                If sCodAreaN = "-1" Then
                    sCodAreaN = sCodArea
                End If
                If sCodCuentaFCN = "-1" Then
                    sCodCuentaFCN = sCodCuentaFC
                End If
                'INSERTANDO EL REGISTRO EN CASO EXISTA
                sSql = " Insert into MovimientoN (IdPeriodo, CodCuentaBanco, NroVoucher, "
                sSql += "NroItem, CodCuentaFC, Observaciones, CodCuentaFCN, CodAreaN) "
                sSql += "Values (" + sIdPeriodo + ", " + sCodCuentaBanco + ", " + sNroVoucher + ", "
                sSql += "" + sNroItem + ", " + sCodCuentaFC + ", '" + sObservaciones + "', " + sCodCuentaFCN + ", " + sCodAreaN + ") "
                EjecutarSentencia(sSql)
            End If
        End If
        rdr.Close()
        cn.Close()
    End Sub

    Public Shared Sub IngresarCboPersonaN(ByVal sCodPersona As String, ByVal sCodPersonaN As String)
        Dim sSql, sPersonaN As String

        sSql = "SELECT DISTINCT PersonaN FROM PersonaN Where CodPersonaN = " + sCodPersonaN
        sPersonaN = ExtraerValor(sSql, "PersonaN")

        'ELIMINANDO EL REGISTRO EN CASO EXISTA
        sSql = " Delete From PersonaN Where CodPersona = " + sCodPersona + " "
        EjecutarSentencia(sSql)

        'INSERTANDO EL REGISTRO EN CASO EXISTA
        sSql = " Insert into PersonaN (CodPersonaN, PersonaN, CodPersona) "
        sSql += "Values (" + sCodPersonaN + ", '" + sPersonaN + "', " + sCodPersona + ") "
        EjecutarSentencia(sSql)
    End Sub

    Public Shared Sub IngresarTxtPersonaN(ByVal sCodPersona As String, ByVal sPersonaN As String)
        Dim sSql, sCodPersonaN As String

        'ELIMINANDO EL REGISTRO EN CASO EXISTA
        sSql = " Delete From PersonaN Where CodPersona = " + sCodPersona + " "
        EjecutarSentencia(sSql)
        If sPersonaN <> "" Then
            'CONSULTANDO POR EL SIGUIENTE REGISTRO
            sSql = "SELECT (max(CodPersonaN) + 1) as CodPersonaN FROM PersonaN "
            sCodPersonaN = ExtraerValor(sSql, "CodPersonaN")
            If (sCodPersonaN = "1") Then
                sCodPersonaN = "90001"
            End If
            'INSERTANDO EL REGISTRO EN CASO EXISTA
            sSql = " Insert into PersonaN (CodPersonaN, PersonaN, CodPersona) "
            sSql += "Values (" + sCodPersonaN + ", '" + sPersonaN + "', " + sCodPersona + ") "
            EjecutarSentencia(sSql)
        End If
    End Sub

    Public Shared Sub ActualizarArea(ByVal sCodArea As String, ByVal sFlagPrograma As String)
        Dim sSql As String

        'ACTUALIZANDO EL REGISTRO EN CASO EXISTA
        sSql = " Update Area Set FlagPrograma = '" + sFlagPrograma + "' "
        sSql += "Where CodArea = " + sCodArea
        EjecutarSentencia(sSql)
    End Sub

    Public Shared Function ActualizarCuentaFC2(ByVal sCodCuenta2 As String, ByVal sCuenta2 As String, ByVal sCodEstado As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM V_CuentaFC WHERE CodCuenta2 <> " + sCodCuenta2 + " AND Cuenta2 ='" + sCuenta2 + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'ACTUALIZANDO EL REGISTRO EN CASO EXISTA
            sSql = " Update CuentaFC2 Set Cuenta2 = '" + sCuenta2 + "', "
            sSql += "CodEstado = " + sCodEstado + " "
            sSql += "Where CodCuenta2 = " + sCodCuenta2
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "La cuenta " + sCuenta2 + " ya se encuentra registrada"
    End Function

    Public Shared Function InsertarCuentaFC2(ByVal sCuenta2 As String, ByVal sCodEstado As String) As String
        Dim sSql, sNumReg, sCodCuenta2 As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM V_CuentaFC WHERE Cuenta2 ='" + sCuenta2 + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'OBTENIENDO EL SIGUIENTE IDENTIFICADOR
            sSql = "SELECT (max(CodCuenta2) + 1) as CodCuenta2 FROM CuentaFC2 "
            sCodCuenta2 = ExtraerValor(sSql, "CodCuenta2")
            If (sCodCuenta2 = "1") Then
                sCodCuenta2 = "901"
            End If
            'INSERTANDO EL REGISTRO EN CASO EXISTA
            sSql = " Insert into CuentaFC2 (CodCuenta2, Cuenta2, CodEstado) "
            sSql += "Values (" + sCodCuenta2 + ", '" + sCuenta2 + "', " + sCodEstado + ") "
            EjecutarSentencia(sSql)
            Return "Nuevo registro grabado: " + sCuenta2
        End If
        Return "La cuenta " + sCuenta2 + " ya se encuentra registrada"
    End Function

    Public Shared Function BorrarCuentaFC2(ByVal sCodCuenta2) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CuentaFC2 Where CodCuenta2 = " + sCodCuenta2
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarCuentaFCN(ByVal sCodCuentaN As String, ByVal sCuentaN As String, ByVal sCodEstado As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM V_CuentaFCN WHERE CodCuentaN <> " + sCodCuentaN + " AND CuentaN ='" + sCuentaN + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'ACTUALIZANDO EL REGISTRO EN CASO EXISTA
            sSql = " Update CuentaFCN Set CuentaN = '" + sCuentaN + "', "
            sSql += "CodEstado = " + sCodEstado + " "
            sSql += "Where CodCuentaN = " + sCodCuentaN
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "La cuenta " + sCuentaN + " ya se encuentra registrada"
    End Function

    Public Shared Function InsertarCuentaFCN(ByVal sCuentaN As String, ByVal sCodEstado As String) As String
        Dim sSql, sNumReg, sCodCuentaN As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM V_CuentaFCN WHERE CuentaN ='" + sCuentaN + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'OBTENIENDO EL SIGUIENTE IDENTIFICADOR
            sSql = "SELECT (max(CodCuentaN) + 1) as CodCuentaN FROM CuentaFCN "
            sCodCuentaN = ExtraerValor(sSql, "CodCuentaN")
            If (sCodCuentaN = "1") Then
                sCodCuentaN = "901"
            End If
            'INSERTANDO EL REGISTRO EN CASO EXISTA
            sSql = " Insert into CuentaFCN (CodCuentaN, CuentaN, CodEstado) "
            sSql += "Values (" + sCodCuentaN + ", '" + sCuentaN + "', " + sCodEstado + ") "
            EjecutarSentencia(sSql)
            Return "Nuevo registro grabado: " + sCuentaN
        End If
        Return "La cuenta " + sCuentaN + " ya se encuentra registrada"
    End Function

    Public Shared Function BorrarCuentaFCN(ByVal sCodCuentaN) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CuentaFCN Where CodCuentaN = " + sCodCuentaN
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarCtaFlujoNRep(ByVal sIdCtaFlujoNRep As String, ByVal sCtaFlujoNRep As String, ByVal sSigno As String, ByVal sCodSeccion As String, ByVal sFlagModif As String, ByVal sCodAuxiliar As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaFlujoNRep WHERE CtaFlujoNRep = '" + sCtaFlujoNRep + "' AND IdCtaFlujoNRep <> " + sIdCtaFlujoNRep
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'ACTUALIZANDO EL REGISTRO
            sSql = " Update CtaFlujoNRep Set CtaFlujoNRep = '" + sCtaFlujoNRep + "', "
            sSql += "Signo = '" + sSigno + "', "
            sSql += "CodSeccion = '" + sCodSeccion + "', "
            sSql += "FlagModif = '" + sFlagModif + "', "
            sSql += "CodAuxiliar = '" + sCodAuxiliar + "', "
            sSql += "Orden = " + sOrden + " "
            sSql += "Where IdCtaFlujoNRep = " + sIdCtaFlujoNRep
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function InsertarCtaFlujoNRep(ByVal sCtaFlujoNRep As String, ByVal sSigno As String, ByVal sCodSeccion As String, ByVal sFlagModif As String, ByVal sCodAuxiliar As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaFlujoNRep WHERE CtaFlujoNRep = '" + sCtaFlujoNRep + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'INSERTANDO EL REGISTRO
            sSql = " Insert into CtaFlujoNRep (CtaFlujoNRep, Signo, CodSeccion, FlagModif, CodAuxiliar, Orden) "
            sSql += "Values ('" + sCtaFlujoNRep + "', '" + sSigno + "', '" + sCodSeccion + "', '" + sFlagModif + "', '" + sCodAuxiliar + "', '" + sOrden + "') "
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function BorrarCtaFlujoNRep(ByVal sIdCtaFlujoNRep) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CtaFlujoNRep Where IdCtaFlujoNRep = " + sIdCtaFlujoNRep
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarCtaDetalleNRep(ByVal sIdCtaDetalleNRep As String, ByVal sCtaDetalleNRep As String, ByVal sSigno As String, ByVal sCodCtaOrigen As String, ByVal sTipoDetalle As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String
        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaDetalleNRep WHERE CtaDetalleNRep = '" + sCtaDetalleNRep + "' AND CodCtaOrigen = '" + sCodCtaOrigen + "' AND IdCtaDetalleNRep <> " + sIdCtaDetalleNRep
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'ACTUALIZANDO EL REGISTRO
            sSql = " Update CtaDetalleNRep Set CtaDetalleNRep = '" + sCtaDetalleNRep + "', "
            sSql += "Signo = '" + sSigno + "', "
            sSql += "CodCtaOrigen = '" + sCodCtaOrigen + "', "
            sSql += "TipoDetalle = '" + sTipoDetalle + "', "
            sSql += "Orden = " + sOrden + " "
            sSql += "Where IdCtaDetalleNRep = " + sIdCtaDetalleNRep
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function InsertarCtaDetalleNRep(ByVal IdCtaFlujoNRep As String, ByVal sCtaDetalleNRep As String, ByVal sSigno As String, ByVal sCodCtaOrigen As String, ByVal sTipoDetalle As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaDetalleNRep WHERE CtaDetalleNRep = '" + sCtaDetalleNRep + "' AND CodCtaOrigen = '" + sCodCtaOrigen + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'INSERTANDO EL REGISTRO
            sSql = " Insert into CtaDetalleNRep (IdCtaFlujoNRep, CtaDetalleNRep, Signo, CodCtaOrigen, TipoDetalle, Orden) "
            sSql += "Values ('" + IdCtaFlujoNRep + "', '" + sCtaDetalleNRep + "', '" + sSigno + "', '" + sCodCtaOrigen + "', '" + sTipoDetalle + "', '" + sOrden + "') "
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function BorrarCtaDetalleNRep(ByVal sIdCtaDetalleNRep) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CtaDetalleNRep Where IdCtaDetalleNRep = " + sIdCtaDetalleNRep
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarCtaFlujoRep(ByVal sIdCtaFlujoRep As String, ByVal sCtaFlujoRep As String, ByVal sSigno As String, ByVal sCodSeccion As String, ByVal sFlagModif As String, ByVal sCodAuxiliar As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaFlujoRep WHERE CtaFlujoRep = '" + sCtaFlujoRep + "' AND IdCtaFlujoRep <> " + sIdCtaFlujoRep
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'ACTUALIZANDO EL REGISTRO
            sSql = " Update CtaFlujoRep Set CtaFlujoRep = '" + sCtaFlujoRep + "', "
            sSql += "Signo = '" + sSigno + "', "
            sSql += "CodSeccion = '" + sCodSeccion + "', "
            sSql += "FlagModif = '" + sFlagModif + "', "
            sSql += "CodAuxiliar = '" + sCodAuxiliar + "', "
            sSql += "Orden = '" + sOrden + "' "
            sSql += "Where IdCtaFlujoRep = " + sIdCtaFlujoRep
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function InsertarCtaFlujoRep(ByVal sCtaFlujoRep As String, ByVal sSigno As String, ByVal sCodSeccion As String, ByVal sFlagModif As String, ByVal sCodAuxiliar As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaFlujoRep WHERE CtaFlujoRep = '" + sCtaFlujoRep + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'INSERTANDO EL REGISTRO
            sSql = " Insert into CtaFlujoRep (CtaFlujoRep, Signo, CodSeccion, FlagModif, CodAuxiliar, Orden) "
            sSql += "Values ('" + sCtaFlujoRep + "', '" + sSigno + "', '" + sCodSeccion + "', '" + sFlagModif + "', '" + sCodAuxiliar + "', '" + sOrden + "') "
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function BorrarCtaFlujoRep(ByVal sIdCtaFlujoRep) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CtaFlujoRep Where IdCtaFlujoRep = " + sIdCtaFlujoRep
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarCtaDetalleRep(ByVal sIdCtaDetalleRep As String, ByVal sCtaDetalleRep As String, ByVal sSigno As String, ByVal sCodCtaOrigen As String, ByVal sTipoDetalle As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String
        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaDetalleRep WHERE CtaDetalleRep = '" + sCtaDetalleRep + "' AND CodCtaOrigen = '" + sCodCtaOrigen + "' AND IdCtaDetalleRep <> " + sIdCtaDetalleRep
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'ACTUALIZANDO EL REGISTRO
            sSql = " Update CtaDetalleRep Set CtaDetalleRep = '" + sCtaDetalleRep + "', "
            sSql += "Signo = '" + sSigno + "', "
            sSql += "CodCtaOrigen = '" + sCodCtaOrigen + "', "
            sSql += "TipoDetalle = '" + sTipoDetalle + "', "
            sSql += "Orden = " + sOrden + " "
            sSql += "Where IdCtaDetalleRep = " + sIdCtaDetalleRep
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function InsertarCtaDetalleRep(ByVal IdCtaFlujoRep As String, ByVal sCtaDetalleRep As String, ByVal sSigno As String, ByVal sCodCtaOrigen As String, ByVal sTipoDetalle As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaDetalleRep WHERE CtaDetalleRep = '" + sCtaDetalleRep + "' AND CodCtaOrigen = '" + sCodCtaOrigen + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'INSERTANDO EL REGISTRO
            sSql = " Insert into CtaDetalleRep (IdCtaFlujoRep, CtaDetalleRep, Signo, CodCtaOrigen, TipoDetalle, Orden) "
            sSql += "Values ('" + IdCtaFlujoRep + "', '" + sCtaDetalleRep + "', '" + sSigno + "', '" + sCodCtaOrigen + "', '" + sTipoDetalle + "', '" + sOrden + "') "
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function BorrarCtaDetalleRep(ByVal sIdCtaDetalleRep) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CtaDetalleRep Where IdCtaDetalleRep = " + sIdCtaDetalleRep
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarCtaFlujoS(ByVal sIdCtaFlujo As String, ByVal sCtaFlujo As String, ByVal sSigno As String, ByVal sCodSeccion As String, ByVal sFlagModif As String, ByVal sCodDetalle As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaFlujoS WHERE CodDetalle IS NOT NULL AND CodDetalle = '" + sCodDetalle + "' AND IdCtaFlujo <> " + sIdCtaFlujo
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'ACTUALIZANDO EL REGISTRO
            sSql = " Update CtaFlujoS Set CtaFlujo = '" + sCtaFlujo + "', "
            If sSigno = 0 Then
                sSql += "Signo = NULL, "
            Else
                sSql += "Signo = '" + sSigno + "', "
            End If
            sSql += "CodSeccion = '" + sCodSeccion + "', "
            If sFlagModif = "N" Then
                sSql += "FlagModif = NULL, "
            Else
                sSql += "FlagModif = '" + sFlagModif + "', "
            End If
            If sCodDetalle = "" Then
                sSql += "CodDetalle = NULL, "
            Else
                sSql += "CodDetalle = '" + sCodDetalle + "', "
            End If
            sSql += "Orden = '" + sOrden + "' "
            sSql += "Where IdCtaFlujo = " + sIdCtaFlujo
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "ERROR: El Codigo del Detalle de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function InsertarCtaFlujoS(ByVal sCtaFlujo As String, ByVal sSigno As String, ByVal sCodSeccion As String, ByVal sFlagModif As String, ByVal sCodDetalle As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaFlujoS WHERE CodDetalle IS NOT NULL AND CodDetalle = '" + sCodDetalle + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'INSERTANDO EL REGISTRO
            sSql = " Insert into CtaFlujoS (CtaFlujo, Signo, CodSeccion, FlagModif, CodDetalle, Orden) "
            sSql += "Values ('" + sCtaFlujo + "', "
            If sSigno = 0 Then
                sSql += "NULL, "
            Else
                sSql += "'" + sSigno + "', "
            End If
            sSql += "'" + sCodSeccion + "', "
            If sFlagModif = "N" Then
                sSql += "NULL, "
            Else
                sSql += "'" + sFlagModif + "', "
            End If
            If sCodDetalle = "" Then
                sSql += "NULL, "
            Else
                sSql += "'" + sCodDetalle + "', "
            End If
            sSql += "'" + sOrden + "') "
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"
        End If
        Return "ERROR: El Codigo del Detalle de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function BorrarCtaFlujoS(ByVal sIdCtaFlujo) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CtaFlujoS Where IdCtaFlujo = " + sIdCtaFlujo
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarCtaDetalleS(ByVal sIdCtaDetalle As String, ByVal sCodDetalle As String, ByVal sSigno As String, ByVal sCodCtaOrigen As String, ByVal sTipoDetalle As String) As String
        Dim sSql, sNumReg As String
        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaDetalleS WHERE CodDetalle = '" + sCodDetalle + "' AND CodCtaOrigen = '" + sCodCtaOrigen + "' AND IdCtaDetalle <> " + sIdCtaDetalle
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'ACTUALIZANDO EL REGISTRO
            sSql = " Update CtaDetalleS Set Signo = '" + sSigno + "', "
            sSql += "CodCtaOrigen = '" + sCodCtaOrigen + "', "
            If sTipoDetalle = "" Then
                sSql += "TipoDetalle = NULL "
            Else
                sSql += "TipoDetalle = '" + sTipoDetalle + "' "
            End If
            sSql += "Where IdCtaDetalle = " + sIdCtaDetalle
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function InsertarCtaDetalleS(ByVal sCodDetalle As String, ByVal sSigno As String, ByVal sCodCtaOrigen As String, ByVal sTipoDetalle As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaDetalleS WHERE CodDetalle = '" + sCodDetalle + "' AND CodCtaOrigen = '" + sCodCtaOrigen + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'INSERTANDO EL REGISTRO
            sSql = " Insert into CtaDetalleS (CodDetalle, Signo, CodCtaOrigen, TipoDetalle) "
            sSql += "Values ('" + sCodDetalle + "', '" + sSigno + "', '" + sCodCtaOrigen + "', "
            If sTipoDetalle = "" Then
                sSql += "NULL) "
            Else
                sSql += "'" + sTipoDetalle + "')"
            End If
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function BorrarCtaDetalleS(ByVal sIdCtaDetalle) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CtaDetalleS Where IdCtaDetalle = " + sIdCtaDetalle
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarCtaEGP(ByVal sIdCtaEGP As String, ByVal sCtaEGP As String, ByVal sSigno As String, ByVal sCodSeccion As String, ByVal sFlagModif As String, ByVal sOrden As String) As String
        Dim sSql As String

        sOrden = ValidarImporte(sOrden)
            'ACTUALIZANDO EL REGISTRO
            sSql = " Update CtaEGP Set CtaEGP = '" + sCtaEGP + "', "
            If sSigno = 0 Then
                sSql += "Signo = NULL, "
            Else
                sSql += "Signo = '" + sSigno + "', "
            End If
            sSql += "CodSeccion = '" + sCodSeccion + "', "
            If sFlagModif = "N" Then
                sSql += "FlagModif = NULL, "
            Else
                sSql += "FlagModif = '" + sFlagModif + "', "
            End If
            sSql += "Orden = '" + sOrden + "' "
            sSql += "Where IdCtaEGP = " + sIdCtaEGP
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        
    End Function

    Public Shared Function InsertarCtaEGP(ByVal sCtaEGP As String, ByVal sSigno As String, ByVal sCodSeccion As String, ByVal sFlagModif As String, ByVal sOrden As String) As String
        Dim sSql As String


        sOrden = ValidarImporte(sOrden)
            'INSERTANDO EL REGISTRO
        sSql = " Insert into CtaEGP (CtaEGP, Signo, CodSeccion, FlagModif, Orden) "
        sSql += "Values ('" + sCtaEGP + "', "
            If sSigno = 0 Then
                sSql += "NULL, "
            Else
                sSql += "'" + sSigno + "', "
            End If
            sSql += "'" + sCodSeccion + "', "
            If sFlagModif = "N" Then
                sSql += "NULL, "
            Else
                sSql += "'" + sFlagModif + "', "
            End If
            sSql += "'" + sOrden + "') "
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"

    End Function

    Public Shared Function BorrarCtaEGP(ByVal sIdCtaEGP) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CtaEGP Where IdCtaEGP = " + sIdCtaEGP
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarCtaEGPDet(ByVal sIdCtaEGPDet As String, ByVal sIdCtaEGP As String, ByVal sSigno As String, ByVal sCodCtaOrigen As String, ByVal sTipoDetalle As String) As String
        Dim sSql, sNumReg As String
        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaEGPDet WHERE IdCtaEGP = '" + sIdCtaEGP + "' AND CodCtaOrigen = '" + sCodCtaOrigen + "' AND IdCtaEGPDet <> '" + sIdCtaEGPDet + "'"
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'ACTUALIZANDO EL REGISTRO
            sSql = " Update CtaEGPDet Set Signo = '" + sSigno + "', "
            sSql += "CodCtaOrigen = '" + sCodCtaOrigen + "', "
            If sTipoDetalle = "" Then
                sSql += "TipoDetalle = NULL "
            Else
                sSql += "TipoDetalle = '" + sTipoDetalle + "' "
            End If
            sSql += "Where IdCtaEGPDet = " + sIdCtaEGPDet
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function InsertarCtaEGPDet(ByVal sIdCtaEGP As String, ByVal sSigno As String, ByVal sCodCtaOrigen As String, ByVal sTipoDetalle As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM CtaEGPDet WHERE IdCtaEGP = '" + sIdCtaEGP + "' AND CodCtaOrigen = '" + sCodCtaOrigen + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'INSERTANDO EL REGISTRO
            sSql = " Insert into CtaEGPDet (IdCtaEGP, Signo, CodCtaOrigen, TipoDetalle) "
            sSql += "Values ('" + sIdCtaEGP + "', '" + sSigno + "', '" + sCodCtaOrigen + "', "
            If sTipoDetalle = "" Then
                sSql += "NULL) "
            Else
                sSql += "'" + sTipoDetalle + "')"
            End If
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"
        End If
        Return "ERROR: El nombre de la Cuenta ya se encuentra registrada"
    End Function

    Public Shared Function BorrarCtaEGPDet(ByVal sIdCtaEGPDet) As String
        Dim sSql As String
        'Eliminando el registro
        sSql = "Delete From CtaEGPDet Where IdCtaEGPDet = " + sIdCtaEGPDet
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarAjuste(ByVal sIdAjuste As String, ByVal sIdPeriodo As String, ByVal sFecha As String, ByVal sCodCuentaFC2 As String, ByVal sMontoUSD As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM Ajuste WHERE IdPeriodo = '" + sIdPeriodo + "' AND CodCuentaFC2 = '" + sCodCuentaFC2 + "' AND IdAjuste <> " + sIdAjuste + " "
        sNumReg = ExtraerValor(sSql, "NumReg")
        'If (sNumReg = "0") Then
        'ACTUALIZANDO EL REGISTRO EN CASO EXISTA
        sSql = " Update Ajuste Set IdPeriodo = '" + sIdPeriodo + "', "
        sSql += "Fecha = '" + sFecha + "', "
        sSql += "CodCuentaFC2 = '" + sCodCuentaFC2 + "', "
        'sSql += "MontoBaseUSD = '" + ValidarImporte(sMontoBaseUSD) + "', "
        'sSql += "MontoIGVUSD = '" + ValidarImporte(sMontoIGVUSD) + "', "
        sSql += "MontoUSD = '" + ValidarImporte(sMontoUSD) + "' "
        sSql += "Where IdAjuste = '" + sIdAjuste + "'"
        EjecutarSentencia(sSql)
        'Return "Registro actualizado correctamente"
        'End If
        Return "ERROR: Para ese Periodo y Código, la información ya se encuentra registrada"
    End Function

    Public Shared Function InsertarAjuste(ByVal sIdPeriodo As String, ByVal sFecha As String, ByVal sCodCuentaFC2 As String, ByVal sMontoUSD As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM Ajuste WHERE IdPeriodo = '" + sIdPeriodo + "' AND CodCuentaFC2 = '" + sCodCuentaFC2 + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        'If (sNumReg = "0") Then
        'INSERTANDO EL REGISTRO
        sSql = " Insert into Ajuste (IdPeriodo, Fecha, CodCuentaFC2, MontoUSD) "
        sSql += "Values ('" + sIdPeriodo + "', '" + sFecha + "', '" + sCodCuentaFC2 + "', '" + ValidarImporte(sMontoUSD) + "') "
        EjecutarSentencia(sSql)
        Return "La información ha sido registrada"
        'End If
        'Return "ERROR: Para ese Periodo y Código, la información ya se encuentra registrada"
    End Function

    Public Shared Function BorrarAjuste(ByVal sIdAjuste As String) As String
        Dim sSql As String
        'ELIMINANDO EL REGISTRO
        sSql = " Delete From Ajuste Where IdAjuste = '" + sIdAjuste + "' "
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarAjusteExt(ByVal sIdAjuste As String, ByVal sIdPeriodo As String, ByVal sFecha As String, ByVal sCodTipoAjuste As String, ByVal sCodCuentaFCOrig As String, ByVal sCodPersonaOrig As String, ByVal sCodCuentaFCDes As String, ByVal sCodPersonaDes As String, ByVal sCodAreaDes As String, ByVal sObservacion As String, ByVal sMontoUSD As String) As String
        Dim sSql, sNumReg As String

        'LAS CUENTAS NO DEBEN SER IGUALES
        'If (sCodCuentaFCOrig <> sCodCuentaFCDes) Then
        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        'sSql = "SELECT count(*) as NumReg FROM AjusteExt WHERE IdPeriodo = '" + sIdPeriodo + "' AND CodCuentaFCOrig = '" + sCodCuentaFCOrig + "' AND  CodCuentaFCDes = '" + sCodCuentaFCDes + "' AND IdAjuste <> " + sIdAjuste + " "
        'sNumReg = ExtraerValor(sSql, "NumReg")
        'If (sNumReg = "0") Then
        'ACTUALIZANDO EL REGISTRO EN CASO EXISTA
        sSql = " Update AjusteExt Set IdPeriodo = '" + sIdPeriodo + "', "
        sSql += "Fecha = '" + sFecha + "', "
        sSql += "CodTipoAjuste = '" + sCodTipoAjuste + "', "
        sSql += "CodCuentaFCOrig = '" + sCodCuentaFCOrig + "', "
        sSql += "CodPersonaOrig = '" + sCodPersonaOrig + "', "
        sSql += "CodCuentaFCDes = '" + sCodCuentaFCDes + "', "
        sSql += "CodPersonaDes = '" + sCodPersonaDes + "', "
        sSql += "CodAreaDes = '" + sCodAreaDes + "', "
        sSql += "Observacion = '" + sObservacion + "', "
        sSql += "MontoUSD = '" + ValidarImporte(sMontoUSD) + "' "
        sSql += "Where IdAjuste = '" + sIdAjuste + "'"
        EjecutarSentencia(sSql)
        Return "Registro actualizado correctamente"
        'End If
        'Return "ERROR: Para ese Periodo y Código, la información ya se encuentra registrada"
        'Else
        'Return "ERROR: Las cuentas de Origen y Destino no deben ser iguales"
        'End If
    End Function

    Public Shared Function InsertarAjusteExt(ByVal sIdPeriodo As String, ByVal sFecha As String, ByVal sCodTipoAjuste As String, ByVal sCodCuentaFCOrig As String, ByVal sCodPersonaOrig As String, ByVal sCodCuentaFCDes As String, ByVal sCodPersonaDes As String, ByVal sCodAreaDes As String, ByVal sObservacion As String, ByVal sMontoUSD As String) As String
        Dim sSql, sNumReg As String
        'If (sCodCuentaFCOrig <> sCodCuentaFCDes) Then
        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        'sSql = "SELECT count(*) as NumReg FROM AjusteExt WHERE IdPeriodo = '" + sIdPeriodo + "' AND CodCuentaFCOrig = '" + sCodCuentaFCOrig + "'  AND CodCuentaFCDes = '" + sCodCuentaFCDes + "' "
        'sNumReg = ExtraerValor(sSql, "NumReg")
        'If (sNumReg = "0") Then
        'INSERTANDO EL REGISTRO
        sSql = " Insert into AjusteExt (IdPeriodo, Fecha, CodTipoAjuste, CodCuentaFCOrig, CodPersonaOrig, CodCuentaFCDes, CodPersonaDes, CodAreaDes, Observacion, MontoUSD) "
        sSql += "Values ('" + sIdPeriodo + "', '" + sFecha + "', '" + sCodTipoAjuste + "', '" + sCodCuentaFCOrig + "', '" + sCodPersonaOrig + "', " & _
                    "'" + sCodCuentaFCDes + "', '" + sCodPersonaDes + "', '" + sCodAreaDes + "', '" + sObservacion + "', '" + ValidarImporte(sMontoUSD) + "') "
        EjecutarSentencia(sSql)
        Return "La información ha sido registrada"
        'End If
        'Return "ERROR: Para ese Periodo y Código, la información ya se encuentra registrada"
        'Else
        'Return "ERROR: Las cuentas de Origen y Destino no deben ser iguales"
        'End If
    End Function

    Public Shared Function BorrarAjusteExt(ByVal sIdAjuste As String) As String
        Dim sSql As String
        'ELIMINANDO EL REGISTRO
        sSql = " Delete From AjusteExt Where IdAjuste = '" + sIdAjuste + "' "
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarDetalle(ByVal sCodVersion As String, ByVal sIdPeriodo As String, ByVal sCodDetalle As String, ByVal sRubro As String, ByVal sMontoNeto As String) As Boolean
        Dim sSql As String
        Dim bResultado As Boolean = True

        'ELIMINANDO EL REGISTRO
        sSql = " Delete From DetalleS Where CodVersion = '" + sCodVersion + "' "
        sSql += "AND IdPeriodo = '" + sIdPeriodo + "' "
        sSql += "AND CodDetalle = '" + sCodDetalle + "' "
        sSql += "AND Rubro = '" + sRubro + "' "
        EjecutarSentencia(sSql)

        'INSERTANDO EL REGISTRO
        sSql = " Insert into DetalleS (CodVersion, IdPeriodo, CodDetalle, Rubro, MontoNeto, IGV, MontoBruto) "
        sSql += "Values ('" + sCodVersion + "', '" + sIdPeriodo + "', '" + sCodDetalle + "', '" + sRubro + "', '" + ValidarImporte(sMontoNeto) + "', '0', '" + ValidarImporte(sMontoNeto) + "') "
        EjecutarSentencia(sSql)
        
        Return bResultado
    End Function

    Public Shared Function ActualizarGrupoATV(ByVal sIdEmpresa As String, ByVal sEmpresa As String, ByVal sCodPersona As String, ByVal sNivel As String, ByVal sNivelN As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM GrupoATV WHERE Empresa = '" + sEmpresa + "' AND IdEmpresa <> " + sIdEmpresa + " "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            'ACTUALIZANDO EL REGISTRO EN CASO EXISTA
            sSql = " Update GrupoATV Set Empresa = '" + sEmpresa + "', "
            sSql += "CodPersona = '" + sCodPersona + "', "
            If sNivel = "" Then
                sSql += "Nivel = NULL, "
            Else
                sSql += "Nivel = '" + sNivel + "', "
            End If
            If sNivelN = "" Then
                sSql += "NivelN = NULL, "
            Else
                sSql += "NivelN = '" + sNivelN + "', "
            End If
            sSql += "Orden = '" + sOrden + "' "
            sSql += "Where IdEmpresa = '" + sIdEmpresa + "'"
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "ERROR: El nombre de la empresa ya se encuentra registrada"
    End Function

    Public Shared Function InsertarGrupoATV(ByVal sEmpresa As String, ByVal sCodPersona As String, ByVal sNivel As String, ByVal sNivelN As String, ByVal sOrden As String) As String
        Dim sSql, sNumReg, sAux1, sAux2 As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM GrupoATV WHERE Empresa = '" + sEmpresa + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            sOrden = ValidarImporte(sOrden)
            If sNivel = "" Then
                sAux1 = "NULL"
            Else
                sAux1 = "'" + sNivel + "'"
            End If
            If sNivelN = "" Then
                sAux2 = "NULL "
            Else
                sAux2 = "'" + sNivelN + "'"
            End If
            'INSERTANDO EL REGISTRO
            sSql = " Insert into GrupoATV (Empresa, CodPersona, Nivel, NivelN, Orden) "
            sSql += "Values ('" + sEmpresa + "', '" + sCodPersona + "', " + sAux1 + ", " + sAux2 + ", '" + sOrden + "') "
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"
        End If
        Return "ERROR: El nombre de la empresa ya se encuentra registrada"
    End Function

    Public Shared Function BorrarGrupoATV(ByVal sIdEmpresa As String) As String
        Dim sSql As String
        'ELIMINANDO EL REGISTRO
        sSql = " Delete From GrupoATV Where IdEmpresa = '" + sIdEmpresa + "'"
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function

    Public Shared Function ActualizarPersonaS(ByVal sIdPersonaS As String, ByVal sCodGrupo As String, ByVal sPersonaS As String, ByVal sCodPersona As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM PersonaS WHERE CodPersona = '" + sCodPersona + "' AND IdPersonaS <> '" + sIdPersonaS + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'ACTUALIZANDO EL REGISTRO EN CASO EXISTA
            sSql = " Update PersonaS Set CodGrupo = '" + sCodGrupo + "', "
            sSql += "PersonaS = '" + sPersonaS + "', "
            sSql += "CodPersona = '" + sCodPersona + "' "
            sSql += "Where IdPersonaS = '" + sIdPersonaS + "'"
            EjecutarSentencia(sSql)
            Return "Registro actualizado correctamente"
        End If
        Return "ERROR: La persona ya ha sido asignado a otro grupo"
    End Function

    Public Shared Function InsertarPersonaS(ByVal sCodGrupo As String, ByVal sPersonaS As String, ByVal sCodPersona As String) As String
        Dim sSql, sNumReg As String

        'CONSULTANDO SI EXISTE UN REGISTRO SIMILAR
        sSql = "SELECT count(*) as NumReg FROM PersonaS WHERE CodPersona = '" + sCodPersona + "' "
        sNumReg = ExtraerValor(sSql, "NumReg")
        If (sNumReg = "0") Then
            'INSERTANDO EL REGISTRO
            sSql = " Insert into PersonaS (CodGrupo, PersonaS, CodPersona) "
            sSql += "Values ('" + sCodGrupo + "', '" + sPersonaS + "', '" + sCodPersona + "') "
            EjecutarSentencia(sSql)
            Return "La información ha sido registrada"
        End If
        Return "ERROR: La persona ya ha sido asignado a otro grupo"
    End Function

    Public Shared Function BorrarPersonaS(ByVal sIdPersonaS As String) As String
        Dim sSql As String
        'ELIMINANDO EL REGISTRO
        sSql = " Delete From PersonaS Where IdPersonaS = '" + sIdPersonaS + "'"
        EjecutarSentencia(sSql)
        Return "El registro ha sido eliminado"
    End Function
End Class
