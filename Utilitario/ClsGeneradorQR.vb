Imports ZXing

Public Class ClsGeneradorQR
    Public Const FormatoFecha As String = "yyyy-MM-dd"
    Public Const FormatoHora As String = "HH:mm:ss-05:00"

    Public Class ClsDatosQR

        Public Property NumFac As String
        Public Property FecFac As String
        Public Property HorFac As String
        Public Property NitFac As String
        Public Property DocAdq As String
        Public Property ValFac As String
        Public Property ValIva As String
        Public Property ValOtroIm As String
        Public Property ValTolFac As String
        Public Property CUFE As String


    End Class

    Private Function SaveQR(ByVal Datos As ClsDatosQR, ByVal rutaimgQR As String, Optional ByRef msgg As String = "") As Boolean

        Try
            Dim cadenaQR As String = ""

            cadenaQR = "NumFac:" + Datos.NumFac + vbCrLf _
                + "FecFac:" + Datos.FecFac + vbCrLf _
                + "HorFac:" + Datos.HorFac + vbCrLf _
                + "NitFac:" + Datos.NitFac + vbCrLf _
                + "DocAdq:" + Datos.DocAdq + vbCrLf _
                + "ValFac:" + Datos.ValFac + vbCrLf _
                + "ValIva:" + Datos.ValIva + vbCrLf _
                + "ValOtroIm:" + Datos.ValOtroIm + vbCrLf _
                + "ValTolFac:" + Datos.ValTolFac _
                + "CUFE:" + Datos.CUFE

            Dim escritor As New BarcodeWriter

            escritor.Format = BarcodeFormat.QR_CODE
            '236 px = 2cm tamano minimos establecido por la Dian
            escritor.Options.Height = 236
            escritor.Options.Width = 236

            Dim lienso As System.Drawing.Bitmap

            lienso = escritor.Write(cadenaQR)

            lienso.Save(rutaimgQR, System.Drawing.Imaging.ImageFormat.Jpeg)

            msgg = "ok"
            Return True
        Catch ex As Exception

            msgg = "Ocurrio un error al generar el QR : " & ex.Message

        End Try

        Return False

    End Function

    Private Shared Function ConvertirDoubleToDecimalForXML(Valor As Decimal, Optional ByVal NumDecimales As Integer = 2) As String
        Dim valorSTR As String = String.Empty
        Dim Formato As String = "##0.00"
        Select Case NumDecimales
            Case 1 : Formato = "##0.0"
            Case 2 : Formato = "##0.00"
            Case 3 : Formato = "##0.000"
            Case 4 : Formato = "##0.0000"
            Case 5 : Formato = "##0.00000"
            Case 6 : Formato = "##0.000000"
        End Select
        Valor = Math.Round(Valor, NumDecimales, MidpointRounding.AwayFromZero)
        valorSTR = (Valor.ToString(Formato).Replace(",", "."))
        Return valorSTR
    End Function

End Class
