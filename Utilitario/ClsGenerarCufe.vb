
Public Class ClsGenerarCufe
    Public Const FormatoFecha As String = "yyyy-MM-dd"
    Public Const FormatoHora As String = "HH:mm:ss-05:00"

    Public Class ClsDatosCUFE
        Public Property NumFac As String
        Public Property FecFac As String
        Public Property ValFac As String
        Public Property ValImp1 As String
        Public Property ValImp2 As String
        Public Property ValImp3 As String
        Public Property ValImp As String
        Public Property NitOFE As String
        Public Property TipAdq As String
        Public Property NumAdq As String
        Public Property ClTec As String
        Public Property TipoAmbiente As String
    End Class

    Public Enum Ambiente
        PRUEBA
        PRODUCCION
    End Enum

    Private Function GetCUFE(ByVal data As ClsDatosCUFE)
        Dim cufe As String = String.Format("{0}{1}{2}01{3}04{4}03{5}{6}{7}{8}{9}{10}{11}", data.NumFac, data.FecFac, data.ValFac, data.ValImp1, data.ValImp2, data.ValImp3, data.ValImp, data.NitOFE, data.TipAdq, data.NumAdq, data.ClTec, data.TipoAmbiente)
        Return New ClsGenerarCufe().getSHA384(cufe)
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

    Private Function getSHA384(ByVal strToHash As String) As String
        Dim cryp As New System.Security.Cryptography.SHA384CryptoServiceProvider
        Dim ByteString() As Byte = System.Text.Encoding.ASCII.GetBytes(strToHash)
        ByteString = cryp.ComputeHash(ByteString)
        Dim strResult As String = ""
        For Each b As Byte In ByteString
            strResult &= b.ToString("x2")
        Next
        Return strResult
    End Function


End Class
