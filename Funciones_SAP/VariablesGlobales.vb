Public Class VariablesGlobales
    Public Shared _gNombreAddOn As String
    Public Shared _gVersionAddOn As String

    Public Shared _CabeceraQRY As String

    Public Shared _EmiReceptorQRY As String

    Public Shared _ImpuestosQRY As String

    Public Shared _DesCargosQRY As String

    Public Shared _InformacionFiscalQRY As String

    Public Shared _ReteQRY As String

    Public Shared _AnticipoQRY As String

    Public Shared _ItotalesQRY As String

    Public Shared _FpagoQRY As String

    Public Shared _DetalleQry As String

    Public Shared _AdicionalFCQry As String

    Public Shared _AdicionalNCQry As String

    Public Shared _AdicionalNDQry As String

    Public Shared _DocsEnviadosQRY As String

    Public Shared _DocsIntegradosQRY As String

    Public Shared _EntregaQRY As String
    'VARIABLES GOBALES PARA LAS URLS DE LOS WS

    Public Shared _gEmisionTipo As String
    Public Shared _gEmisionClave As String
    Public Shared _gWS_RecepcionClave As String

    Public Shared _gGuardaLogEmision As String

    Public Shared _gWS_EmisionFC As String = ""
    Public Shared _gWS_EmisionND As String = ""
    Public Shared _gWS_EmisionNC As String = ""
    Public Shared _gWS_EmisionConsulta As String = ""
    Public Shared _gWS_EmisionConsultaFiles As String = ""
    Public Shared _gWS_ReenvioMail As String = ""
    Public Shared _gWS_Utilidades As String = ""

    'RECEPCION
    Public Shared _gWS_RecepcionConsulta As String = ""
    Public Shared _gWS_RecepcionEstado As String = ""
    Public Shared _gWS_RecepcionMR As String = ""

    'OTRAS AJUSTES GLOBALES
    Public Shared _gTimeOut_Emision As String = ""
    'Public Shared _gIsInHouse As String = ""
    Public Shared _gNO_ConsumirMetodoHTTPS As String = ""
    Public Shared _gRUCEmisor As String = ""
    Public Shared _gTipoSocioNegocio As String = ""
    Public Shared _gGeneracion_Cufe_QR As String = ""
    Public Shared _gRutaRPT As String = ""
    Public Shared _gUsuarioDB As String = ""
    Public Shared _gPasswordDB As String = ""
    Public Shared _gAdjuntos As String = ""
    '- Necesario para Generacion de Cufe
    Public Shared _gAmbiente As Integer = 0
    Public Shared _gPinSoftware As String = ""

    'LICENCIA
    Public Shared _gTipoLicenciaAddOn As String = ""
    Public Shared _gTieneLicenciaActivaAddOn As Boolean = False
    Public Shared _CorreoResponsable As String = ""
    Public Shared _gVersiondelMinisterioDeHacienda As String = ""

    Public Shared _gVersionDisponibleAddOn As String = ""
    Public Shared _gReleaseNoteAddOn As String = ""

    'VALIDACIONES
    Public Shared _gValidacionNit As String = ""
    Public Shared _gValidacionObligacionFiscales As String = ""
    Public Shared _gVTipoFactura As String = ""
    Public Shared _gVTipoOperacionDoc As String = ""
    Public Shared _gVTipoNotaCredito As String = ""
    Public Shared _VTipoNotaDebito As String = ""
    Public Shared _gVtipoDescuento As String = ""
    Public Shared _gVMediodePago As String = ""
    Public Shared _gVInfoReferencia As String = ""
    Public Shared _gVImpuestosMapeados As String = ""
    Public Shared _gVCamposExportacion As String = ""
    Public Shared _gVmailReceptor As String = ""
    Public Shared _gNombreCampoTipoIdentificacion As String = ""
    Public Shared _gNombreCampoDireccion As String = ""
    Public Shared _gNombreCampoMunicipio As String = ""

End Class
