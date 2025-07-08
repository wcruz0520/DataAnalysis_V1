Option Strict Off
Option Explicit On

Public Class Estructura

    'Public WithEvents rSboApp As SAPbouiCOM.Application
    Public LicenciaAddon As String = "" 'EMISION, RECEPCION, FULL

    Public VersionAddon As String = "1.0.0"

    ''' <summary>
    '''  Inicializaciòn de la Clase
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        Try
            rSboApp = rSboGui.GetApplication
            oFuncionesB1 = New Funciones_SAP.FuncionesB1(rCompany, rSboApp)
            oFuncionesB1.mostrarMensajesError = False
            oFuncionesB1.mostrarMensajesExito = True
            oFuncionesB1.mantenerLogErrores = True

            oFuncionesB1.validarVersion_SoloCrearTabla()

        Catch ex As Exception
        End Try
    End Sub

    Public Sub CreacionDeEstructura()
        If Not oFuncionesB1.validarVersion(NombreAddon, VersionAddon) Then
            rSboApp.StatusBar.SetText(NombreAddon + " - Validando la estructura necesaria para el correcto funcionamiento del Addon.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            FUN_CreaTablas()
            FUN_CreaCampos()

            'FUN_CreaUDO_LOG()  'LOG DEL PROCESO
            FUN_CreaUDO_CONF() 'CATALOGO
            FUN_CreaUDO()

            'CREACION UDO PLANT RPT
            FUN_CreaUDO_PlantRPT()

            FUN_CreaUDO_Cartera()

            'oFuncionesB1.confirmarVersion(NombreAddon, VersionAddon)
        Else
            rSboApp.SetStatusBarMessage(NombreAddon + " - Su estructura de la base de datos esta actualizada para la version del AddOn: " + VersionAddon, SAPbouiCOM.BoMessageTime.bmt_Short, False)
        End If
    End Sub

    ''' <summary>
    ''' Crea las tablas de usuario 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FUN_CreaTablas()
        Try

            '' CONFIGURACIÓN - CATALOGO
            'oFuncionesB1.creaTablaMD("SSPEINFODET", "(SS) Info Pago Ef. Enviada Cab", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("SSUDOCONF", "(GS) CONFIGURACION", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("SSUDOCONFD", "(GS) CONFIGURACION DETALLE", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            oFuncionesB1.creaTablaMD("SSVALPLANTCAB", "(SS) Val Plant Cab", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("SSVALPLANTDET", "(SS) Val Plant Det", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            oFuncionesB1.creaTablaMD("SSTERCEROS", "(SS) Tabla de Terceros", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            'oFuncionesB1.creaTablaMD("SSTERCEROS", "(SS) Tabla de Terceros", SAPbobsCOM.BoUTBTableType.bott_Document)
            'oFuncionesB1.creaTablaMD("SSTERCEROSDET", "(SS) Tabla detalle Terceros", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            'TABLAS PLANTILLAS RPT
            oFuncionesB1.creaTablaMD("SSPLANT_RPT", "CAB Plant RPT", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("SSPLANT_RPT_DET1", "DET1 Plant RPT", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            oFuncionesB1.creaTablaMD("SS_CARTERA_CAB", "(SS) Antiguedad Cartera", SAPbobsCOM.BoUTBTableType.bott_Document)
            oFuncionesB1.creaTablaMD("SS_CARTERA_DET1", "(SS) Detalle 1 Cartera", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaTablas_Catch , Error: " & ex.Message.ToString(), "Estructura")
        Finally
            GC.Collect()
        End Try
    End Sub

    ''' <summary>
    ''' Crea los campos de Usuario (UDF)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FUN_CreaCampos()
        Try
            'Dim StrcbdVal() As String = {"SI", "NO"}
            'Dim StrcbdDes() As String = {"SI", "NO"}

            '' CONFIGURACION ADDON
            oFuncionesB1.creaCampoMD("SSUDOCONF", "Modulo", "(SS) Modulo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSUDOCONF", "Tipo", "(SS) Tipo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSUDOCONF", "Subtipo", "(SS) Subtipo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSUDOCONFD", "Nombre", "(SS) Nombre", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSUDOCONFD", "Valor", "(SS) Valor", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SSVALPLANTCAB", "Anio", "(SS) Año", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTCAB", "CodSuc", "(SS) Cod Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTCAB", "NomSuc", "(SS) Nom Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTCAB", "Adicional", "(SS) Adicional", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTCAB", "IdTercero", "(SS) Id Tercero", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTCAB", "NomTercero", "(SS) Nombre Tercero", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'U_NomTercero

            'oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Nivel1", "(SS) Nivel1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Nivel2", "(SS) Nivel2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Nivel3", "(SS) Nivel3", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "NomCuenta", "(SS) Nombre Cuenta", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "CtaAsoc", "(SS) Cta. Asociada", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "NomCtaAsc", "(SS) Nombre Cuenta Asociada", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "CodAgrup", "(SS) Cod. Agrupador", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            Dim StrcbdVal() As String = {"P", "A", "R"}
            Dim StrcbdDes() As String = {"Principal", "Adicional", "Real"}
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Tipo", "(SS) Tipo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 7, SAPbobsCOM.BoYesNoEnum.tNO, StrcbdVal, StrcbdDes, "P")
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "LineaNeg", "(SS) Linea de Negocio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "SubLinea", "(SS) SubLinea", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Sucursal", "(SS) Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Espacio", "(SS) Espacio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Espacio2", "(SS) Espacio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Proyecto", "(SS) Proyecto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "NomProy", "(SS) Nomb Proyecto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Enero", "(SS) Enero", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Febrero", "(SS) Febrero", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Marzo", "(SS) Marzo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Abril", "(SS) Abril", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Mayo", "(SS) Mayo", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Junio", "(SS) Junio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Julio", "(SS) Julio", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Agosto", "(SS) Agosto", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Septiembre", "(SS) Septiembre", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Octubre", "(SS) Octubre", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Noviembre", "(SS) Noviembre", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSVALPLANTDET", "Diciembre", "(SS) Diciembre", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)

            'CREACION CAMPOS PLANTILLA RPT
            'CABECERA
            oFuncionesB1.creaCampoMD("SSPLANT_RPT", "NAME_RPT", "(SS) Nombre Informe", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT", "CODE_RPT", "(SS) Código Informe", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT", "ANEXO", "(SS) Anexo/#Hoja", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            'DETALLE
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "TITULO1", "(SS) Titulo 1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "NIVEL1", "(SS) Cuenta Nivel 1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "TITULO2", "(SS) Titulo 2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "NIVEL2", "(SS) Cuenta Nivel 2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "TITULO3", "(SS) Titulo 3", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "NIVEL3", "(SS) Cuenta Nivel 3", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "TITULO4", "(SS) Titulo 4", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "NIVEL4", "(SS) Cuenta Nivel 4", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "TITULO5", "(SS) Titulo 5", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "NIVEL5", "(SS) Cuenta Nivel 5", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "TITULO6", "(SS) Titulo 6", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "NIVEL6", "(SS) Cuenta Nivel 6", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "PROYECTO", "(SS) Proyecto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "LINEANEGOCIO", "(SS) Linea Negocio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "SUBLINEA", "(SS) SubLinea Negocio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "SUCURSAL", "(SS) Sucursal", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSPLANT_RPT_DET1", "AREA_ESPACIO", "(SS) Area/Espacio", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            'TABLA TERCEROS
            'oFuncionesB1.creaCampoMD("SSTERCEROSDET", "IdTercero", "(SS) Identificacion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            'oFuncionesB1.creaCampoMD("SSTERCEROSDET", "NomTercero", "(SS) Nombre Tercero", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSTERCEROS", "IdTercero", "(SS) Identificacion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SSTERCEROS", "NomTercero", "(SS) Nombre Tercero", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            'TABLAS CARTERA CAB
            oFuncionesB1.creaCampoMD("SS_CARTERA_CAB", "FECHA", "(SS) Fecha Generación", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)

            'TABLA CARTERA DET
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "NumInter", "(SS) N° Interno", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "DocNum", "(SS) N° Doc. Consulta", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "CardCode", "(SS) Código Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "CardName", "(SS) Nombre Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Tipo_Doc", "(SS) Tipo Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "FolioNum", "(SS) FolioNum", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "DocDate", "(SS) Fecha Documento", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "DocDueDate", "(SS) Fecha Vencimiento", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Total_Cart", "(SS) Total Cartera", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Dias_Venc", "(SS) Días vencidos", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "CartxVenc", "(SS) Cartera x Vencer", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Cart_Venc", "(SS) Cartera Vencida", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "TotalAbon", "(SS) Total Abono", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Venc_30", "(SS) Vencida 30", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Venc_60", "(SS) Vencida 60", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Venc_90", "(SS) Vencida 90", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Venc_120", "(SS) Vencida 120", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Venc_m120", "(SS) Vencida > 120", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 100, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "CodVend", "(SS) Código Vendedor", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "NomVend", "(SS) Nombre Vendedor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "NumAut", "(SS) Número Autorización", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Denomi", "(SS) Denominacion", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "Estab", "(SS) Establecimiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "PtoEm", "(SS) Punto Emisión", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
            oFuncionesB1.creaCampoMD("SS_CARTERA_DET1", "ObjType", "(SS) Tipo Objeto Doc", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, SAPbobsCOM.BoYesNoEnum.tNO)
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaCampos_Catch , Error: " & ex.Message.ToString(), "Estructura")
        Finally
            GC.Collect()
        End Try

    End Sub

    Private Sub FUN_CreaUDO_CONF()
        Dim oUdo As SAPbobsCOM.UserObjectsMD
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SSUDOCONF") Then
                oUdo.Code = "SSUDOCONF"
                oUdo.Name = "(SS) Configuracion UDO"
                oUdo.TableName = "SSUDOCONF"
                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                'oUdo.LogTableName = "A_GS_CONF"

                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document

                oUdo.ChildTables.TableName = "SSUDOCONFD"

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_CONF , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear UDO CONF" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SS_CONF, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_CONF_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub FUN_CreaUDO()
        Dim oUdo As SAPbobsCOM.UserObjectsMD
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SSVALPLANTCAB") Then
                oUdo.Code = "SSVALPLANTCAB"
                oUdo.Name = "SSVALPLANTCAB"
                oUdo.TableName = "SSVALPLANTCAB"
                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_SSVALPLANTCAB"

                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document

                oUdo.ChildTables.TableName = "SSVALPLANTDET"

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_HORASEQUIPO , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear SSVALPLANTCAB" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SSVALPLANTCAB, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_HORASEQUIPO_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub FUN_CreaUDO_PlantRPT()
        Dim oUdo As SAPbobsCOM.UserObjectsMD
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SSPLANT_RPT") Then
                oUdo.Code = "SSPLANT_RPT"
                oUdo.Name = "SSPLANT_RPT"
                oUdo.TableName = "SSPLANT_RPT"
                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_SSPLANT_RPT"

                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document

                oUdo.ChildTables.TableName = "SSPLANT_RPT_DET1"

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_PlantRPT , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear SSPLANT_RPT" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SSPLANT_RPT, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_PlantRPT_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub FUN_CreaUDO_Cartera()
        Dim oUdo As SAPbobsCOM.UserObjectsMD
        Dim lRetCode As Integer = 0
        Dim sErrMsg As String = ""
        Dim lErrCode As Integer = 0

        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD)
            GC.Collect() 'Release the handle to the table

            oUdo = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If Not oUdo.GetByKey("SS_CARTERA_CAB") Then
                oUdo.Code = "SS_CARTERA_CAB"
                oUdo.Name = "SS_CARTERA_CAB"
                oUdo.TableName = "SS_CARTERA_CAB"
                oUdo.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUdo.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUdo.LogTableName = "A_SS_CARTERA_CAB"

                oUdo.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document

                oUdo.ChildTables.TableName = "SS_CARTERA_DET1"

                lRetCode = oUdo.Add
                '// check for errors in the process
                If lRetCode <> 0 Then
                    rCompany.GetLastError(lErrCode, sErrMsg)
                    Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_HORASEQUIPO , Error: " + sErrMsg.ToString(), "Estructura")
                    rSboApp.StatusBar.SetText(NombreAddon + " - Error al Crear SS_CARTERA_CAB" + sErrMsg.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    rSboApp.StatusBar.SetText(NombreAddon + " - UDO : SS_CARTERA_CAB, Creado Exitosamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("FUN_CreaUDO_HORASEQUIPO_Catch , Error: " + ex.Message.ToString(), "Estructura")
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUdo)
            oUdo = Nothing
            GC.Collect()
        End Try
    End Sub

End Class


