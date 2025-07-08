Option Strict Off
Option Explicit On

Imports System.Xml
Imports System.Windows.Forms
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security

Module SubMain

    Public Const NombreAddon As String = "DATA_ANALYSIS"
    Public Const VersionAddon As String = "1.0.0"
    Public Const CodigoPais As String = "EC"

    Public Const sKey As String = "S01s7p1" ' CLAVE DE ENCRIPTACION LICENCIA


    Public rSboApp As SAPbouiCOM.Application
    Public rSboGui As SAPbouiCOM.SboGuiApi
    Public rCompany As SAPbobsCOM.Company
    Public rEstructura As Estructura
    Public rEvento As Eventos
    Private Const conectarse As Boolean = True

    Dim strSQL As String
    Dim ret As Integer

    Public oFuncionesB1 As Funciones_SAP.FuncionesB1

    Public ofrmAcercaDe As frmAcercaDe
    Public ofrmConfClave As frmConfClave
    Public ofrmConfMenu As frmConfMenu
    Public ofrmParametrosAddon As frmParametrosAddon
    Public ofrmPlantilla As frmPlantilla
    Public ofrmEstadistico As frmEstadistica
    Public ofrmCartera As frmCartera

    Public ofrmPlantillaRPT As frmPlantRPT

    '--------------FILTROS---------------------------
    Public oFiltros As SAPbouiCOM.EventFilters
    Public oFiltro As SAPbouiCOM.EventFilter

    Public Sub main()
        Try

            Dim strTest(4) As String, sCookie As String
            Dim strConnString As String
            Dim textoMensajeSinLicencia As String = ""

            strConnString = vbNullString
            strTest = System.Environment.GetCommandLineArgs()

            ' Validaciones de seguridad del AddOn
            If strTest.Length > 0 Then
                If strTest.Length > 1 Then
                    If strTest(0).LastIndexOf("\") > 0 Then
                        strConnString = strTest(1)
                    Else
                        strConnString = strTest(0)
                    End If
                Else
                    If strTest(0).LastIndexOf("\") = -1 Then
                        strConnString = strTest(0)
                    Else
                        System.Windows.Forms.MessageBox.Show("El Add-on se debe ejecutar desde SAP Business One. (" & NombreAddon & "-Err1)")
                        End
                    End If
                End If
            Else
                System.Windows.Forms.MessageBox.Show("El Add-on se debe ejecutar desde SAP Business One. (" & NombreAddon & "-Err2)")
                End
            End If

            ' Conexión
            If strConnString.Length > 0 Then
                Try

                    rSboGui = New SAPbouiCOM.SboGuiApi
                    rCompany = New SAPbobsCOM.Company()
                    ' Conexión con el UI
                    rSboGui.Connect(strConnString)
                    rSboApp = rSboGui.GetApplication()

                    ' Conexión con el DI
                    If conectarse Then
                        rCompany = rSboApp.Company.GetDICompany()
                    Else
                        sCookie = rCompany.GetContextCookie
                        ret = rCompany.SetSboLoginContext(rSboApp.Company.GetConnectionContext(sCookie))
                        If ret = 0 Then
                            ret = rCompany.Connect()
                            If ret <> 0 Then
                                rCompany.GetLastError(ret, strSQL)
                                rSboApp.StatusBar.SetText("Error al Conectar el Add-On " & NombreAddon & ": " & strSQL, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Utilitario.Util_Log.Escribir_Log("Error al Conectar el Add-On " & NombreAddon & ": " & strSQL, "SubMain")
                                End
                            End If
                        Else
                            rSboApp.StatusBar.SetText("No se ha Conectado con AddOn " & NombreAddon & ": Error " & ret, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Utilitario.Util_Log.Escribir_Log("No se ha Conectado con AddOn " & NombreAddon & ": Error " & ret, "SubMain")
                            End
                        End If

                    End If

                    ' **** INICIALIZACIÓN DE LAS CLASES ****

                    'ESTRUCTURA DE DATOS / BUSQ. FORMATEADAS
                    rEstructura = New Estructura
                    Funciones_SAP.VariablesGlobales._gNombreAddOn = NombreAddon
                    Funciones_SAP.VariablesGlobales._gVersionAddOn = VersionAddon

                    'EVENTOS GENERALES - ACERCA DE
                    rEvento = New Eventos


                    ' FORMULARIOS DE CONFIGURACION / PARAMETRIZACION
                    ofrmAcercaDe = New frmAcercaDe(rCompany, rSboApp)
                    ofrmConfClave = New frmConfClave(rCompany, rSboApp)
                    ofrmConfMenu = New frmConfMenu(rCompany, rSboApp)
                    ofrmParametrosAddon = New frmParametrosAddon(rCompany, rSboApp)
                    ofrmPlantilla = New frmPlantilla(rCompany, rSboApp)
                    ofrmPlantillaRPT = New frmPlantRPT(rCompany, rSboApp)

                    ofrmEstadistico = New frmEstadistica(rCompany, rSboApp)
                    ofrmCartera = New frmCartera(rCompany, rSboApp)

                    Funciones_SAP.VariablesGlobales._gTipoLicenciaAddOn = "FULL"

                    Menu_LicenciaFULL()

                    rEstructura.LicenciaAddon = "FULL"
                    rSboApp.StatusBar.SetText("Se Conecto el add-On Data Analysis", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    If Not textoMensajeSinLicencia = "" Then
                        rSboApp.MessageBox(textoMensajeSinLicencia)
                    End If

                Catch ex As Exception
                    System.Windows.Forms.MessageBox.Show("No se ha Conectado con AddOn " & NombreAddon & ": " & ex.Message)
                    End
                End Try

            End If

            System.Windows.Forms.Application.Run()

            Exit Sub

        Catch exMain As Exception
            System.Windows.Forms.MessageBox.Show("Error iniciando el Add-On " & NombreAddon & ": " & exMain.Message)
        End Try
    End Sub

    Private Sub Menu_LicenciaFULL()
        Dim sPath As String
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        oMenus = rSboApp.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        Try
            oMenuItem = rSboApp.Menus.Item("43520") 'Menu principal
            sPath = Application.StartupPath
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "mnPrincipalUdoSeruvi"
            oCreationPackage.String = "Datos para Analíticas"
            oCreationPackage.Enabled = True
            oCreationPackage.Image = sPath & "\" & "logoPPTO3_redim.png" ' "logoV2.jpg"
            oCreationPackage.Position = 15
            oMenus = oMenuItem.SubMenus
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oMenuItem = rSboApp.Menus.Item("mnPrincipalUdoSeruvi")
            oMenus = oMenuItem.SubMenus

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "UDO1"
            oCreationPackage.String = "Presupuesto"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "UDO9"
            oCreationPackage.String = "Estadísticos"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "UDO7"
            oCreationPackage.String = "Cartera"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "UDO3"
            oCreationPackage.String = "Plantilla Reportes"
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            'oCreationPackage.UniqueID = "MNU_DEFINICIONES"
            'oCreationPackage.String = "Definiciones"
            'oCreationPackage.Enabled = True
            'Try
            '    oMenus.AddEx(oCreationPackage)
            'Catch ex As Exception
            'End Try

            '' Submenú: "Terceros" dentro de "Definiciones"
            'oMenuItem = rSboApp.Menus.Item("MNU_DEFINICIONES")
            'oMenus = oMenuItem.SubMenus

            'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            'oCreationPackage.UniqueID = "MNU_TERCEROS"
            'oCreationPackage.String = "Terceros"
            'oCreationPackage.Enabled = True
            'Try
            '    oMenus.AddEx(oCreationPackage)
            'Catch ex As Exception
            'End Try

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING  ' CONFIGURACION
            oCreationPackage.UniqueID = "UDO2"
            'oCreationPackage.Image = sPath & "\" & "conf_redim.png"
            oCreationPackage.String = "Acerca De.."
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try

            'oMenuItem = rSboApp.Menus.Item("UDO1")
            'oMenus = oMenuItem.SubMenus
            'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            'oCreationPackage.UniqueID = "UDO11"
            'oCreationPackage.String = "Pantalla Principal"
            'Try
            '    oMenus.AddEx(oCreationPackage)
            'Catch ex As Exception
            'End Try


        Catch ex As Exception
            rSboApp.SetStatusBarMessage("Error al cargar Menu Licencia FULL" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub
    Private Sub Menu_SinLicencia()
        Dim sPath As String
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        oMenus = rSboApp.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = rSboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        Try
            oMenuItem = rSboApp.Menus.Item("43520") 'Menu principal
            sPath = Application.StartupPath
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "mnPrincipal1"
            oCreationPackage.String = "Integración Shopify"
            oCreationPackage.Enabled = True
            oCreationPackage.Image = sPath & "\" & "logo11.png"
            oCreationPackage.Position = 16
            oMenus = oMenuItem.SubMenus
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try


            oMenuItem = rSboApp.Menus.Item("mnPrincipal1")
            oMenus = oMenuItem.SubMenus

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oCreationPackage.UniqueID = "GS301"
            oCreationPackage.String = "Acerca De.."
            oCreationPackage.Enabled = True
            Try
                oMenus.AddEx(oCreationPackage)
            Catch ex As Exception
            End Try


        Catch ex As Exception
            rSboApp.SetStatusBarMessage("Error al cargar Menu SIN Licencia" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub SetFiltros()

        oFiltros = New SAPbouiCOM.EventFilters()


        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        oFiltro.AddEx("frmPlantilla")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
        oFiltro.AddEx("frmPlantilla")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFiltro.AddEx("frmPlantilla")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFiltro.AddEx("frmPlantilla")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD)
        oFiltro.AddEx("frmPlantilla")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
        oFiltro.AddEx("frmPlantilla")


        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD)
        oFiltro.AddEx("frmEstadistica")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
        oFiltro.AddEx("frmEstadistica")


        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD)
        oFiltro.AddEx("frmCartera")

        oFiltro = oFiltros.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
        oFiltro.AddEx("frmCartera")

        rSboApp.SetFilter(oFiltros)

    End Sub

    Function customCertValidation(ByVal sender As Object, _
                                     ByVal cert As X509Certificate, _
                                     ByVal chain As X509Chain, _
                                     ByVal errors As SslPolicyErrors) As Boolean
        Return True
    End Function

End Module