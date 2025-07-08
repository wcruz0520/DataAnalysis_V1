Imports SAPbobsCOM

Public Class frmParametrosAddon
    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Dim mensaje As String = ""
    Private mors As SAPbobsCOM.Recordset = Nothing
    Dim oUserDataSource As SAPbouiCOM.UserDataSource
    Dim odt As SAPbouiCOM.DataTable
    Dim Alia As String = ""

    Dim cbxCO As SAPbouiCOM.ComboBox
    Dim cbxINH As SAPbouiCOM.ComboBox

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioParametrosADDON()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmParametrosAddon") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmParametrosAddon.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmParametrosAddon").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmParametrosAddon")

            Dim CHK_GLE As SAPbouiCOM.CheckBox ' Activar Que se Guarde Log de Emision en GS_LOG
            CHK_GLE = oForm.Items.Item("CHK_GLE").Specific
            oForm.DataSources.UserDataSources.Add("CHK_GLE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            CHK_GLE.ValOn = "Y"
            CHK_GLE.ValOff = "N"
            CHK_GLE.DataBind.SetBound(True, "", "CHK_GLE")

            'Label que muestra la version tributaria , se lo coloca en Negrita
            Dim LBVH As SAPbouiCOM.StaticText
            LBVH = oForm.Items.Item("LBVH").Specific
            LBVH.Item.TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            LBVH.Item.FontSize = 13
            LBVH.Caption = Funciones_SAP.VariablesGlobales._gVersiondelMinisterioDeHacienda


            CargaDatos()


            oForm.Visible = True
            oForm.Select()

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " - Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub CargaDatos()
        oForm = rsboApp.Forms.Item("frmParametrosAddon")
        oForm.Freeze(True)
        Try
            Dim ACTUALIZA As Integer = 0
            ' DATA TABLE CABECERA
            Try
                oForm.DataSources.DataTables.Add("odt")
            Catch ex As Exception
            End Try
            Dim QueryFC As String = ""
            If rCompany.DbServerType = BoDataServerTypes.dst_HANADB Then
                QueryFC = "SELECT A.""U_Nombre"",A.""U_Valor"" "
                QueryFC += "FROM ""@SSUDOCONFD"" A INNER JOIN "
                QueryFC += """@SSUDOCONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                QueryFC += " WHERE  B.""U_Modulo"" = '" + Funciones_SAP.VariablesGlobales._gNombreAddOn + "' AND B.""U_Tipo"" = 'PARAMETROS' "
                QueryFC += " AND B.""U_Subtipo"" = 'CONFIGURACION'"
            Else
                QueryFC = "SELECT A.U_Nombre,A.U_Valor "
                QueryFC += "FROM ""@SSUDOCONFD"" A WITH(NOLOCK) INNER JOIN "
                QueryFC += """@SSUDOCONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                QueryFC += " WHERE  B.U_Modulo = '" + Funciones_SAP.VariablesGlobales._gNombreAddOn + "' AND  B.U_Tipo = 'PARAMETROS' "
                QueryFC += " AND  B.U_Subtipo = 'CONFIGURACION'"
            End If

            ' CARGANDO CONFIGURACION DE FACTURAS
            oForm.DataSources.DataTables.Item("odt").ExecuteQuery(QueryFC)
            odt = oForm.DataSources.DataTables.Item("odt")
            cbxINH = oForm.Items.Item("cbxINH").Specific

            Dim i As Integer
            For i = 0 To odt.Rows.Count - 1

                If odt.GetValue("U_Nombre", i).ToString().Equals("Param_Ambiente") Then
                    cbxCO = oForm.Items.Item("cbxCO").Specific
                    If Not odt.GetValue("U_Valor", i).ToString() = "" Then
                        cbxCO.Select(odt.GetValue("U_Valor", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Param_Inhouse") Then
                    If Not odt.GetValue("U_Valor", i).ToString() = "" Then
                        cbxINH.Select(odt.GetValue("U_Valor", i).ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ActivarQu eSeGuardeLogEmision") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_GLE")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("WS_Licencia") Then
                    If Not odt.GetValue("U_Valor", i).ToString() = "" Then
                        oForm.Items.Item("ws_LIC").Specific.value = odt.GetValue("U_Valor", i).ToString()
                    Else
                        oForm.Items.Item("ws_LIC").Specific.value = "http://labcr.guru-soft.com/eDocCR/Sitios_Solsap/WS_LICENCIA_SAP/Licencia.svc" 'LICENCIA
                    End If
                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("Param_Ruc") Then
                    oForm.Items.Item("txtNIT").Specific.value = odt.GetValue("U_Valor", i).ToString()

                ElseIf odt.GetValue("U_Nombre", i).ToString().Equals("ActivarQueSeGuardeLogEmision") Then
                    oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_GLE")
                    oUserDataSource.ValueEx = odt.GetValue("U_Valor", i).ToString()


                End If


                ACTUALIZA = 1
            Next

            If ACTUALIZA = 1 Then
                Dim obtnGrabar As SAPbouiCOM.Button
                obtnGrabar = oForm.Items.Item("obtnGrabar").Specific
                obtnGrabar.Caption = "Actualizar"
            End If

        Catch ex As Exception
            rsboApp.MessageBox(ex.Message.ToString())
        Finally
            oForm.Freeze(False)
            mors = Nothing
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormTypeEx = "frmParametrosAddon" Then
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "cbxINH"
                                    oForm = rsboApp.Forms.Item("frmParametrosAddon")

                                Case "cbxCO"
                                    oForm = rsboApp.Forms.Item("frmParametrosAddon")
                                    Dim cbxCO As SAPbouiCOM.ComboBox
                                    cbxCO = oForm.Items.Item("cbxCO").Specific
                                    Dim ws_LIC As SAPbouiCOM.EditText
                                    ws_LIC = oForm.Items.Item("ws_LIC").Specific 'WS_LICENCIA

                                    If cbxCO.Value = "PRUEBAS" Then
                                        ws_LIC.Value = "https://labcr.guru-soft.com/eDocCR/Sitios_Solsap/WS_LICENCIA_SAP/Licencia.svc" 'LICENCIA
                                    ElseIf cbxCO.Value = "PRODUCCION" Then
                                        ws_LIC.Value = "https://cr.edocnube.com/eDocCR/Sitios_Solsap/WS_LICENCIA_SAP/Licencia.svc"
                                    End If


                            End Select
                        End If


                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If Not pVal.Before_Action Then
                            Select Case pVal.ItemUID
                                Case "obtnGrabar"

                                    Try
                                        Dim oConfiguracion As Entidades.Configuracion
                                        Dim olistaDetalleConfiguracion As List(Of Entidades.ConfiguracionDetalle)

                                        oForm = rsboApp.Forms.Item("frmParametrosAddon")

                                        Dim cbxCO As SAPbouiCOM.ComboBox
                                        cbxCO = oForm.Items.Item("cbxCO").Specific 'Compania

                                        Dim cbxINH As SAPbouiCOM.ComboBox
                                        cbxINH = oForm.Items.Item("cbxINH").Specific 'INHOUSE

                                        Dim ws_LIC As SAPbouiCOM.EditText
                                        ws_LIC = oForm.Items.Item("ws_LIC").Specific 'Licencia WS

                                        Dim txtNIT As SAPbouiCOM.EditText
                                        txtNIT = oForm.Items.Item("txtNIT").Specific 'NIT,RUC,IDENTIFICACION

                                        'GrabaParametrizacion("01", "Factura de Proveedor", txtFPref.Value, txtFCue.Value, lCuentaF.Caption)
                                        oConfiguracion = New Entidades.Configuracion
                                        oConfiguracion.Modulo = Funciones_SAP.VariablesGlobales._gNombreAddOn
                                        oConfiguracion.Tipo = "PARAMETROS"
                                        oConfiguracion.SubTipo = "CONFIGURACION"
                                        olistaDetalleConfiguracion = New List(Of Entidades.ConfiguracionDetalle)

                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Param_Ambiente", cbxCO.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Param_Inhouse", cbxINH.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("WS_Licencia", ws_LIC.Value))
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("Param_Ruc", txtNIT.Value))

                                        oUserDataSource = oForm.DataSources.UserDataSources.Item("CHK_GLE")
                                        olistaDetalleConfiguracion.Add(New Entidades.ConfiguracionDetalle("ActivarQueSeGuardeLogEmision", oUserDataSource.ValueEx.ToString()))

                                        oConfiguracion.Detalle = olistaDetalleConfiguracion
                                        GuardaCONF(oConfiguracion)

                                        oForm.Items.Item("obtnGrabar").Visible = False
                                        oForm.Items.Item("2").Left = oForm.Items.Item("obtnGrabar").Left
                                        Dim oB As SAPbouiCOM.Button
                                        oB = oForm.Items.Item("2").Specific
                                        oB.Caption = "OK"

                                    Catch ex As Exception
                                        Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent - obtnGrabar: " + ex.Message.ToString(), "frmParametrosAddon")
                                    Finally

                                    End Try

                            End Select
                        End If
                End Select
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmParametrosAddon")
            System.Windows.Forms.MessageBox.Show("Error rSboApp_ItemEvent :" & ex.Message.ToString())
        End Try
    End Sub

    Public Sub GuardaCONF(ByVal oConfiguracion As Entidades.Configuracion)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim sqlstr As String = ""
        Dim mRst As SAPbobsCOM.Recordset = Nothing
        Dim conta As String = Nothing
        Dim sDocEntry As String = Nothing

        Try
            Dim query As String
            Dim CodeExist As String = "0"
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                query = "Select ""DocEntry"" From """ & rCompany.CompanyDB & """.""@SSUDOCONF"" Where ""U_Modulo"" = '" + oConfiguracion.Modulo + "' AND ""U_Tipo"" = '" + oConfiguracion.Tipo + "' AND ""U_Subtipo"" = '" + oConfiguracion.SubTipo + "'"
            Else
                query = "Select DocEntry From [@SSUDOCONF] Where U_Modulo = '" + oConfiguracion.Modulo + "' AND U_Tipo = '" + oConfiguracion.Tipo + "' AND U_Subtipo = '" + oConfiguracion.SubTipo + "'"
            End If
            CodeExist = oFuncionesB1.getRSvalue(query, "DocEntry")

            'mRst = fCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Not CodeExist = "0" Then ' SI EXISTE, ELIMINO Y ACTUALIZO

                ' SI EXISTE ELIMINA PARA VOLVER A CREAR
                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONF")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", CodeExist)
                oGeneralService.Delete(oGeneralParams)

                'CREA NUEVAMENTE EL REGISTRO
                oGeneralService = oCompanyService.GetGeneralService("SS_CONF")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralData.SetProperty("U_Modulo", oConfiguracion.Modulo)
                oGeneralData.SetProperty("U_Tipo", oConfiguracion.Tipo)
                oGeneralData.SetProperty("U_Subtipo", oConfiguracion.SubTipo)

                oChildren = oGeneralData.Child("SSUDOCONFD")
                For Each oItem As Entidades.ConfiguracionDetalle In oConfiguracion.Detalle
                    oChild = oChildren.Add
                    oChild.SetProperty("U_Nombre", oItem.Nombre)
                    oChild.SetProperty("U_Valor", oItem.Valor)
                Next
                oGeneralParams = oGeneralService.Add(oGeneralData)

            Else

                oCompanyService = rCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("SS_CONF")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralData.SetProperty("U_Modulo", oConfiguracion.Modulo)
                oGeneralData.SetProperty("U_Tipo", oConfiguracion.Tipo)
                oGeneralData.SetProperty("U_Subtipo", oConfiguracion.SubTipo)

                oChildren = oGeneralData.Child("SSUDOCONFD")
                For Each oItem As Entidades.ConfiguracionDetalle In oConfiguracion.Detalle
                    oChild = oChildren.Add
                    oChild.SetProperty("U_Nombre", oItem.Nombre)
                    oChild.SetProperty("U_Valor", oItem.Valor)
                Next
                oGeneralParams = oGeneralService.Add(oGeneralData)
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Function ConsultaParametro(ByVal Modulo As String, ByVal Tipo As String, ByVal Subtipo As String, ByVal Nombre As String) As String
        Try
            Dim valor As String = ""
            Dim sQueryPrefijo As String = ""
            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sQueryPrefijo = "SELECT A.""U_Valor"" "
                sQueryPrefijo += "FROM ""@SSUDOCONFD"" A INNER JOIN "
                sQueryPrefijo += """@SSUDOCONF"" B ON A.""DocEntry"" = B.""DocEntry"""
                sQueryPrefijo += " WHERE  B.""U_Modulo"" = '" + Modulo + "' AND B.""U_Tipo"" = '" + Tipo + "' "
                sQueryPrefijo += " AND B.""U_Subtipo"" = '" + Subtipo + "'"
                sQueryPrefijo += " AND A.""U_Nombre"" = '" + Nombre + "'"
            Else
                sQueryPrefijo = "SELECT A.U_Valor "
                sQueryPrefijo += "FROM ""@SSUDOCONFD"" A WITH(NOLOCK) INNER JOIN "
                sQueryPrefijo += """@SSUDOCONF"" B WITH(NOLOCK) ON A.DocEntry = B.DocEntry"
                sQueryPrefijo += " WHERE B.U_Modulo = '" + Modulo + "' AND  B.U_Tipo = '" + Tipo + "' "
                sQueryPrefijo += " AND B.U_Subtipo = '" + Subtipo + "'"
                sQueryPrefijo += " AND A.U_Nombre = '" + Nombre + "'"
            End If

            valor = oFuncionesB1.getRSvalue(sQueryPrefijo, "U_Valor", "")
            Return valor
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function RecorreFormulario(ByVal oApp As SAPbouiCOM.Application, ByVal Formulario As String) As Boolean
        Try
            For Each oForm In oApp.Forms
                Select Case oForm.UniqueID
                    Case Formulario
                        oForm.Visible = True
                        oForm.Select()
                        Return True
                End Select
            Next

            For Each oForm In oApp.Forms
                If oForm.UniqueID = Formulario Then
                    oForm.Visible = True
                    oForm.Select()
                    ' oForm.Close()
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Class
