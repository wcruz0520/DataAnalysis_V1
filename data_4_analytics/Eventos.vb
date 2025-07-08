Imports System.IO

Public Class Eventos
    Public WithEvents rSboApp As SAPbouiCOM.Application

    Sub New()
        rSboApp = rSboGui.GetApplication
    End Sub

    ''' <summary>
    '''  Eventos de Aplicacion
    ''' </summary>
    ''' <param name="EventType"></param>
    ''' <remarks></remarks>
    Private Sub rSboApp_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles rSboApp.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    End
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                    End
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    End
                Case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged
                    End
            End Select

        Catch ex As Exception
        End Try
    End Sub

    Private Sub rSboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.MenuEvent
        Try
            '1284
            If pVal.MenuUID = "UDO2" And pVal.BeforeAction = False Then
                ' Acerca de
                ofrmAcercaDe.CargaFormularioAcercaDe()

            ElseIf pVal.MenuUID = "UDO1" And pVal.BeforeAction = False Then
                ofrmPlantilla.CargaFormularioPlantilla()

            ElseIf pVal.MenuUID = "UDO3" And pVal.BeforeAction = False Then
                ofrmPlantillaRPT.CargaFormularioPlantillaRPT()

            ElseIf pVal.MenuUID = "UDO7" And pVal.BeforeAction = False Then
                ofrmCartera.CargaFormCartera()

            ElseIf pVal.MenuUID = "UDO9" And pVal.BeforeAction = False Then
                ofrmEstadistico.CargaFormEstadistico()

            ElseIf pVal.MenuUID = "GS121" And pVal.BeforeAction = False Then
                'Historico
                'ofrmHistorico.CargaFormularioHistorico()
            ElseIf pVal.MenuUID = "GS131" And pVal.BeforeAction = False Then
                'Configuración de Parametros
                'ofrmAcercaDe.CargaFormularioAcercaDe()
                'ofrmConfiguracion.CargaFormularioConfiguracion()
                'rSboApp.MessageBox("fbsfndfnfd") ' mensaje para validar 

                'ElseIf pVal.MenuUID = "MNU_TERCEROS" AndAlso pVal.BeforeAction = False Then
                '    ' Abre el formulario del UDO SSTERCEROS
                '    rSboApp.ActivateMenuItem("FT_SSTERCEROS")

            End If

            'If pVal.MenuUID = "1282" Or pVal.MenuUID = "1287" And pVal.BeforeAction = False Then ' NUEVO, DUPLICAR
            '    Try
            '        Dim typeExx, idFormm As String
            '        typeExx = oFuncionesB1.FormularioActivo(idFormm)
            '        If typeExx = "141" Then ' FACTURA DE PROEVEEDORES
            '            If Not pVal.BeforeAction Then
            '                Try
            '                    Dim mForm As SAPbouiCOM.Form = rSboApp.Forms.Item(idFormm)
            '                    mForm.Items.Item("txtFE").Visible = False
            '                    mForm.Items.Item("LinkFE").Visible = False
            '                Catch ex As Exception
            '                End Try
            '            End If

            '        End If
            '    Catch ex As Exception
            '    End Try
            'End If

        Catch ex As Exception
            rSboApp.MessageBox(ex.Message)
        End Try
    End Sub

End Class
