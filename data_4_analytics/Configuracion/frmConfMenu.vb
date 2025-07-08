Public Class frmConfMenu

    Private oForm As SAPbouiCOM.Form
    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioMenuDeConfiguraciones()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmConfMenu") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmConfMenu.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmConfMenu").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try
            oForm = rsboApp.Forms.Item("frmConfMenu")
            oForm.Visible = True
            oForm.Select()
        Catch ex As Exception
            rsboApp.MessageBox("Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rSboApp.ItemEvent
        Try
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK _
                             And pVal.FormTypeEx = "frmConfMenu" Then
                If pVal.BeforeAction = False And pVal.ItemUID = "linkParam" Then

                    ofrmParametrosAddon.CargaFormularioParametrosADDON()

                    'ElseIf pVal.BeforeAction = False And pVal.ItemUID = "linkProx" Then

                    '    ofrmProxy.CargaFormularioParametrosProxy()

                    'ElseIf pVal.BeforeAction = False And pVal.ItemUID = "linkCon" Then


                    '    ofrmconsultasDB.CargaFormularioParametrosConsultasBD()

                End If
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmConfMenu")
            System.Windows.Forms.MessageBox.Show("Error rSboApp_ItemEvent :" & ex.Message.ToString())
        End Try

    End Sub


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
