Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Security.Permissions
Imports System.Windows.Forms
Imports SAPbouiCOM
Imports System.Globalization
Imports Microsoft.Office.Interop
Imports SAPbobsCOM

Public Class frmCreaTercero
    Private oForm As SAPbouiCOM.Form
    Private WithEvents rsboApp As SAPbouiCOM.Application
    Private rCompany As SAPbobsCOM.Company
    Private codigosEliminados As New List(Of String)

    Public Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargarFormulario()
        Try
            Dim xmlDoc As New XmlDocument
            Dim path As String = System.Windows.Forms.Application.StartupPath & "\frmCreaTercero.srf"
            xmlDoc.Load(path)

            rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            oForm = rsboApp.Forms.Item("frmCreaTercero")

            oForm.Freeze(True)
            oForm.Mode = BoFormMode.fm_OK_MODE

            Dim oDT As DataTable = oForm.DataSources.DataTables.Add("DT_TERCEROS")
            oDT.ExecuteQuery("SELECT ""Code"", ""Name"", ""U_IdTercero"", ""U_NomTercero"" FROM ""@SSTERCEROS""")

            Dim oGrid As Grid = oForm.Items.Item("tb_ter").Specific
            oGrid.DataTable = oDT
            oGrid.SelectionMode = BoMatrixSelect.ms_Single

            For i As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(i).Editable = True
            Next

            Dim oRS As Recordset = rCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRS.DoQuery("SELECT ""Code"" FROM ""@SSTERCEROS""")
            Dim codigosExistentes As New HashSet(Of String)
            While Not oRS.EoF
                codigosExistentes.Add(oRS.Fields.Item(0).Value.ToString())
                oRS.MoveNext()
            End While

            oGrid.Columns.Item("Code").Editable = True

            oGrid.Columns.Item("Code").TitleObject.Caption = "Code"
            oGrid.Columns.Item("Name").TitleObject.Caption = "Name"
            oGrid.Columns.Item("U_IdTercero").TitleObject.Caption = "ID Tercero"
            oGrid.Columns.Item("U_NomTercero").TitleObject.Caption = "Nombre Tercero"

            oForm.Visible = True
            oForm.Freeze(False)
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error al cargar el formulario: " & ex.Message, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub GuardarCambios()
        Try
            Dim oDT As DataTable = oForm.DataSources.DataTables.Item("DT_TERCEROS")
            Dim oRS As Recordset = rCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            oRS.DoQuery("SELECT ""Code"" FROM ""@SSTERCEROS""")
            Dim codigosExistentes As New HashSet(Of String)
            While Not oRS.EoF
                codigosExistentes.Add(oRS.Fields.Item(0).Value.ToString())
                oRS.MoveNext()
            End While

            Dim nuevosCodigos As New HashSet(Of String)

            For i As Integer = 0 To oDT.Rows.Count - 1
                Dim code = oDT.GetValue("Code", i).ToString().Trim()
                Dim name = oDT.GetValue("Name", i).ToString().Trim()
                Dim id = oDT.GetValue("U_IdTercero", i).ToString().Trim()
                Dim nombre = oDT.GetValue("U_NomTercero", i).ToString().Trim()

                If String.IsNullOrEmpty(code) OrElse String.IsNullOrEmpty(id) OrElse String.IsNullOrEmpty(nombre) Then
                    Continue For
                End If

                nuevosCodigos.Add(code)

                If codigosExistentes.Contains(code) Then
                    Try
                        Dim query As String = $"UPDATE ""@SSTERCEROS"" SET ""Name"" = '{name}', ""U_IdTercero"" = '{id}', ""U_NomTercero"" = '{nombre}' WHERE ""Code"" = '{code}'"
                        oRS.DoQuery(query)
                    Catch ex As Exception
                        rsboApp.StatusBar.SetText("Error actualizando tercero [" & code & "]: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    End Try
                    'oRS.DoQuery($"UPDATE ""@SSTERCEROS"" SET ""Name"" = '{name}', ""U_IdTercero"" = '{id}', ""U_NomTercero"" = '{nombre}' WHERE ""Code"" = '{code}'")
                Else
                    Dim oUserTable As UserTable = rCompany.UserTables.Item("SSTERCEROS")
                    oUserTable.Code = code
                    oUserTable.Name = name
                    oUserTable.UserFields.Fields.Item("U_IdTercero").Value = id
                    oUserTable.UserFields.Fields.Item("U_NomTercero").Value = nombre
                    Dim result = oUserTable.Add()
                    If result <> 0 Then
                        Dim errMsg As String = ""
                        rCompany.GetLastError(result, errMsg)
                        Throw New Exception("Error al insertar tercero: " & errMsg)
                    End If
                End If
            Next

            For Each cod In codigosEliminados
                If codigosExistentes.Contains(cod) Then
                    oRS.DoQuery($"DELETE FROM ""@SSTERCEROS"" WHERE ""Code"" = '{cod}'")
                End If
            Next
            codigosEliminados.Clear()

            rsboApp.StatusBar.SetText("Cambios guardados correctamente.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
            RefrescarGrid()
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error al guardar cambios: " & ex.Message, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub RefrescarGrid()
        Try
            Dim oDT As DataTable = oForm.DataSources.DataTables.Item("DT_TERCEROS")
            oDT.ExecuteQuery("SELECT ""Code"", ""Name"", ""U_IdTercero"", ""U_NomTercero"" FROM ""@SSTERCEROS""")
            CType(oForm.Items.Item("tb_ter").Specific, Grid).DataTable = oDT
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error al refrescar grid: " & ex.Message, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub rsboApp_ItemEvent(FormUID As String, ByRef pVal As ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormTypeEx <> "frmCreaTercero" Then Return

            If pVal.EventType = BoEventTypes.et_ITEM_PRESSED AndAlso Not pVal.BeforeAction Then
                Select Case pVal.ItemUID
                    Case "1" ' Grabar
                        GuardarCambios()
                    Case "2" ' Cancelar
                        oForm.Close()
                    Case "btn_addLn"
                        Dim oGrid As Grid = oForm.Items.Item("tb_ter").Specific
                        oGrid.DataTable.Rows.Add()
                        If oForm.Mode = BoFormMode.fm_OK_MODE Then
                            oForm.Mode = BoFormMode.fm_UPDATE_MODE
                        End If
                    Case "btn_delLn"
                        Dim oGrid As Grid = oForm.Items.Item("tb_ter").Specific
                        If oGrid.Rows.SelectedRows.Count > 0 Then
                            Dim index As Integer = oGrid.Rows.SelectedRows.Item(0, BoOrderType.ot_SelectionOrder)
                            Dim code As String = oGrid.DataTable.GetValue("Code", index).ToString().Trim()

                            If Not String.IsNullOrEmpty(code) Then
                                codigosEliminados.Add(code)
                            End If

                            oGrid.DataTable.Rows.Remove(index)

                            If oForm.Mode = BoFormMode.fm_OK_MODE Then
                                oForm.Mode = BoFormMode.fm_UPDATE_MODE
                            End If
                        Else
                            rsboApp.StatusBar.SetText("Debe seleccionar una fila para eliminar.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                        End If
                End Select
            End If

            If pVal.EventType = BoEventTypes.et_CLICK AndAlso pVal.ItemUID = "tb_ter" AndAlso pVal.ColUID = "Code" Then
                Dim oGrid As Grid = oForm.Items.Item("tb_ter").Specific
                Dim rowIndex As Integer = pVal.Row
                Dim currentCode As String = oGrid.DataTable.GetValue("Code", rowIndex).ToString().Trim()

                If Not String.IsNullOrEmpty(currentCode) Then
                    BubbleEvent = False
                    rsboApp.StatusBar.SetText("No se puede modificar el código de un registro existente.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                End If
            End If
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error en evento: " & ex.Message, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class
