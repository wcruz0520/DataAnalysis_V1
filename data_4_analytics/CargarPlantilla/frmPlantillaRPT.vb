Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Security.Permissions
Imports System.Windows.Forms
Imports SAPbouiCOM
Imports System.Globalization
Imports Microsoft.Office.Interop


Public Class frmPlantRPT

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Dim lineaMatrix As Integer = 0
    Dim customCulture As CultureInfo

    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
    Dim oConditions As SAPbouiCOM.Conditions
    Dim oCondition As SAPbouiCOM.Condition
    Dim oUserDataSource As SAPbouiCOM.UserDataSource

    Dim columnasEsperadas As Integer = 12

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioPlantillaRPT()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmPlantRpt") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmPlantRpt.srf" 'frmPlantRpt
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)

            Catch exx As Exception
                rsboApp.Forms.Item("frmPlantRpt").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmPlantRpt")
            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = Windows.Forms.Application.StartupPath & "\LogoSS.png"

            oForm.Mode = BoFormMode.fm_ADD_MODE

            Dim txtCodSuc As SAPbouiCOM.EditText
            txtCodSuc = oForm.Items.Item("txtCodInf").Specific
            txtCodSuc.DataBind.SetBound(True, "@SSPLANT_RPT", "U_CODE_RPT")

            txtCodSuc = oForm.Items.Item("txtNomInf").Specific
            txtCodSuc.DataBind.SetBound(True, "@SSPLANT_RPT", "U_NAME_RPT")

            txtCodSuc = oForm.Items.Item("txtAnexo").Specific
            txtCodSuc.DataBind.SetBound(True, "@SSPLANT_RPT", "U_ANEXO")

            'Instancia elementos consultar linea
            Dim lbl As SAPbouiCOM.StaticText = oForm.Items.Item("Item_0").Specific
            lbl.Item.Visible = False

            Dim txtBL As SAPbouiCOM.EditText = oForm.Items.Item("txtBL").Specific
            txtBL.Item.Visible = False

            Dim btnBusL As SAPbouiCOM.Button = oForm.Items.Item("btnBusL").Specific
            btnBusL.Item.Visible = False


            Dim screenWidth As Integer = rsboApp.Desktop.Width
            Dim screenHeight As Integer = rsboApp.Desktop.Height

            ' Obtener el tamaño del formulario
            Dim formWidth As Integer = oForm.Width
            Dim formHeight As Integer = oForm.Height

            ' Calcular la posición para centrar el formulario
            Dim centeredLeft As Integer = (screenWidth - formWidth) / 2
            Dim centeredTop As Integer = (screenHeight - formHeight) / 2

            ' Ajustar la posición del formulario
            oForm.Left = centeredLeft
            oForm.Top = centeredTop - 75

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

            'rsboApp.MessageBox(NombreAddon + " Pantalla cargada con exito: ")

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                Utilitario.Util_Log.Escribir_Log("evento click", "frmPlantRpt")
                If pVal.FormTypeEx = "frmPlantRpt" Then

                    Utilitario.Util_Log.Escribir_Log("evento click frmPlantRpt ", "frmPlantRpt")
                    If pVal.BeforeAction = False And pVal.ItemUID = "btnMas" Then
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantRpt ItemUID = btnMas", "frmPlantRpt")
                        rsboApp.Forms.ActiveForm.Freeze(True)

                        Dim oMatrix As SAPbouiCOM.Matrix = CType(oForm.Items.Item("MTX_UDOrpt").Specific, SAPbouiCOM.Matrix)

                        If oMatrix.RowCount = 0 Then

                            oMatrix.AddRow()
                            For i As Integer = 1 To oMatrix.RowCount
                                oMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i
                            Next

                        Else

                            AddRowAtSelectedPosition("frmPlantRpt", oMatrix, lineaMatrix)

                        End If
                        rsboApp.Forms.ActiveForm.Freeze(False)

                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btnMenos" Then
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantRpt ItemUID = btnMenos", "frmPlantRpt")
                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_UDOrpt").Specific

                        If mMatrix.RowCount > 0 Then
                            mMatrix.DeleteRow(lineaMatrix)
                        End If

                        For i As Integer = 1 To mMatrix.RowCount
                            mMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i
                        Next

                        If rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_OK_MODE Then
                            rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Mode = BoFormMode.fm_UPDATE_MODE
                        End If

                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btnCP" Then
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantRpt ItemUID = btnCP", "frmPlantRpt")
                        Dim selectFileDialog As New SelectFileDialog("C:\", "", "CSV files (*.csv)|*.csv|All files (*.*)|*.*", DialogTypeRPT.OPEN)
                        selectFileDialog.Open()

                        If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFile) Then
                            Dim ruta As String = ""
                            Dim validacion As Integer
                            ruta = selectFileDialog.SelectedFile

                            validacion = ValidarCSV(ruta, columnasEsperadas)

                            Select Case validacion
                                Case 0
                                    rsboApp.StatusBar.SetText("✅ El archivo es válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    leerCSV(ruta)
                                Case 1
                                    rsboApp.StatusBar.SetText(" Error: El archivo no está delimitado por ';'. Cod. Error: {ValidarCSV(ruta, columnasEsperadas).ToString}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Case 2
                                    rsboApp.StatusBar.SetText($" Error: Se esperaban {columnasEsperadas} columnas, pero se encontraron menos o más columnas. Cod. Error: {validacion}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Case 3
                                    rsboApp.StatusBar.SetText($" Archivo vacío. Cod. Error: {validacion}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Case Else
                                    rsboApp.StatusBar.SetText($" Error desconocido o archivo vacío. Cod. Error: {validacion}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Select


                        End If

                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "1" Then
                        If oForm.Mode = BoFormMode.fm_FIND_MODE Then

                        Else
                            oForm.EnableMenu("1283", True)
                        End If

                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "MTX_UDOrpt" Then
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantRpt ItemUID = MTX_UDOrpt", "frmPlantRpt")
                        Dim selectedRow As Integer = pVal.Row
                        lineaMatrix = pVal.Row
                        If selectedRow > 0 Then
                            rsboApp.StatusBar.SetText("Fila seleccionada: " & selectedRow.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If

                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btnBusL" Then
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantRpt ItemUID = btnBusL", "frmPlantRpt")
                        Dim BuscarLinea As SAPbouiCOM.EditText = oForm.Items.Item("txtBL").Specific

                        Dim LineaSelect As Integer = CInt(BuscarLinea.Value)
                        Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_UDOrpt").Specific

                        oForm.Freeze(True)
                        mMatrix.FlushToDataSource() ' Deshabilita la actualización en tiempo real.
                        mMatrix.SelectRow(LineaSelect, True, False) ' Selecciona la fila.
                        oForm.Freeze(False)
                    End If
                End If
            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                Utilitario.Util_Log.Escribir_Log("evento et_ITEM_PRESSED", "frmPlantRpt")
                If pVal.FormTypeEx = "frmPlantRpt" Then

                    If pVal.BeforeAction = False And pVal.ItemUID = "chk1FC" Then
                        Dim chk1FC As SAPbouiCOM.CheckBox = oForm.Items.Item("chk1FC").Specific

                        Dim oMatrix As SAPbouiCOM.Matrix
                        oMatrix = oForm.Items.Item("MTX_UDOrpt").Specific

                        If chk1FC.Checked Then
                            Dim oColumns As SAPbouiCOM.Columns
                            oColumns = oMatrix.Columns

                            oColumns.Item("1FC").Editable = True
                        Else
                            Dim oColumns As SAPbouiCOM.Columns
                            oColumns = oMatrix.Columns

                            oColumns.Item("1FC").Editable = False
                        End If
                    End If

                    Dim oFor As SAPbouiCOM.Form
                    oFor = rsboApp.Forms.Item("frmPlantRpt")
                    If pVal.ItemUID = "btnCP" And oFor.Mode <> BoFormMode.fm_ADD_MODE Then 'pVal.BeforeAction = False And
                        'Dim btnCP As SAPbouiCOM.Button = oForm.Items.Item("btnCP").Specific
                        oFor.Mode = BoFormMode.fm_UPDATE_MODE
                    End If

                End If

            End If

            If pVal.FormTypeEx = "frmPlantRpt" Then

                If pVal.BeforeAction And pVal.ItemUID = "1" Then

                    Try
                        Dim oFor As SAPbouiCOM.Form
                        oFor = rsboApp.Forms.Item("frmPlantRpt")
                        Dim CodInform As SAPbouiCOM.EditText = oFor.Items.Item("txtCodInf").Specific


                        If IsNothing(CodInform.Value) Then
                            If oFor.Mode = BoFormMode.fm_ADD_MODE Then
                                BubbleEvent = False
                                rsboApp.StatusBar.SetText("Por favor Código de Informe", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If


                        Else
                            If oFor.Mode = BoFormMode.fm_ADD_MODE Then

                                Dim txtCodInf As SAPbouiCOM.EditText = oForm.Items.Item("txtCodInf").Specific
                                Dim CodInfor As String = IIf(String.IsNullOrEmpty(txtCodInf.Value), "", txtCodInf.Value)

                                Dim txtAnexo As SAPbouiCOM.EditText = oForm.Items.Item("txtAnexo").Specific
                                Dim Anexo As String = IIf(String.IsNullOrEmpty(txtAnexo.Value), "", txtAnexo.Value)
                                Dim ConRegistro As String = ""

                                If CodInfor = "" Then
                                    BubbleEvent = False
                                    rsboApp.StatusBar.SetText("No Olvidar colocar Código de Informe..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Else

                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        If Anexo = "" Then
                                            ' Anexo vacío: validar si existe uno con Anexo vacío
                                            ConRegistro = "SELECT COUNT(1) AS ""Contador"" FROM """ + rCompany.CompanyDB + """.""@SSPLANT_RPT"" " &
                                                          "WHERE ""U_CODE_RPT"" = '" + CodInfor.ToString + "' AND IFNULL(""U_ANEXO"", '') = ''"
                                        Else
                                            ' Anexo lleno: validar si existe uno igual
                                            ConRegistro = "SELECT COUNT(1) AS ""Contador"" FROM """ + rCompany.CompanyDB + """.""@SSPLANT_RPT"" " &
                                                          "WHERE ""U_CODE_RPT"" = '" + CodInfor.ToString + "' AND IFNULL(""U_ANEXO"", '') = '" + Anexo.ToString + "'"
                                        End If
                                    Else
                                        If Anexo = "" Then
                                            ConRegistro = "SELECT COUNT(1) AS ""Contador"" FROM ""@SSPLANT_RPT"" " &
                                                          "WHERE ""U_CODE_RPT"" = '" + CodInfor.ToString + "' AND ISNULL(""U_ANEXO"", '') = ''"
                                        Else
                                            ConRegistro = "SELECT COUNT(1) AS ""Contador"" FROM ""@SSPLANT_RPT"" " &
                                                          "WHERE ""U_CODE_RPT"" = '" + CodInfor.ToString + "' AND ISNULL(""U_ANEXO"", '') = '" + Anexo.ToString + "'"
                                        End If
                                    End If

                                    Utilitario.Util_Log.Escribir_Log("Consulta cantidad registro por Cod. Informe: " + ConRegistro.ToString, "frmPlantRpt")

                                    Dim Val As Integer = CInt(oFuncionesB1.getRSvalue(ConRegistro, "Contador", "0"))
                                    If Val > 0 Then
                                        BubbleEvent = False
                                        rsboApp.StatusBar.SetText("Solo se permite un registro por Código de Informe y Anexo/#Hoja..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If

                                End If

                            Else
                                If oFor.Mode = BoFormMode.fm_UPDATE_MODE Then

                                    Dim txtCodInf As SAPbouiCOM.EditText = oForm.Items.Item("txtCodInf").Specific
                                    Dim CodInfor As String = IIf(String.IsNullOrEmpty(txtCodInf.Value), "", txtCodInf.Value)

                                    Dim txtAnexo As SAPbouiCOM.EditText = oForm.Items.Item("txtAnexo").Specific
                                    Dim Anexo As String = IIf(String.IsNullOrEmpty(txtAnexo.Value), "", txtAnexo.Value)
                                    Dim ConRegistro As String = ""

                                    If CodInfor = "" Then
                                        BubbleEvent = False
                                        rsboApp.StatusBar.SetText("No Olvidar colocar Código de Informe..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Else

                                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            If Anexo = "" Then
                                                ' Anexo vacío: validar si existe uno con Anexo vacío
                                                ConRegistro = "SELECT COUNT(1) AS ""Contador"" FROM """ + rCompany.CompanyDB + """.""@SSPLANT_RPT"" " &
                                                              "WHERE ""U_CODE_RPT"" = '" + CodInfor.ToString + "' AND IFNULL(""U_ANEXO"", '') = ''"
                                            Else
                                                ' Anexo lleno: validar si existe uno igual
                                                ConRegistro = "SELECT COUNT(1) AS ""Contador"" FROM """ + rCompany.CompanyDB + """.""@SSPLANT_RPT"" " &
                                                              "WHERE ""U_CODE_RPT"" = '" + CodInfor.ToString + "' AND IFNULL(""U_ANEXO"", '') = '" + Anexo.ToString + "'"
                                            End If
                                        Else
                                            If Anexo = "" Then
                                                ConRegistro = "SELECT COUNT(1) AS ""Contador"" FROM ""@SSPLANT_RPT"" " &
                                                              "WHERE ""U_CODE_RPT"" = '" + CodInfor.ToString + "' AND ISNULL(""U_ANEXO"", '') = ''"
                                            Else
                                                ConRegistro = "SELECT COUNT(1) AS ""Contador"" FROM ""@SSPLANT_RPT"" " &
                                                              "WHERE ""U_CODE_RPT"" = '" + CodInfor.ToString + "' AND ISNULL(""U_ANEXO"", '') = '" + Anexo.ToString + "'"
                                            End If
                                        End If

                                        Utilitario.Util_Log.Escribir_Log("Consulta cantidad registro por Cod. Informe: " + ConRegistro.ToString, "frmPlantRpt")

                                        Dim Val As Integer = CInt(oFuncionesB1.getRSvalue(ConRegistro, "Contador", "0"))
                                        If Val > 0 Then
                                            BubbleEvent = False
                                            rsboApp.StatusBar.SetText("Solo se permite un registro por Código de Informe y Anexo/#Hoja..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        End If

                                    End If


                                End If

                            End If

                        End If



                    Catch ex As Exception
                        BubbleEvent = False
                        rsboApp.StatusBar.SetText("Error al validar informacion: " + ex.Message.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End Try

                End If



            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD Then
                Utilitario.Util_Log.Escribir_Log("evento et_MATRIX_LOAD", "frmPlantRpt")
                If pVal.FormTypeEx = "frmPlantRpt" Then
                    If pVal.BeforeAction Then
                        Dim EditText As SAPbouiCOM.EditText
                        EditText = oForm.Items.Item("Item_4").Specific
                        EditText.Item.Click()

                        EditText = oForm.Items.Item("Item_10").Specific
                        EditText.Item.Enabled = False
                    End If
                End If





            End If

            If pVal.EventType = BoEventTypes.et_KEY_DOWN Then
                Utilitario.Util_Log.Escribir_Log("evento et_KEY_DOWN", "frmPlantRpt")
                If pVal.FormTypeEx = "frmPlantRpt" Then

                    If pVal.BeforeAction = False Then
                        If pVal.CharPressed = Keys.Enter Then

                            Dim BuscarLinea As SAPbouiCOM.EditText = oForm.Items.Item("txtBL").Specific

                            Dim LineaSelect As Integer = CInt(BuscarLinea.Value)
                            Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_UDOrpt").Specific

                            mMatrix.SelectRow(LineaSelect, True, False)

                        End If
                    End If

                End If

            End If

        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmPlantRpt")
        End Try


    End Sub

    Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent
        Try
            If pVal.MenuUID = "1282" Then 'Buscar
                If pVal.BeforeAction = False Then
                    If rsboApp.Forms.ActiveForm.UniqueID = "frmPlantRpt" Then
                        Try
                            Dim NumeroUdo As SAPbouiCOM.EditText = oForm.Items.Item("Item_10").Specific
                            NumeroUdo.Item.Enabled = True

                        Catch ex As Exception

                        End Try
                    End If

                End If
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error en MenuEvent: " + ex.Message.ToString, "frmPlantRpt")
        End Try

    End Sub

    Private Sub rsboApp_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent

        Try
            If eventInfo.FormUID = "frmPlantRpt" Then

                If eventInfo.ItemUID = "MTX_UDOrpt" Then

                    If eventInfo.ColUID = "COL" Then

                        Dim oMenuItem As SAPbouiCOM.MenuItem = Nothing
                        Dim oMenus As SAPbouiCOM.Menus = Nothing

                        If eventInfo.BeforeAction = True Then

                            Dim mForm As SAPbouiCOM.Form = rsboApp.Forms.Item(eventInfo.FormUID)

                            If mForm.Mode = BoFormMode.fm_ADD_MODE Or mForm.Mode = BoFormMode.fm_OK_MODE Or mForm.Mode = BoFormMode.fm_FIND_MODE Then
                                oMenuItem = rsboApp.Menus.Item("1280")
                                If oMenuItem.SubMenus.Exists("Agregar") Then
                                    rsboApp.Menus.RemoveEx("Agregar")

                                End If
                                If oMenuItem.SubMenus.Exists("Eliminar") Then
                                    rsboApp.Menus.RemoveEx("Eliminar")
                                End If
                                If oMenuItem.SubMenus.Exists("1283") Then
                                    rsboApp.Menus.RemoveEx("1283")
                                End If
                            End If
                        Else

                            Try
                                oMenuItem = rsboApp.Menus.Item("1280")
                                If oMenuItem.SubMenus.Exists("Agregar") Then
                                    rsboApp.Menus.RemoveEx("Agregar")

                                End If
                                If oMenuItem.SubMenus.Exists("Eliminar") Then
                                    rsboApp.Menus.RemoveEx("Eliminar")
                                End If
                                If oMenuItem.SubMenus.Exists("1283") Then
                                    rsboApp.Menus.RemoveEx("1283")
                                End If
                            Catch ex As Exception
                                rsboApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End Try

                        End If

                    End If

                End If

            End If
        Catch ex As Exception

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

    Private Sub AddRowAtSelectedPosition(ByVal FormUID As String, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal SelectedRow As Integer)

        Try

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item(FormUID)
            'Dim oMatrix As SAPbouiCOM.Matrix = CType(oForm.Items.Item(MatrixID).Specific, SAPbouiCOM.Matrix)
            'oForm.Freeze(True)
            ' Validar la fila seleccionada
            If SelectedRow <= 0 Or SelectedRow > oMatrix.RowCount Then
                rsboApp.StatusBar.SetText("Seleccione una fila válida.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If

            ' Insertar una nueva fila en la posición seleccionada
            oMatrix.AddRow(1, SelectedRow)

            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SSPLANT_RPT_DET1")
            oMatrix.FlushToDataSource()

            Dim LineaSSetear As Integer = SelectedRow


            'oDataSource.InsertRecord(SelectedRow - 1)
            oDataSource.SetValue("LineId", LineaSSetear, SelectedRow + 1)
            oDataSource.SetValue("U_TITULO1", LineaSSetear, String.Empty)
            'oDataSource.SetValue("U_NIVEL1", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_TITULO2", LineaSSetear, String.Empty)
            'oDataSource.SetValue("U_NIVEL2", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_TITULO3", LineaSSetear, String.Empty)
            'oDataSource.SetValue("U_NIVEL3", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_TITULO4", LineaSSetear, String.Empty)
            'oDataSource.SetValue("U_NIVEL4", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_TITULO5", LineaSSetear, String.Empty)
            'oDataSource.SetValue("U_NIVEL5", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_TITULO6", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_NIVEL6", LineaSSetear, String.Empty)

            oDataSource.SetValue("U_PROYECTO", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_LINEANEGOCIO", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_SUBLINEA", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_SUCURSAL", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_AREA_ESPACIO", LineaSSetear, String.Empty)

            Dim numLineaInsert = SelectedRow + 1

            For i As Integer = numLineaInsert To oDataSource.Size - 1
                oDataSource.SetValue("LineId", i, (i).ToString())
            Next

            oMatrix.LoadFromDataSource()

            For i As Integer = SelectedRow To oMatrix.RowCount
                oMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i.ToString()
            Next

            oMatrix.SelectRow(SelectedRow, True, False)
            oForm.Update()
            'oForm.Freeze(False)
            rsboApp.StatusBar.SetText("Fila añadida correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            rsboApp.StatusBar.SetText("Error al añadir la fila: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            rsboApp.Forms.ActiveForm.Freeze(False)

        End Try
    End Sub

    Function leerCSV(ByVal ruta As String) As Boolean
        Try

            rsboApp.StatusBar.SetText("Leyendo archivo csv, por favor esperar un momento..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)
            customCulture = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
            customCulture.NumberFormat.NumberDecimalSeparator = "."
            customCulture.NumberFormat.NumberGroupSeparator = ","
            Utilitario.Util_Log.Escribir_Log("Cultura configurada: Separador decimal= " & customCulture.NumberFormat.NumberDecimalSeparator & " Separador de Miles= " & customCulture.NumberFormat.NumberGroupSeparator, "frmPlantRpt")


            ' Suponiendo que tienes una referencia a la Matrix
            Dim oMatrix As Matrix = oForm.Items.Item("MTX_UDOrpt").Specific
            'Dim source As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_DETUDO")
            Dim fil As Integer = 0
            ' Ruta al archivo CSV
            Dim csvPath As String = ruta
            Dim DatosEnExcel As New List(Of String)
            Using reader As New StreamReader(csvPath, Encoding.GetEncoding("iso-8859-1"))
                Dim rowIndex As Integer = 1

                If rowIndex = 1 Then
                    reader.ReadLine()
                End If

                While Not reader.EndOfStream
                    ' Leer cada línea del archivo CSV
                    Dim line As String = reader.ReadLine()
                    'Dim fields As String = line.Split(";"c)
                    DatosEnExcel.Add(line)
                    rowIndex += 1
                End While
            End Using

            Dim oProgressBar As SAPbouiCOM.ProgressBar
            oProgressBar = rsboApp.StatusBar.CreateProgressBar("Procesando CSV..", DatosEnExcel.Count, False)

            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SSPLANT_RPT_DET1")

            oDataSource.Clear()
            oMatrix.Clear()

            Dim j As Integer = 0

            Dim TotalLineas = DatosEnExcel.Count
            Dim ValorUpdate = TotalLineas / 100
            For Each filas In DatosEnExcel
                oDataSource.InsertRecord(j)
                oDataSource.SetValue("LineId", j, j + 1)
                oDataSource.SetValue("U_TITULO1", j, filas.Split(";")(0))
                'oDataSource.SetValue("U_NIVEL1", j, filas.Split(";")(1))
                oDataSource.SetValue("U_TITULO2", j, filas.Split(";")(1))
                'oDataSource.SetValue("U_NIVEL2", j, filas.Split(";")(3))
                oDataSource.SetValue("U_TITULO3", j, filas.Split(";")(2))
                'oDataSource.SetValue("U_NIVEL3", j, filas.Split(";")(5))
                oDataSource.SetValue("U_TITULO4", j, filas.Split(";")(3))
                'oDataSource.SetValue("U_NIVEL4", j, filas.Split(";")(7))
                oDataSource.SetValue("U_TITULO5", j, filas.Split(";")(4))
                'oDataSource.SetValue("U_NIVEL5", j, filas.Split(";")(9))
                oDataSource.SetValue("U_TITULO6", j, filas.Split(";")(5))
                oDataSource.SetValue("U_NIVEL6", j, filas.Split(";")(6))

                oDataSource.SetValue("U_PROYECTO", j, filas.Split(";")(7))
                oDataSource.SetValue("U_LINEANEGOCIO", j, filas.Split(";")(8))
                oDataSource.SetValue("U_SUBLINEA", j, filas.Split(";")(9))
                oDataSource.SetValue("U_SUCURSAL", j, filas.Split(";")(10))
                oDataSource.SetValue("U_AREA_ESPACIO", j, filas.Split(";")(11))



                oProgressBar.Value = j


                j += 1
            Next
            oMatrix.LoadFromDataSource()

            oForm.Freeze(False)
            oProgressBar.Stop()
            rsboApp.StatusBar.SetText("Datos cargados desde el archivo Excel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Return True

        Catch ex As Exception
            rsboApp.StatusBar.SetText(" Error al Leer archivo .cvs, " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
            Return False
        Finally

        End Try

    End Function

    Function ValidarCSV(rutaArchivo As String, columnasEsperadas As Integer) As Integer
        Try
            Using sr As New StreamReader(rutaArchivo)
                Dim linea As String = sr.ReadLine()

                If String.IsNullOrEmpty(linea) Then Return 3 ' Archivo vacío

                ' Validar si la línea contiene al menos un ';'
                If Not linea.Contains(";"c) Then Return 1 ' No está delimitado por ';'

                ' Separar la línea por ';' y contar columnas
                Dim columnas() As String = linea.Split(";"c)

                ' Validar cantidad de columnas
                If columnas.Length <> columnasEsperadas Then Return 2 ' Número incorrecto de columnas

                Return 0 ' Archivo válido
            End Using
        Catch ex As Exception
            'Console.WriteLine("Error al leer el archivo: " & ex.Message)
            rsboApp.StatusBar.SetText(" Error al Leer archivo .cvs, " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return 4 ' Error al leer archivo
        End Try
    End Function


End Class
