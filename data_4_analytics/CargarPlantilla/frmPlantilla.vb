Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Security.Permissions
Imports System.Windows.Forms
Imports SAPbouiCOM
Imports System.Globalization
Imports Microsoft.Office.Interop


Public Class frmPlantilla

    Private oForm As SAPbouiCOM.Form

    Private rCompany As SAPbobsCOM.Company
    Private WithEvents rsboApp As SAPbouiCOM.Application

    Dim lineaMatrix As Integer = 0
    Dim customCulture As CultureInfo

    Dim oCFL As SAPbouiCOM.ChooseFromList
    Dim oCFL_Ter As SAPbouiCOM.ChooseFromList
    Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
    Dim oCFLCreationParamsTer As SAPbouiCOM.ChooseFromListCreationParams
    Dim oConditions As SAPbouiCOM.Conditions
    Dim oCondition As SAPbouiCOM.Condition
    Dim oUserDataSource As SAPbouiCOM.UserDataSource

    Dim columnasEsperadas As Integer = 24

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormularioPlantilla()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String
        If RecorreFormulario(rsboApp, "frmPlantilla") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmPlantilla.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)

            Catch exx As Exception
                rsboApp.Forms.Item("frmPlantilla").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmPlantilla")
            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = Windows.Forms.Application.StartupPath & "\LogoSS.png"

            oForm.Mode = BoFormMode.fm_ADD_MODE

            ' Obtener los ChooseFromLists del formulario
            oCFLs = oForm.ChooseFromLists

            ' Crear parámetros para el CFL
            oCFLCreationParams = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = 62 ' Código para OOCR (centros de coste)
            oCFLCreationParams.UniqueID = "CFL_Dim3"

            ' Crear el CFL
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Configurar condiciones del CFL para filtrar por DimCode = 3
            oConditions = oCFL.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "DimCode" ' Campo a filtrar
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "3" ' Valor del filtro

            oCFL.SetConditions(oConditions)


            'oCFLCreationParamsTer = rsboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            'oCFLCreationParamsTer.MultiSelection = False
            'oCFLCreationParamsTer.ObjectType = "SSTERCEROS" ' UDO prefix obligatorio
            'oCFLCreationParamsTer.UniqueID = "CFL_Ter"

            'oCFL_Ter = oCFLs.Add(oCFLCreationParamsTer)


            Dim txtCodSuc As SAPbouiCOM.EditText
            txtCodSuc = oForm.Items.Item("txtCodSuc").Specific
            'oForm.DataSources.UserDataSources.Add("EditDSP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            txtCodSuc.DataBind.SetBound(True, "@SSVALPLANTCAB", "U_CodSuc")
            txtCodSuc.ChooseFromListUID = "CFL_Dim3"

            txtCodSuc.ChooseFromListAlias = "OcrCode"

            Dim txtCodTer As SAPbouiCOM.EditText
            txtCodTer = oForm.Items.Item("txtCodTer").Specific
            'txtCodTer.DataBind.SetBound(True, "@SSVALPLANTCAB", "U_IdTercero")
            'txtCodTer.ChooseFromListUID = "CFL_Ter"
            'txtCodTer.ChooseFromListAlias = "DocEntry"
            'txtCodTer.ChooseFromListAlias = "U_IdTercero"
            'txtCodTer.ChooseFromListAlias = "U_NomTercero"

            Dim txtNomTer As SAPbouiCOM.EditText
            txtNomTer = oForm.Items.Item("txtNomTer").Specific
            'txtNomTer.DataBind.SetBound(True, "@SSVALPLANTCAB", "U_NomTercero")


            llena_comboAnio()
            llena_comboAdicional()

            Dim lbl As SAPbouiCOM.StaticText = oForm.Items.Item("Item_0").Specific
            lbl.Item.Visible = False

            Dim txtBL As SAPbouiCOM.EditText = oForm.Items.Item("txtBL").Specific
            txtBL.Item.Visible = False

            Dim btnBusL As SAPbouiCOM.Button = oForm.Items.Item("btnBusL").Specific
            btnBusL.Item.Visible = False

            'Dim btn_addTer As SAPbouiCOM.Button = oForm.Items.Item("btn_addTer").Specific
            'btn_addTer.Item.Visible = True

            Dim screenWidth As Integer = rsboApp.Desktop.Width
            Dim screenHeight As Integer = rsboApp.Desktop.Height

            ' Obtener el tamaño del formulario
            Dim formWidth As Integer = oForm.Width
            Dim formHeight As Integer = oForm.Height

            ' Calcular la posición para centrar el formulario
            Dim centeredLeft As Integer = (screenWidth - formWidth) / 2
            Dim centeredTop As Integer = (screenHeight - formHeight) / 2

            ' Ajustar la posición del formulario
            'oForm.Left = centeredLeft
            'oForm.Top = centeredTop - 75

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        End Try

    End Sub

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            'If pVal.BeforeAction = False AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_MENU_CLICK AndAlso pVal.FormTypeEx = "frmPlantilla" Then
            '    If pVal.MenuUID = "1283" Then ' 1283 = Eliminar
            '        Dim respuesta As Integer = rsboApp.MessageBox("¿Está seguro que desea eliminar el registro?", 1, "Sí", "No", "")
            '        If respuesta <> 1 Then
            '            ' Usuario presionó "No"
            '            BubbleEvent = False
            '        End If
            '    End If
            'End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                Utilitario.Util_Log.Escribir_Log("evento click", "frmPlantilla")
                If pVal.FormTypeEx = "frmPlantilla" Then

                    Utilitario.Util_Log.Escribir_Log("evento click frmPlantilla ", "frmPlantilla")
                    If pVal.BeforeAction = False And pVal.ItemUID = "btnMas" Then
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantilla ItemUID = btnMas", "frmPlantilla")
                        rsboApp.Forms.ActiveForm.Freeze(True)

                        Dim oMatrix As SAPbouiCOM.Matrix = CType(oForm.Items.Item("MTX_UDO").Specific, SAPbouiCOM.Matrix)

                        If oMatrix.RowCount = 0 Then

                            oMatrix.AddRow()
                            For i As Integer = 1 To oMatrix.RowCount
                                oMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i
                            Next

                        Else

                            AddRowAtSelectedPosition("frmPlantilla", oMatrix, lineaMatrix)

                        End If
                        rsboApp.Forms.ActiveForm.Freeze(False)


                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btnMenos" Then
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantilla ItemUID = btnMenos", "frmPlantilla")
                        Dim mMatrix As SAPbouiCOM.Matrix = rsboApp.Forms.Item(rsboApp.Forms.ActiveForm.UniqueID).Items.Item("MTX_UDO").Specific

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
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantilla ItemUID = btnCP", "frmPlantilla")
                        Dim selectFileDialog As New SelectFileDialog("C:\", "", "CSV files (*.csv)|*.csv|All files (*.*)|*.*", DialogType.OPEN)
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
                                    rsboApp.StatusBar.SetText(" [Error]: El archivo no está delimitado por ';'. Cod. Error: {ValidarCSV(ruta, columnasEsperadas).ToString}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Case 2
                                    rsboApp.StatusBar.SetText($" [Error]: Se esperaban {columnasEsperadas} columnas, pero se encontraron menos o más columnas. Cod. Error: {validacion}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Case 3
                                    rsboApp.StatusBar.SetText($" [Error]: Archivo vacío: {validacion}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Case Else
                                    rsboApp.StatusBar.SetText($" [Error]: Puede que el archivo que intenta cargar, se encuentre abierto: {validacion}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Select


                        End If

                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "1" Then
                        If oForm.Mode = BoFormMode.fm_FIND_MODE Then




                        Else
                            oForm.EnableMenu("1282", True)
                        End If

                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "MTX_UDO" Then
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantilla ItemUID = MTX_UDO", "frmPlantilla")
                        Dim selectedRow As Integer = pVal.Row
                        lineaMatrix = pVal.Row
                        If selectedRow > 0 Then
                            rsboApp.StatusBar.SetText("Fila seleccionada: " & selectedRow.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If

                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btnBusL" Then
                        Utilitario.Util_Log.Escribir_Log("evento click frmPlantilla ItemUID = btnBusL", "frmPlantilla")
                        Dim BuscarLinea As SAPbouiCOM.EditText = oForm.Items.Item("txtBL").Specific

                        Dim LineaSelect As Integer = CInt(BuscarLinea.Value)
                        Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_UDO").Specific

                        oForm.Freeze(True)
                        mMatrix.FlushToDataSource() ' Deshabilita la actualización en tiempo real.
                        mMatrix.SelectRow(LineaSelect, True, False) ' Selecciona la fila.
                        'mMatrix.LoadFromDataSource() ' Carga nuevamente los datos.
                        oForm.Freeze(False)
                    ElseIf pVal.BeforeAction = False And pVal.ItemUID = "btn_addTer" Then
                        Dim frmTer As New frmCreaTercero(rCompany, rsboApp)
                        frmTer.CargarFormulario()
                    End If

                    'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED AndAlso pVal.ItemUID = "btn_addTer" AndAlso pVal.BeforeAction = False Then
                    '    ' Lógica para abrir frmCreaTercero
                    '    Dim frmTer As New frmCreaTercero(rCompany, rsboApp)
                    '    frmTer.CargarFormulario()
                    'End If

                End If
            End If

            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                Utilitario.Util_Log.Escribir_Log("evento et_ITEM_PRESSED", "frmPlantilla")
                If pVal.FormTypeEx = "frmPlantilla" Then

                    If pVal.BeforeAction = False And pVal.ItemUID = "chk1FC" Then
                        Dim chk1FC As SAPbouiCOM.CheckBox = oForm.Items.Item("chk1FC").Specific

                        Dim oMatrix As SAPbouiCOM.Matrix
                        oMatrix = oForm.Items.Item("MTX_UDO").Specific

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
                    oFor = rsboApp.Forms.Item("frmPlantilla")
                    If pVal.ItemUID = "btnCP" And oFor.Mode <> BoFormMode.fm_ADD_MODE Then 'pVal.BeforeAction = False And
                        'Dim btnCP As SAPbouiCOM.Button = oForm.Items.Item("btnCP").Specific
                        oFor.Mode = BoFormMode.fm_UPDATE_MODE
                    End If

                End If

            End If

            If pVal.FormTypeEx = "frmPlantilla" Then

                If pVal.BeforeAction And pVal.ItemUID = "1" Then

                    Try
                        Dim oFor As SAPbouiCOM.Form
                        oFor = rsboApp.Forms.Item("frmPlantilla")
                        Dim cmbAnio As SAPbouiCOM.ComboBox = oFor.Items.Item("cbmAnio").Specific


                        If IsNothing(cmbAnio.Selected) Then
                            If oFor.Mode = BoFormMode.fm_ADD_MODE Then
                                BubbleEvent = False
                                rsboApp.StatusBar.SetText("Por favor seleccionar un Año", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If


                        Else
                            If oFor.Mode = BoFormMode.fm_ADD_MODE Then


                                Dim año As Integer = CInt(cmbAnio.Selected.Value)
                                Dim AñoActual As Integer = Year(Date.Now)
                                Dim cbmAnio As SAPbouiCOM.ComboBox = oForm.Items.Item("cbmAnio").Specific
                                AñoActual = CInt(cbmAnio.Selected.Value)
                                Dim txtCodSuc As SAPbouiCOM.EditText = oForm.Items.Item("txtCodSuc").Specific
                                Dim txtCodTer As SAPbouiCOM.EditText = oForm.Items.Item("txtCodTer").Specific
                                Dim CodSucursal As String = IIf(String.IsNullOrEmpty(txtCodSuc.Value), "", txtCodSuc.Value)
                                Dim CodTer As String = IIf(String.IsNullOrEmpty(txtCodTer.Value), "", txtCodTer.Value)
                                Dim ConRegistro As String = ""


                                If CodSucursal = "" Then
                                    BubbleEvent = False
                                    rsboApp.StatusBar.SetText("Por favor colocar la sucursal..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Else
                                    ' TEMPORALMENTE DESACTIVADO

                                    'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    '    ConRegistro = "select COUNT(1) as ""Contador"" from """ + rCompany.CompanyDB + """.""@SSVALPLANTCAB"" where ""U_Anio""='" + AñoActual.ToString + "'" + " and ""U_CodSuc""='" + CodSucursal.ToString + "'"
                                    'Else
                                    '    ConRegistro = "select COUNT(1) as ""Contador"" from ""@SSVALPLANTCAB"" where ""U_Anio""='" + AñoActual.ToString + "'" + " and ""U_CodSuc""='" + CodSucursal.ToString + "'"
                                    'End If

                                    'Utilitario.Util_Log.Escribir_Log("Consulta cantidad registro por año: " + ConRegistro.ToString, "frmPlantilla")

                                    'Dim Val As Integer = CInt(oFuncionesB1.getRSvalue(ConRegistro, "Contador", "0"))
                                    ''If año <> AñoActual Then
                                    'If Val > 0 Then
                                    '    BubbleEvent = False
                                    '    rsboApp.StatusBar.SetText("Solo se permite un registro por Año y Sucursal..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    'End If

                                End If

                            End If

                            If oFor.Mode = BoFormMode.fm_UPDATE_MODE Then


                                Dim año As Integer = CInt(cmbAnio.Selected.Value)
                                Dim AñoActual As Integer = Year(Date.Now)
                                Dim cbmAnio As SAPbouiCOM.ComboBox = oForm.Items.Item("cbmAnio").Specific
                                AñoActual = CInt(cbmAnio.Selected.Value)
                                Dim txtCodSuc As SAPbouiCOM.EditText = oForm.Items.Item("txtCodSuc").Specific
                                Dim CodSucursal As String = IIf(String.IsNullOrEmpty(txtCodSuc.Value), "", txtCodSuc.Value)
                                Dim txtCodTer As SAPbouiCOM.EditText = oForm.Items.Item("txtCodTer").Specific
                                Dim CodTer As String = IIf(String.IsNullOrEmpty(txtCodTer.Value), "", txtCodTer.Value)
                                Dim ConRegistro As String = ""


                                If CodSucursal = "" Then
                                    BubbleEvent = False
                                    rsboApp.StatusBar.SetText("Por favor colocar la sucursal..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Else
                                    ' TEMPORALMENTE DESACTIVADO

                                    'If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                    '    ConRegistro = "select COUNT(1) as ""Contador"" from """ + rCompany.CompanyDB + """.""@SSVALPLANTCAB"" where ""U_Anio""='" + AñoActual.ToString + "'" + " and ""U_CodSuc""='" + CodSucursal.ToString + "'"
                                    'Else
                                    '    ConRegistro = "select COUNT(1) as ""Contador"" from ""@SSVALPLANTCAB"" where ""U_Anio""='" + AñoActual.ToString + "'" + " and ""U_CodSuc""='" + CodSucursal.ToString + "'"
                                    'End If

                                    'Utilitario.Util_Log.Escribir_Log("Consulta cantidad registro por año: " + ConRegistro.ToString, "frmPlantilla")

                                    'Dim Val As Integer = CInt(oFuncionesB1.getRSvalue(ConRegistro, "Contador", "0"))
                                    ''If año <> AñoActual Then
                                    'If Val > 1 Then
                                    '    BubbleEvent = False
                                    '    rsboApp.StatusBar.SetText("Solo se permite un registro por Año y Sucursal..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    'End If

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
                Utilitario.Util_Log.Escribir_Log("evento et_MATRIX_LOAD", "frmPlantilla")
                If pVal.FormTypeEx = "frmPlantilla" Then
                    If pVal.BeforeAction Then
                        Dim EditText As SAPbouiCOM.EditText
                        'EditText = oForm.Items.Item("Item_4").Specific
                        'EditText.Item.Click()

                        EditText = oForm.Items.Item("Item_10").Specific
                        EditText.Item.Enabled = False
                    End If
                End If





            End If

            If pVal.EventType = BoEventTypes.et_KEY_DOWN Then
                Utilitario.Util_Log.Escribir_Log("evento et_KEY_DOWN", "frmPlantilla")
                If pVal.FormTypeEx = "frmPlantilla" Then

                    If pVal.BeforeAction = False Then
                        If pVal.CharPressed = Keys.Enter Then

                            Dim BuscarLinea As SAPbouiCOM.EditText = oForm.Items.Item("txtBL").Specific

                            Dim LineaSelect As Integer = CInt(BuscarLinea.Value)
                            Dim mMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_UDO").Specific

                            mMatrix.SelectRow(LineaSelect, True, False)

                            'MessageBox.Show("Enter key pressed")
                        End If
                    End If

                End If

            End If


            If pVal.EventType = BoEventTypes.et_CHOOSE_FROM_LIST Then
                Utilitario.Util_Log.Escribir_Log("evento et_CHOOSE_FROM_LIST", "frmPlantilla")
                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                oCFLEvento = pVal
                If oCFLEvento.BeforeAction = False Then
                    If pVal.FormTypeEx = "frmPlantilla" Then

                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item("frmPlantilla")

                        If sCFL_ID = "CFL_Dim3" Then
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            Dim val As String = String.Empty
                            Dim val1 As String = String.Empty

                            If Not oDataTable Is Nothing Then
                                val = oDataTable.GetValue(0, 0)   ' OcrCode
                                val1 = oDataTable.GetValue(1, 0)  ' OcrName
                                Try
                                    oForm.DataSources.DBDataSources.Item("@SSVALPLANTCAB").SetValue("U_CodSuc", 0, val)
                                    oForm.DataSources.DBDataSources.Item("@SSVALPLANTCAB").SetValue("U_NomSuc", 0, val1)

                                    ' Actualizar el EditText de nombre de sucursal
                                    Dim txtNomSuc As SAPbouiCOM.EditText
                                    txtNomSuc = oForm.Items.Item("txtNomSuc").Specific
                                    txtNomSuc.Value = val1

                                    ' *** AGREGAR ESTA PARTE PARA CAMBIAR EL FORMULARIO A MODO UPDATE ***
                                    If oForm.Mode = BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = BoFormMode.fm_UPDATE_MODE
                                    End If

                                Catch ex As Exception
                                    Utilitario.Util_Log.Escribir_Log("Error al actualizar sucursal: " + ex.Message, "frmPlantilla")
                                End Try
                            Else
                                ' Si no hay selección, limpiar el campo de nombre
                                Dim txtNomSuc As SAPbouiCOM.EditText
                                txtNomSuc = oForm.Items.Item("txtNomSuc").Specific
                                txtNomSuc.Value = String.Empty
                                oForm.DataSources.DBDataSources.Item("@SSVALPLANTCAB").SetValue("U_NomSuc", 0, String.Empty)

                                ' *** CAMBIAR A MODO UPDATE TAMBIÉN AL LIMPIAR ***
                                If oForm.Mode = BoFormMode.fm_OK_MODE Then
                                    oForm.Mode = BoFormMode.fm_UPDATE_MODE
                                End If
                            End If

                            'ElseIf sCFL_ID = "CFL_Ter" Then
                            '    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                            '    Dim oDataTable As SAPbouiCOM.DataTable
                            '    oDataTable = oCFLEvento.SelectedObjects
                            '    Dim docEntry As String = String.Empty
                            '    Dim nombreTercero As String = String.Empty

                            '    If Not oDataTable Is Nothing Then
                            '        docEntry = oDataTable.GetValue(0, 0)  ' DocEntry

                            '        Try
                            '            ' Actualizar el DocEntry en el datasource
                            '            oForm.DataSources.DBDataSources.Item("@SSVALPLANTCAB").SetValue("U_IdTercero", 0, docEntry)

                            '            ' Obtener el nombre del tercero consultando la tabla detalle
                            '            Dim consultaNombre As String = ""
                            '            If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            '                consultaNombre = "SELECT TOP 1 ""U_NomTercero"" FROM """ + rCompany.CompanyDB + """.""@SSTERCEROSDET"" WHERE ""DocEntry"" = '" + docEntry + "'"
                            '            Else
                            '                consultaNombre = "SELECT TOP 1 ""U_NomTercero"" FROM ""@SSTERCEROSDET"" WHERE ""DocEntry"" = '" + docEntry + "'"
                            '            End If

                            '            nombreTercero = oFuncionesB1.getRSvalue(consultaNombre, "U_NomTercero", "")

                            '            ' Actualizar el datasource y el EditText
                            '            oForm.DataSources.DBDataSources.Item("@SSVALPLANTCAB").SetValue("U_NomTercero", 0, nombreTercero)

                            '            Dim txtNomTer As SAPbouiCOM.EditText
                            '            txtNomTer = oForm.Items.Item("txtNomTer").Specific
                            '            txtNomTer.Value = nombreTercero

                            '            ' *** AGREGAR ESTA PARTE PARA CAMBIAR EL FORMULARIO A MODO UPDATE ***
                            '            If oForm.Mode = BoFormMode.fm_OK_MODE Then
                            '                oForm.Mode = BoFormMode.fm_UPDATE_MODE
                            '            End If

                            '        Catch ex As Exception
                            '            Utilitario.Util_Log.Escribir_Log("Error al actualizar tercero: " + ex.Message, "frmPlantilla")
                            '        End Try
                            '    Else
                            '        ' Si no hay selección, limpiar los campos
                            '        Dim txtNomTer As SAPbouiCOM.EditText
                            '        txtNomTer = oForm.Items.Item("txtNomTer").Specific
                            '        txtNomTer.Value = String.Empty
                            '        oForm.DataSources.DBDataSources.Item("@SSVALPLANTCAB").SetValue("U_NomTercero", 0, String.Empty)

                            '        ' *** CAMBIAR A MODO UPDATE TAMBIÉN AL LIMPIAR ***
                            '        If oForm.Mode = BoFormMode.fm_OK_MODE Then
                            '            oForm.Mode = BoFormMode.fm_UPDATE_MODE
                            '        End If
                            '    End If
                        End If
                    End If
                End If
            End If

            If pVal.FormTypeEx = "frmPlantilla" AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN AndAlso pVal.BeforeAction = False Then
                If pVal.CharPressed = 9 Then
                    Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item(pVal.FormUID)
                    Dim txtCodTer As SAPbouiCOM.EditText = oForm.Items.Item("txtCodTer").Specific
                    Dim txtNomTer As SAPbouiCOM.EditText = oForm.Items.Item("txtNomTer").Specific

                    Dim codTer As String = txtCodTer.Value.Trim
                    If codTer <> "" Then
                        Dim query As String
                        If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            query = $"SELECT TOP 1 ""U_NomTercero"" FROM ""@SSTERCEROS"" WHERE ""Code"" = '{codTer}'"
                        Else
                            query = $"SELECT TOP 1 ""U_NomTercero"" FROM ""@SSTERCEROS"" WHERE ""Code"" = '{codTer}'"
                        End If

                        Dim oRs As SAPbobsCOM.Recordset = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRs.DoQuery(query)

                        If oRs.RecordCount > 0 Then
                            Dim nomTer As String = oRs.Fields.Item("U_NomTercero").Value.ToString()
                            txtNomTer.Value = nomTer
                            oForm.DataSources.DBDataSources.Item("@SSVALPLANTCAB").SetValue("U_NomTercero", 0, nomTer)
                        Else
                            txtNomTer.Value = ""
                            oForm.DataSources.DBDataSources.Item("@SSVALPLANTCAB").SetValue("U_NomTercero", 0, "")
                            rsboApp.StatusBar.SetText("No se encontró el tercero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If

                        ' Asegurar que el formulario esté en modo Update si estaba en OK
                        If oForm.Mode = BoFormMode.fm_OK_MODE Then
                            oForm.Mode = BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                End If
            End If




        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("ex rSboApp_ItemEvent: " + ex.Message.ToString(), "frmPlantilla")
            'System.Windows.Forms.MessageBox.Show("Error rSboApp_ItemEvent :" & ex.Message.ToString())
        End Try


    End Sub

    Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent
        Try
            If pVal.MenuUID = "1281" Then 'Buscar
                If pVal.BeforeAction = False Then
                    If rsboApp.Forms.ActiveForm.UniqueID = "frmPlantilla" Then
                        Try
                            Dim NumeroUdo As SAPbouiCOM.EditText = oForm.Items.Item("Item_10").Specific
                            NumeroUdo.Item.Enabled = True

                        Catch ex As Exception

                        End Try

                    End If

                End If
            End If
        Catch ex As Exception
            Utilitario.Util_Log.Escribir_Log("Error en MenuEvent: " + ex.Message.ToString, "frmPlantilla")
        End Try

    End Sub

    Private Sub rsboApp_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles rsboApp.RightClickEvent

        Try
            If eventInfo.FormUID = "frmPlantilla" Then

                If eventInfo.ItemUID = "MTX_UDO" Then

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

    Private Sub llena_comboAnio()

        Dim queryAnio As String = "select * from ""@SS_ANIO"""
        'Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing

        oForm = rsboApp.Forms.Item("frmPlantilla")
        Dim cboAnio As SAPbouiCOM.ComboBox
        cboAnio = oForm.Items.Item("cbmAnio").Specific

        Dim oRecordSet As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet("select * from ""@SS_ANIO"" order by ""Code""")
        ValoresValidos = cboAnio.ValidValues
        While cboAnio.ValidValues.Count > 0
            cboAnio.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End While
        If oRecordSet.RecordCount > 1 Then
            While (oRecordSet.EoF = False)
                'rsboApp.SetStatusBarMessage("Valor" + oRecordSet.Fields.Item("Code").Value.ToString + " descripcion: " + oRecordSet.Fields.Item("Name").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                ValoresValidos.Add(Convert.ToString(oRecordSet.Fields.Item("Code").Value), Convert.ToString(oRecordSet.Fields.Item("Name").Value))
                oRecordSet.MoveNext()
            End While
        End If
    End Sub

    Private Sub llena_comboAdicional()

        Dim queryAdi As String = "select * from ""@SS_PPTO_ADIC"""
        'Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim ValoresValidos As SAPbouiCOM.ValidValues = Nothing

        oForm = rsboApp.Forms.Item("frmPlantilla")
        Dim cboAdi As SAPbouiCOM.ComboBox
        cboAdi = oForm.Items.Item("cbmAdi").Specific

        Dim oRecordSet As SAPbobsCOM.Recordset = oFuncionesB1.getRecordSet("select * from ""@SS_PPTO_ADIC"" order by ""Code""")
        ValoresValidos = cboAdi.ValidValues
        While cboAdi.ValidValues.Count > 0
            cboAdi.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
        End While
        If oRecordSet.RecordCount > 1 Then
            While (oRecordSet.EoF = False)
                'rsboApp.SetStatusBarMessage("Valor" + oRecordSet.Fields.Item("Code").Value.ToString + " descripcion: " + oRecordSet.Fields.Item("Name").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                ValoresValidos.Add(Convert.ToString(oRecordSet.Fields.Item("Code").Value), Convert.ToString(oRecordSet.Fields.Item("Name").Value))
                oRecordSet.MoveNext()
            End While
        End If
    End Sub


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

            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SSVALPLANTDET")
            oMatrix.FlushToDataSource()

            Dim LineaSSetear As Integer = SelectedRow


            'oDataSource.InsertRecord(SelectedRow - 1)
            oDataSource.SetValue("LineId", LineaSSetear, SelectedRow + 1)
            oDataSource.SetValue("U_Tipo", LineaSSetear, "P")
            'oDataSource.SetValue("U_Nivel1", LineaSSetear, String.Empty)
            'oDataSource.SetValue("U_Nivel2", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_Nivel3", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_NomCuenta", LineaSSetear, String.Empty)

            'AÑADIDO 2025/06/04
            oDataSource.SetValue("U_CtaAsoc", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_NomCtaAsc", LineaSSetear, String.Empty)
            'AÑADIDO 2025/06/04
            oDataSource.SetValue("U_CodAgrup", LineaSSetear, String.Empty)

            oDataSource.SetValue("U_Proyecto", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_NomProy", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_LineaNeg", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_SubLinea", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_Sucursal", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_Espacio", LineaSSetear, String.Empty)
            'oDataSource.SetValue("U_Espacio2", LineaSSetear, String.Empty)

            oDataSource.SetValue("U_Enero", LineaSSetear, "0")
            oDataSource.SetValue("U_Febrero", LineaSSetear, "0")
            oDataSource.SetValue("U_Marzo", LineaSSetear, "0")
            oDataSource.SetValue("U_Abril", LineaSSetear, "0")
            oDataSource.SetValue("U_Mayo", LineaSSetear, "0")
            oDataSource.SetValue("U_Junio", LineaSSetear, "0")
            oDataSource.SetValue("U_Julio", LineaSSetear, "0")
            oDataSource.SetValue("U_Agosto", LineaSSetear, "0")
            oDataSource.SetValue("U_Septiembre", LineaSSetear, "0")
            oDataSource.SetValue("U_Octubre", LineaSSetear, "0")
            oDataSource.SetValue("U_Noviembre", LineaSSetear, "0")
            oDataSource.SetValue("U_Diciembre", LineaSSetear, "0")

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
            Utilitario.Util_Log.Escribir_Log("Cultura configurada: Separador decimal= " & customCulture.NumberFormat.NumberDecimalSeparator & " Separador de Miles= " & customCulture.NumberFormat.NumberGroupSeparator, "frmPlantilla")


            ' Suponiendo que tienes una referencia a la Matrix
            Dim oMatrix As Matrix = oForm.Items.Item("MTX_UDO").Specific
            'Dim source As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_DETUDO")
            Dim fil As Integer = 0
            ' Ruta al archivo CSV
            Dim csvPath As String = ruta
            Dim DatosEnExcel As New List(Of String)
            Using reader As New StreamReader(csvPath)
                Dim rowIndex As Integer = 1

                If rowIndex = 1 Then
                    reader.ReadLine()
                End If

                While Not reader.EndOfStream
                    ' Leer cada línea del archivo CSV
                    Dim line As String = reader.ReadLine()
                    DatosEnExcel.Add(line)
                    rowIndex += 1
                End While
            End Using

            Dim oProgressBar As SAPbouiCOM.ProgressBar
            oProgressBar = rsboApp.StatusBar.CreateProgressBar("Procesando CSV..", DatosEnExcel.Count, False)

            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SSVALPLANTDET")

            oDataSource.Clear()
            oMatrix.Clear()

            Dim j As Integer = 0

            Dim TotalLineas = DatosEnExcel.Count
            Dim ValorUpdate = TotalLineas / 100
            For Each filas In DatosEnExcel
                oDataSource.InsertRecord(j)
                oDataSource.SetValue("LineId", j, j + 1)
                oDataSource.SetValue("U_Tipo", j, filas.Split(";")(0))
                'oDataSource.SetValue("U_Nivel1", j, filas.Split(";")(0))
                'oDataSource.SetValue("U_Nivel2", j, filas.Split(";")(1))
                oDataSource.SetValue("U_Nivel3", j, filas.Split(";")(1))
                oDataSource.SetValue("U_NomCuenta", j, filas.Split(";")(2))

                oDataSource.SetValue("U_CtaAsoc", j, filas.Split(";")(3))
                oDataSource.SetValue("U_NomCtaAsc", j, filas.Split(";")(4))

                oDataSource.SetValue("U_CodAgrup", j, filas.Split(";")(5))

                oDataSource.SetValue("U_Proyecto", j, filas.Split(";")(6))
                oDataSource.SetValue("U_NomProy", j, filas.Split(";")(7))
                oDataSource.SetValue("U_LineaNeg", j, filas.Split(";")(8))
                oDataSource.SetValue("U_SubLinea", j, filas.Split(";")(9))
                oDataSource.SetValue("U_Sucursal", j, filas.Split(";")(10))
                oDataSource.SetValue("U_Espacio", j, filas.Split(";")(11))
                'oDataSource.SetValue("U_Espacio2", j, filas.Split(";")(10))

                oDataSource.SetValue("U_Enero", j, filas.Split(";")(12))
                oDataSource.SetValue("U_Febrero", j, filas.Split(";")(13))
                oDataSource.SetValue("U_Marzo", j, filas.Split(";")(14))
                oDataSource.SetValue("U_Abril", j, filas.Split(";")(15))
                oDataSource.SetValue("U_Mayo", j, filas.Split(";")(16))
                oDataSource.SetValue("U_Junio", j, filas.Split(";")(17))
                oDataSource.SetValue("U_Julio", j, filas.Split(";")(18))
                oDataSource.SetValue("U_Agosto", j, filas.Split(";")(19))
                oDataSource.SetValue("U_Septiembre", j, filas.Split(";")(20))
                oDataSource.SetValue("U_Octubre", j, filas.Split(";")(21))
                oDataSource.SetValue("U_Noviembre", j, filas.Split(";")(22))
                oDataSource.SetValue("U_Diciembre", j, filas.Split(";")(23))


                oProgressBar.Value = j


                j += 1
            Next
            oMatrix.LoadFromDataSource()

            oForm.Freeze(False)
            oProgressBar.Stop()
            rsboApp.StatusBar.SetText("Datos cargados desde el archivo Excel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Return True

        Catch ex As Exception
            ' mensaje = "Error al Leer archivo .cvs : " + ex.Message.ToString()
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
