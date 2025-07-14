Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Security.Permissions
Imports System.Windows.Forms
Imports SAPbouiCOM
Imports System.Globalization
Imports Microsoft.Office.Interop


Public Class frmEstadistica

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
    Dim oMatrix As SAPbouiCOM.Matrix
    Dim txtDocEntry As SAPbouiCOM.EditText
    Dim txtAnio As SAPbouiCOM.EditText

    Dim columnasEsperadas As Integer = 6

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormEstadistico()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmEstadistica") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmEstadistica.srf"
        xmlDoc.Load(strPath)
        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)

            Catch exx As Exception
                rsboApp.Forms.Item("frmEstadistica").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmEstadistica")
            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = Windows.Forms.Application.StartupPath & "\LogoSS.png"

            oForm.Height = 530
            oForm.Width = 600

            oForm = rsboApp.Forms.Item("frmEstadistica")
            oForm.Freeze(True)

            txtDocEntry = oForm.Items.Item("txtDEntry").Specific

            txtAnio = oForm.Items.Item("txtAnio").Specific
            txtAnio.Value = DateTime.Now.Year.ToString()

            oMatrix = oForm.Items.Item("MTX_UDO").Specific

            Try
                Dim colCC As SAPbouiCOM.Column = oMatrix.Columns.Item("B2CComer")
                colCC.DisplayDesc = True
            Catch ex As Exception
                'Si no existe la columna simplemente continuamos.
            End Try

            txtDocEntry.Item.Enabled = False

            Dim anchoTotal As Integer = oForm.Items.Item("MTX_UDO").Width

            'oMatrix.Columns.Item(0).Width = 20

            'anchoTotal = anchoTotal - 20

            Dim anchoPorColumna As Integer = anchoTotal \ oMatrix.Columns.Count

            For i As Integer = 0 To oMatrix.Columns.Count - 1
                If i = 0 Then
                    oMatrix.Columns.Item(i).Width = 30
                Else
                    oMatrix.Columns.Item(i).Width = 85
                End If
            Next

            oForm.Mode = BoFormMode.fm_ADD_MODE

            oForm.Visible = True
            oForm.Select()

            oForm.Freeze(False)

        Catch ex As Exception
            rsboApp.MessageBox(NombreAddon + " Ocurrio un Error al Cargar la Pantalla: " + ex.Message.ToString())
        Finally
            txtAnio.Value = DateTime.Now.Year.ToString()
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

    Private Sub rsboApp_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.ItemEvent
        Try
            If pVal.FormUID = "frmEstadistica" Then
                Select Case pVal.ItemUID

                    'Case "btn_cons"
                    '    If pVal.BeforeAction = False AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                    '    End If

                    Case "btnCP"
                        If pVal.BeforeAction = True AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim Anio As String = If(String.IsNullOrEmpty(txtAnio.Value), "", txtAnio.Value)
                            Try
                                Dim AnioInt As Integer = CInt(Anio)

                                If Anio = "" Then
                                    rsboApp.StatusBar.SetText("Ingrese el año...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                ' Diálogo de archivo
                                Utilitario.Util_Log.Escribir_Log("evento click frmEstadistica ItemUID = btnCP", "frmEstadistica")
                                Dim selectFileDialog As New SelectFileDialog("C:\", "", "CSV files (*.csv)|*.csv|All files (*.*)|*.*", DialogType.OPEN)
                                selectFileDialog.Open()

                                If Not String.IsNullOrWhiteSpace(selectFileDialog.SelectedFile) Then
                                    Dim ruta As String = selectFileDialog.SelectedFile
                                    Dim validacion As Integer = ValidarCSV(ruta, columnasEsperadas)

                                    Select Case validacion
                                        Case 0
                                            If ValidarContenidoCSV(ruta) Then
                                                rsboApp.StatusBar.SetText("✅ El archivo es válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                leerCSV(ruta)
                                            End If
                                        Case 1
                                            rsboApp.StatusBar.SetText("[Error]: El archivo no está delimitado por ';'.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Case 2
                                            rsboApp.StatusBar.SetText($"[Error]: Se esperaban {columnasEsperadas} columnas, pero se encontraron menos o más.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Case 3
                                            rsboApp.StatusBar.SetText("[Error]: Archivo vacío.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Case Else
                                            rsboApp.StatusBar.SetText("[Error]: Puede que el archivo esté abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End Select
                                End If

                            Catch ex As Exception
                                rsboApp.StatusBar.SetText("Ingrese un valor válido para año...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                        End If

                    Case "1" ' Botón grabar
                        If pVal.BeforeAction = True AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim AnioIngresado As String = txtAnio.Value.Trim()

                            If oMatrix.RowCount = 0 OrElse oMatrix.VisualRowCount = 0 Then
                                rsboApp.MessageBox("No se puede grabar. La Matrix está vacía.")
                                BubbleEvent = False
                                Exit Sub
                            End If

                            If String.IsNullOrEmpty(AnioIngresado) OrElse Not IsNumeric(AnioIngresado) Then
                                rsboApp.StatusBar.SetText("[Error] Ingrese un año válido antes de grabar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If

                            ' Validar fechas de la matrix
                            'Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_ESTADISTICAD")

                            'For index As Integer = 0 To oDataSource.Size - 1
                            '    Dim fechaStr As String = oDataSource.GetValue("U_FECHA", index).Trim()
                            '    If fechaStr.Length >= 4 Then
                            '        Dim anioFecha As String = fechaStr.Substring(0, 4)
                            '        If anioFecha <> AnioIngresado Then
                            '            rsboApp.StatusBar.SetText($"[Error] El año ingresado ({AnioIngresado}) no coincide con el año de la línea {index + 1}: {anioFecha}.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '            BubbleEvent = False
                            '            Exit Sub
                            '        End If
                            '    End If
                            'Next

                            '' Si pasa todo: OK
                            'rsboApp.StatusBar.SetText("✅ Validación de año correcta.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                Dim resp As Integer = rsboApp.MessageBox("¿Está seguro que desea actualizar el registro?", 1, "Sí", "No")
                                If resp <> 1 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If

                        End If

                    Case "btn_add"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
                            rsboApp.Forms.ActiveForm.Freeze(True)
                            Dim oMatrix As SAPbouiCOM.Matrix = CType(oForm.Items.Item("MTX_UDO").Specific, SAPbouiCOM.Matrix)

                            If oMatrix.RowCount = 0 Then
                                oMatrix.AddRow()
                                For i As Integer = 1 To oMatrix.RowCount
                                    oMatrix.Columns.Item("COL").Cells.Item(i).Specific.String = i
                                Next
                            Else
                                AddRowAtSelectedPosition("frmEstadistica", oMatrix, lineaMatrix)
                            End If
                            rsboApp.Forms.ActiveForm.Freeze(False)
                        End If

                    Case "btn_del"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
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
                        End If

                    Case "MTX_UDO"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
                            Dim selectedRow As Integer = pVal.Row
                            lineaMatrix = pVal.Row
                            If selectedRow > 0 Then
                                rsboApp.StatusBar.SetText("Fila seleccionada: " & selectedRow.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If
                        End If


                End Select
            End If

        Catch ex As Exception
            rsboApp.MessageBox("Error en evento: " & ex.Message)
        End Try
    End Sub

    Private Sub AddRowAtSelectedPosition(ByVal FormUID As String, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal SelectedRow As Integer)

        Try

            Dim oForm As SAPbouiCOM.Form = rsboApp.Forms.Item(FormUID)

            If SelectedRow <= 0 Or SelectedRow > oMatrix.RowCount Then
                rsboApp.StatusBar.SetText("Seleccione una fila válida.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return
            End If

            oMatrix.AddRow(1, SelectedRow)

            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_ESTADISTICAD")
            oMatrix.FlushToDataSource()

            Dim LineaSSetear As Integer = SelectedRow

            oDataSource.SetValue("LineId", LineaSSetear, SelectedRow + 1)
            oDataSource.SetValue("U_CC", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_FECHA", LineaSSetear, String.Empty)
            oDataSource.SetValue("U_VEHICULOS", LineaSSetear, "0")
            oDataSource.SetValue("U_BANDEJAS", LineaSSetear, "0")
            oDataSource.SetValue("U_CINES", LineaSSetear, "0")
            oDataSource.SetValue("U_CLIENTES", LineaSSetear, "0")

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

            Dim oMatrix As Matrix = oForm.Items.Item("MTX_UDO").Specific
            Dim DatosEnExcel As New List(Of String)

            Using reader As New StreamReader(ruta)
                Dim rowIndex As Integer = 1

                If rowIndex = 1 Then
                    reader.ReadLine()
                End If

                While Not reader.EndOfStream
                    Dim line As String = reader.ReadLine()
                    DatosEnExcel.Add(line)
                    rowIndex += 1
                End While
            End Using

            Dim oProgressBar As SAPbouiCOM.ProgressBar
            oProgressBar = rsboApp.StatusBar.CreateProgressBar("Procesando CSV..", DatosEnExcel.Count, False)

            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_ESTADISTICAD")

            oDataSource.Clear()
            oMatrix.Clear()

            Dim j As Integer = 0
            For Each filas In DatosEnExcel
                oDataSource.InsertRecord(j)
                oDataSource.SetValue("LineId", j, j + 1)
                oDataSource.SetValue("U_CC", j, filas.Split(";")(0))
                oDataSource.SetValue("U_FECHA", j, filas.Split(";")(1))
                oDataSource.SetValue("U_VEHICULOS", j, filas.Split(";")(2))
                oDataSource.SetValue("U_BANDEJAS", j, filas.Split(";")(3))
                oDataSource.SetValue("U_CINES", j, filas.Split(";")(4))
                oDataSource.SetValue("U_CLIENTES", j, filas.Split(";")(5))

                oProgressBar.Value = j
                j += 1
            Next

            Dim anioIngresado As String = txtAnio.Value.Trim()
            Dim anioFecha As String = ""

            For index As Integer = 0 To oDataSource.Size - 1
                Dim fechaStr As String = oDataSource.GetValue("U_FECHA", index).Trim()

                If fechaStr.Length >= 4 Then
                    anioFecha = fechaStr.Substring(0, 4)

                    If anioFecha <> anioIngresado Then
                        rsboApp.StatusBar.SetText($"[Error] La fecha en la línea {index + 1} no coincide con el año ingresado ({anioIngresado}). Valor encontrado: {fechaStr}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.Freeze(False)
                        oProgressBar.Stop()
                        Return False
                    End If
                Else
                    rsboApp.StatusBar.SetText($"[Error] Formato de fecha inválido en línea {index + 1}: {fechaStr}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.Freeze(False)
                    oProgressBar.Stop()
                    Return False
                End If
            Next

            oMatrix.LoadFromDataSource()
            Try
                Dim colCC As SAPbouiCOM.Column = oMatrix.Columns.Item("B2CComer")
                colCC.DisplayDesc = True
            Catch ex As Exception
                'Si no existe la columna simplemente continuamos.
            End Try
            oForm.Freeze(False)
            oProgressBar.Stop()
            rsboApp.StatusBar.SetText("Datos cargados desde el archivo CSV.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return True

        Catch ex As Exception
            rsboApp.StatusBar.SetText(" Error al Leer archivo .csv: " + ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
            Return False
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

    Private Function ValidarContenidoCSV(ruta As String) As Boolean
        Try
            customCulture = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
            customCulture.NumberFormat.NumberDecimalSeparator = "."
            customCulture.NumberFormat.NumberGroupSeparator = ","

            Using sr As New StreamReader(ruta)
                Dim lineNumber As Integer = 1

                If Not sr.EndOfStream Then sr.ReadLine() 'encabezado

                While Not sr.EndOfStream
                    lineNumber += 1
                    Dim line As String = sr.ReadLine()
                    Dim campos() As String = line.Split(";"c)

                    If campos.Length <> columnasEsperadas Then
                        rsboApp.StatusBar.SetText($"[Error] Línea {lineNumber}: cantidad de columnas incorrecta.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If

                    Dim valorEntero As Long
                    If Not Long.TryParse(campos(0).Trim(), NumberStyles.Integer, customCulture, valorEntero) Then
                        rsboApp.StatusBar.SetText($"[Error] Línea {lineNumber}: 'Centro comercial' no es numérico válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If

                    Dim fechaTmp As Date
                    Dim formatos() As String = {"yyyyMMdd", "yyyy-MM-dd"}
                    If Not Date.TryParseExact(campos(1).Trim(), formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, fechaTmp) Then
                        rsboApp.StatusBar.SetText($"[Error] Línea {lineNumber}: 'FECHA' inválida. Use 'yyyyMMdd' o 'yyyy-MM-dd'.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If

                    Dim decimalTmp As Decimal
                    If Not Decimal.TryParse(campos(2).Trim(), NumberStyles.AllowDecimalPoint Or NumberStyles.AllowThousands, customCulture, decimalTmp) Then
                        rsboApp.StatusBar.SetText($"[Error] Línea {lineNumber}: 'Vehiculos' no es número válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If

                    If Not Decimal.TryParse(campos(3).Trim(), NumberStyles.AllowDecimalPoint Or NumberStyles.AllowThousands, customCulture, decimalTmp) Then
                        rsboApp.StatusBar.SetText($"[Error] Línea {lineNumber}: 'Bandejas' no es número válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If

                    If Not Decimal.TryParse(campos(4).Trim(), NumberStyles.AllowDecimalPoint Or NumberStyles.AllowThousands, customCulture, decimalTmp) Then
                        rsboApp.StatusBar.SetText($"[Error] Línea {lineNumber}: 'Cines' no es número válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If

                    If Not Decimal.TryParse(campos(5).Trim(), NumberStyles.AllowDecimalPoint Or NumberStyles.AllowThousands, customCulture, decimalTmp) Then
                        rsboApp.StatusBar.SetText($"[Error] Línea {lineNumber}: 'Clientes' no es número válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If

                End While
            End Using
            Return True
        Catch ex As Exception
            rsboApp.StatusBar.SetText(" Error al validar archivo .csv: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub rsboApp_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles rsboApp.MenuEvent
        Try
            If pVal.BeforeAction AndAlso pVal.MenuUID = "1283" Then
                If rsboApp.Forms.ActiveForm.UniqueID = "frmEstadistica" Then
                    Dim respuesta As Integer = rsboApp.MessageBox("¿Está seguro que desea eliminar el registro?", 1, "Sí", "No")
                    If respuesta <> 1 Then
                        BubbleEvent = False
                    End If
                End If
            End If
        Catch ex As Exception
            rsboApp.MessageBox("Error en MenuEvent: " & ex.Message)
        End Try
    End Sub

End Class
