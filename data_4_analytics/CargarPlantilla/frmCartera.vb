Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Xml
Imports System.Security.Permissions
Imports System.Windows.Forms
Imports SAPbouiCOM
Imports System.Globalization
Imports Microsoft.Office.Interop


Public Class frmCartera

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
    Dim oGrid As SAPbouiCOM.Grid
    Dim txtDocEntry As SAPbouiCOM.EditText
    Dim text_fecha As SAPbouiCOM.EditText
    Dim btnConsultar As SAPbouiCOM.Button
    Dim txtTotCartera As SAPbouiCOM.EditText
    Dim txtTotAbono As SAPbouiCOM.EditText


    Dim columnasEsperadas As Integer = 6

    Sub New(ByVal Company As SAPbobsCOM.Company, ByVal sboApp As SAPbouiCOM.Application)
        rCompany = Company
        rsboApp = sboApp
    End Sub

    Public Sub CargaFormCartera()
        Dim xmlDoc As New Xml.XmlDocument
        Dim strPath As String

        If RecorreFormulario(rsboApp, "frmCartera") Then
            Exit Sub
        End If

        strPath = System.Windows.Forms.Application.StartupPath & "\frmCartera.srf"
        xmlDoc.Load(strPath)

        Try
            Try
                rsboApp.LoadBatchActions(xmlDoc.InnerXml)
            Catch exx As Exception
                rsboApp.Forms.Item("frmCartera").Close()
                xmlDoc = Nothing
                Exit Sub
            End Try

            oForm = rsboApp.Forms.Item("frmCartera")
            oForm.Freeze(True)

            Dim ipLogo As SAPbouiCOM.PictureBox
            ipLogo = oForm.Items.Item("ipLogo").Specific
            ipLogo.Picture = Windows.Forms.Application.StartupPath & "\LogoSS.png"

            txtDocEntry = oForm.Items.Item("txt_DEntry").Specific
            txtDocEntry.Item.Enabled = False

            'txtTotCartera = oForm.Items.Item("txtTCart").Specific
            'txtTotAbono = oForm.Items.Item("txtTAbon").Specific


            oMatrix = oForm.Items.Item("MTX_UDO").Specific

            oForm.Mode = BoFormMode.fm_ADD_MODE

            btnConsultar = oForm.Items.Item("btn_cons").Specific

            If oForm.Mode = BoFormMode.fm_ADD_MODE Then
                btnConsultar.Item.Enabled = True
            Else
                btnConsultar.Item.Enabled = False
            End If


            text_fecha = oForm.Items.Item("text_fecha").Specific
            text_fecha.Value = DateTime.Now.ToString("yyyyMMdd")
            text_fecha.Item.Enabled = False

            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Dim query As String = "EXEC ""Cartera_SP"""
            'Dim query As String = "SELECT * FROM ""_SYS_BIC"".""sap.sbose/SV_ANTIGUEDAD_CARTERA_CLIENTES"""
            Dim query As String = "SELECT * FROM ""Cartera"""
            oRecordSet.DoQuery(query)

            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_CARTERA_DET1")
            oDataSource.Clear()

            Dim i As Integer = 0

            While Not oRecordSet.EoF
                oDataSource.InsertRecord(i)
                oDataSource.SetValue("LineId", i, (i + 1).ToString)

                oDataSource.SetValue("U_DocNum", i, oRecordSet.Fields.Item("DocNum").Value)
                oDataSource.SetValue("U_CardCode", i, oRecordSet.Fields.Item("CardCode").Value)
                oDataSource.SetValue("U_CardName", i, oRecordSet.Fields.Item("CardName").Value)
                oDataSource.SetValue("U_Tipo_Doc", i, oRecordSet.Fields.Item("TIPO_DOC").Value)
                oDataSource.SetValue("U_FolioNum", i, oRecordSet.Fields.Item("FolioNum").Value)

                Dim fechaCont As String = ""
                Dim fechaVenc As String = ""

                If Not IsDBNull(oRecordSet.Fields.Item("DocDate").Value) Then
                    Dim f As Date = oRecordSet.Fields.Item("DocDate").Value
                    fechaCont = f.ToString("yyyyMMdd")
                End If

                If Not IsDBNull(oRecordSet.Fields.Item("DocDueDate").Value) Then
                    Dim f As Date = oRecordSet.Fields.Item("DocDueDate").Value
                    fechaVenc = f.ToString("yyyyMMdd")
                End If

                oDataSource.SetValue("U_DocDate", i, fechaCont)
                oDataSource.SetValue("U_DocDueDate", i, fechaVenc)
                oDataSource.SetValue("U_Total_Cart", i, oRecordSet.Fields.Item("TOTALCARTERA").Value)
                oDataSource.SetValue("U_Dias_Venc", i, oRecordSet.Fields.Item("DIA_VENCIMIENTO").Value)
                oDataSource.SetValue("U_CartxVenc", i, oRecordSet.Fields.Item("CARTERA_POR_VENCER").Value)
                oDataSource.SetValue("U_Cart_Venc", i, oRecordSet.Fields.Item("CARTERA_VENCIDA").Value)
                oDataSource.SetValue("U_TotalAbon", i, oRecordSet.Fields.Item("TOTAL_ABONO").Value)
                oDataSource.SetValue("U_Venc_30", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_30").Value)
                oDataSource.SetValue("U_Venc_60", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_60").Value)
                oDataSource.SetValue("U_Venc_90", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_90").Value)
                oDataSource.SetValue("U_Venc_120", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_120").Value)
                oDataSource.SetValue("U_Venc_m120", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_MAS_120").Value)
                oDataSource.SetValue("U_CodVend", i, oRecordSet.Fields.Item("SlpCode").Value)
                oDataSource.SetValue("U_NomVend", i, oRecordSet.Fields.Item("SlpName").Value)
                oDataSource.SetValue("U_NumAut", i, oRecordSet.Fields.Item("NumAtCard").Value)
                oDataSource.SetValue("U_Denomi", i, oRecordSet.Fields.Item("U_Denominacion").Value)
                oDataSource.SetValue("U_Estab", i, oRecordSet.Fields.Item("U_SS_Est").Value)
                oDataSource.SetValue("U_PtoEm", i, oRecordSet.Fields.Item("U_SS_Pemi").Value)
                oDataSource.SetValue("U_ObjType", i, oRecordSet.Fields.Item("ObjType").Value)
                oDataSource.SetValue("U_NumInter", i, oRecordSet.Fields.Item("DocEntry").Value)


                oRecordSet.MoveNext()
                i += 1
            End While

            oMatrix.LoadFromDataSource()

            oMatrix.Columns.Item("U21CVend").Visible = False
            oMatrix.Columns.Item("V22NVend").Visible = False
            oMatrix.Columns.Item("X23Estb").Visible = False
            oMatrix.Columns.Item("Y24PtoEm").Visible = False
            oMatrix.Columns.Item("H8Folio").Visible = False
            oMatrix.Columns.Item("Z25Obj").Visible = False

            'Dim totalCartera As Double = 0
            'Dim totalAbono As Double = 0

            'For p As Integer = 1 To oMatrix.RowCount
            '    Dim valCart As Double = 0
            '    Dim valAbono As Double = 0

            '    Double.TryParse(oMatrix.Columns.Item("K11TCart").Cells.Item(p).Specific.Value, valCart)
            '    Double.TryParse(oMatrix.Columns.Item("O15TAbo").Cells.Item(p).Specific.Value, valAbono)

            '    totalCartera += valCart
            '    totalAbono += valAbono
            'Next

            'txtTotCartera.Value = totalCartera.ToString("N2")
            'txtTotAbono.Value = totalAbono.ToString("N2")



            'Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            'oEditTextColumn = oMatrix.Columns.Item("D4CrdCod")
            'oEditTextColumn.LinkedObjectType = "2"

            oForm.Visible = True
            oForm.Select()

            rsboApp.StatusBar.SetText("Datos cargados correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            rsboApp.MessageBox("Error al cargar la pantalla: " & ex.Message)
        Finally
            oForm.Freeze(False)
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
            If pVal.FormUID = "frmCartera" Then
                Select Case pVal.ItemUID
                    Case "btn_cons"
                        If pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                LlenarMatrixCartera()
                            Else
                                Dim fechaFormulario As String = text_fecha.Value.Trim
                                Dim fechaHoy As String = DateTime.Now.ToString("yyyyMMdd")
                                If fechaFormulario = fechaHoy Then
                                    LlenarMatrixCartera()
                                Else
                                    rsboApp.StatusBar.SetText("No se puede actualizar documentos con fechas anteriores.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                            End If
                        End If

                    Case "1"
                        If pVal.BeforeAction = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then

                                If oMatrix.RowCount = 0 OrElse oMatrix.VisualRowCount = 0 Then
                                    rsboApp.MessageBox("No se puede grabar. La Matrix está vacía.")
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                Dim fechaCorte As String = text_fecha.Value
                                Dim fechaCorteConvert As String = DateTime.ParseExact(fechaCorte, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("dd-MM-yyyy")

                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    Dim ConRegistro As String = ""
                                    If rCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        ConRegistro = "select COUNT(1) as ""Contador"" from """ + rCompany.CompanyDB + """.""@SS_CARTERA_CAB"" where ""U_FECHA""='" + fechaCorte + "'"
                                    Else
                                        ConRegistro = "select COUNT(1) as ""Contador"" from ""@SS_CARTERA_CAB"" where ""U_FECHA""='" + fechaCorte + "'"
                                    End If

                                    Dim Val As Integer = CInt(oFuncionesB1.getRSvalue(ConRegistro, "Contador", "0"))
                                    If Val > 0 Then
                                        BubbleEvent = False
                                        rsboApp.StatusBar.SetText("Solo se permite un registro por día, ya existe un registro con fecha " + fechaCorteConvert + "..!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                End If

                                Dim respuesta As Integer = rsboApp.MessageBox(
                                    "¿Está seguro de grabar/actualizar con la Fecha de Corte: " & fechaCorteConvert & " ?", 1, "Sí", "No")

                                If respuesta = 2 Then ' Si elige No
                                    BubbleEvent = False ' Cancela el evento
                                    Exit Sub
                                End If

                            End If

                        End If
                    Case "MTX_UDO" 'SAPbouiCOM.BoEventTypes.et_LINK_PRESSED
                        If pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED And pVal.ColUID = "A1DEntry" Then
                            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("MTX_UDO").Specific
                            Dim docEntry As String = oMatrix.Columns.Item("A1DEntry").Cells.Item(pVal.Row).Specific.value
                            Dim objType As String = oMatrix.Columns.Item("Z25Obj").Cells.Item(pVal.Row).Specific.value

                            If docEntry <> "" AndAlso objType <> "" Then
                                rsboApp.OpenForm(objType, "", docEntry)
                            Else
                                rsboApp.StatusBar.SetText("No se encontró ObjType o DocNum.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                End Select

            End If

        Catch ex As Exception
            rsboApp.MessageBox("Error en evento: " & ex.Message)
        End Try
    End Sub

    Private Function GetFormEnum(ByVal objType As String) As SAPbouiCOM.BoFormObjectEnum
        Select Case objType
            Case "13"
                Return SAPbouiCOM.BoFormObjectEnum.fo_Invoice
            Case "14"
                Return SAPbouiCOM.BoFormObjectEnum.fo_InvoiceCreditMemo
            Case Else
                Throw New Exception("ObjType no reconocido: " & objType)
        End Select
    End Function


    Public Sub LlenarMatrixCartera()
        Try
            rsboApp.StatusBar.SetText("Consultando Cartera, por favor espere...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oForm.Freeze(True)

            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = rCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Dim query As String = "EXEC ""Cartera_SP"""
            'Dim query As String = "SELECT * FROM ""_SYS_BIC"".""sap.sbose/SV_ANTIGUEDAD_CARTERA_CLIENTES"""
            Dim query As String = "SELECT * FROM ""Cartera"""
            oRecordSet.DoQuery(query)

            Dim oDataSource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("@SS_CARTERA_DET1")

            oDataSource.Clear()
            oMatrix.Clear()

            Dim i As Integer = 0

            While Not oRecordSet.EoF
                oDataSource.InsertRecord(i)
                oDataSource.SetValue("LineId", i, (i + 1).ToString)

                oDataSource.SetValue("U_DocNum", i, oRecordSet.Fields.Item("DocNum").Value)
                oDataSource.SetValue("U_CardCode", i, oRecordSet.Fields.Item("CardCode").Value)
                oDataSource.SetValue("U_CardName", i, oRecordSet.Fields.Item("CardName").Value)
                oDataSource.SetValue("U_Tipo_Doc", i, oRecordSet.Fields.Item("TIPO_DOC").Value)
                oDataSource.SetValue("U_FolioNum", i, oRecordSet.Fields.Item("FolioNum").Value)

                Dim fechaCont As String = ""
                Dim fechaVenc As String = ""

                If Not IsDBNull(oRecordSet.Fields.Item("DocDate").Value) Then
                    Dim f As Date = oRecordSet.Fields.Item("DocDate").Value
                    fechaCont = f.ToString("yyyyMMdd")
                End If

                If Not IsDBNull(oRecordSet.Fields.Item("DocDueDate").Value) Then
                    Dim f As Date = oRecordSet.Fields.Item("DocDueDate").Value
                    fechaVenc = f.ToString("yyyyMMdd")
                End If

                oDataSource.SetValue("U_DocDate", i, fechaCont)
                oDataSource.SetValue("U_DocDueDate", i, fechaVenc)

                oDataSource.SetValue("U_Total_Cart", i, oRecordSet.Fields.Item("TOTALCARTERA").Value)
                oDataSource.SetValue("U_Dias_Venc", i, oRecordSet.Fields.Item("DIA_VENCIMIENTO").Value)
                oDataSource.SetValue("U_CartxVenc", i, oRecordSet.Fields.Item("CARTERA_POR_VENCER").Value)
                oDataSource.SetValue("U_Cart_Venc", i, oRecordSet.Fields.Item("CARTERA_VENCIDA").Value)
                oDataSource.SetValue("U_TotalAbon", i, oRecordSet.Fields.Item("TOTAL_ABONO").Value)
                oDataSource.SetValue("U_Venc_30", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_30").Value)
                oDataSource.SetValue("U_Venc_60", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_60").Value)
                oDataSource.SetValue("U_Venc_90", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_90").Value)
                oDataSource.SetValue("U_Venc_120", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_120").Value)
                oDataSource.SetValue("U_Venc_m120", i, oRecordSet.Fields.Item("CARTERA_VENCIDA_MAS_120").Value)
                oDataSource.SetValue("U_CodVend", i, oRecordSet.Fields.Item("SlpCode").Value)
                oDataSource.SetValue("U_NomVend", i, oRecordSet.Fields.Item("SlpName").Value)
                oDataSource.SetValue("U_NumAut", i, oRecordSet.Fields.Item("NumAtCard").Value)
                oDataSource.SetValue("U_Denomi", i, oRecordSet.Fields.Item("U_Denominacion").Value)
                oDataSource.SetValue("U_Estab", i, oRecordSet.Fields.Item("U_SS_Est").Value)
                oDataSource.SetValue("U_PtoEm", i, oRecordSet.Fields.Item("U_SS_Pemi").Value)
                oDataSource.SetValue("U_ObjType", i, oRecordSet.Fields.Item("ObjType").Value)
                oDataSource.SetValue("U_NumInter", i, oRecordSet.Fields.Item("DocEntry").Value)

                oRecordSet.MoveNext()
                i += 1
            End While

            oMatrix.LoadFromDataSource()

            oMatrix.Columns.Item("U21CVend").Visible = False
            oMatrix.Columns.Item("V22NVend").Visible = False
            oMatrix.Columns.Item("X23Estb").Visible = False
            oMatrix.Columns.Item("Y24PtoEm").Visible = False
            oMatrix.Columns.Item("H8Folio").Visible = False
            oMatrix.Columns.Item("Z25Obj").Visible = False

            rsboApp.StatusBar.SetText("Datos cargados correctamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            rsboApp.MessageBox("Error al llenar la Matrix: " & ex.Message)
        Finally
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                text_fecha = oForm.Items.Item("text_fecha").Specific
                text_fecha.Value = DateTime.Now.ToString("yyyyMMdd")
            End If

            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
            oForm.Freeze(False)
        End Try
    End Sub


End Class
