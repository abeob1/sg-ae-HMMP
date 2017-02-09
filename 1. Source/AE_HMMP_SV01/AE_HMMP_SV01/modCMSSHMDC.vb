Module modCMSSHMDC

    Private dtCardCode As DataTable
    Private dtItemCode As DataTable
    Private dtInvoice_ARDetails As DataTable
    Private dtInvoice_APDetails As DataTable

    Public Function ProcessHMDCDatas(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessHMDCDatas"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSql = "SELECT ""CardCode"",""VatGroup"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OCRD"""
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtCardCode = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMDCSAPDbName)

            sSql = "SELECT ""ItemCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OITM"""
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtItemCode = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMDCSAPDbName)

            sSql = "SELECT ""U_invoice"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AR_DETAILS"" WHERE IFNULL(""U_Inv_DocEntry"",'') <> ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtInvoice_ARDetails = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMDCSAPDbName)

            sSql = "SELECT ""U_invoice"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AP_DETAILS"" WHERE IFNULL(""U_AP_Inv_DocEntry"",'') <> ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtInvoice_APDetails = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMDCSAPDbName)

            Dim odtDatatable As DataTable
            odtDatatable = oDv.ToTable
            odtDatatable.Columns.Add("IncuredMonth", GetType(Date))
            odtDatatable.Columns.Add("ArCode", GetType(String))
            odtDatatable.Columns.Add("ApCode", GetType(String))
            odtDatatable.Columns.Add("InvoiceDate", GetType(Date))
            odtDatatable.Columns.Add("CostCenter", GetType(String))

            Dim sFileDate As String = file.Name.Substring(11, 8)

            For intRow As Integer = 0 To odtDatatable.Rows.Count - 1
                If Not (odtDatatable.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or odtDatatable.Rows(intRow).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    Console.WriteLine("Processing excel line " & intRow)

                    Dim sCompCode As String = odtDatatable.Rows(intRow).Item(25).ToString
                    If sCompCode = "" Then
                        sErrDesc = "Company Code should not be empty / Check Line " & intRow
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine(sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Dim sTreatment As String = odtDatatable.Rows(intRow).Item(11).ToString
                    sTreatment = sTreatment.Replace("'", " ")

                    Dim sInvoice As String = odtDatatable.Rows(intRow).Item(0).ToString.Trim
                    dtInvoice_ARDetails.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtInvoice_ARDetails.DefaultView.Count > 0 Then
                        sErrDesc = "A/R Invoice has been created previously for invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    dtInvoice_APDetails.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtInvoice_APDetails.DefaultView.Count > 0 Then
                        sErrDesc = "A/p Invoice has been created previously for invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Dim sArCode As String = "C" & sCompCode
                    Dim sClinicCode As String = odtDatatable.Rows(intRow).Item(1).ToString
                    Dim sSubCode As String = odtDatatable.Rows(intRow).Item(2).ToString
                    Dim sAPCode As String = "V" & sClinicCode & sSubCode

                    Dim iIndex As Integer = odtDatatable.Rows(intRow).Item(4).ToString.IndexOf(" ")
                    Dim sDate As String = odtDatatable.Rows(intRow).Item(4).ToString.Substring(0, iIndex)
                    Dim dt As Date
                    Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                    Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

                    Dim dIncurMnth As Date = CDate(dt.Date.AddDays(-(dt.Day - 1)).AddMonths(1).AddDays(-1).ToString())

                    Dim dInvoiceDate As Date
                    Date.TryParseExact(sFileDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dInvoiceDate)

                    odtDatatable.Rows(intRow)("F2") = sClinicCode.ToUpper()
                    odtDatatable.Rows(intRow)("F12") = sTreatment
                    odtDatatable.Rows(intRow)("IncuredMonth") = dIncurMnth
                    odtDatatable.Rows(intRow)("ArCode") = sArCode
                    odtDatatable.Rows(intRow)("ApCode") = sAPCode
                    odtDatatable.Rows(intRow)("InvoiceDate") = dInvoiceDate
                End If
            Next

            Dim oDvFinalView As DataView
            oDvFinalView = New DataView(odtDatatable)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToCompany(p_oCompany, p_oCompDef.sHMDCSAPDbName, p_oCompDef.sHMDCSAPUserName, p_oCompDef.sHMDCSAPPassword, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then
                Console.WriteLine("Company connection Successful")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction", sFuncName)

                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If oDvFinalView.Count > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoTable_RD001_AR()", sFuncName)

                    Console.WriteLine("Inserting datas in YOT Table")
                    If InsertIntoTable_RD001_AR(oDvFinalView, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Console.WriteLine("Insert into YOT table successful")

                    oDvFinalView.RowFilter = "F34 LIKE 'CONTRACT*' AND F16 <> '0'"
                    Dim odtContract As New DataTable
                    odtContract = oDvFinalView.ToTable

                    Dim oContract As DataView = New DataView(odtContract)

                    '*************PROCESSING DATAS IF LESS_DIS_PAY_CLIENT HAS AMOUNT FOR CUSTOMER TYPE IS CONTRACT*****************************
                    If oContract.Count > 0 Then

                        oContract.RowFilter = Nothing
                        'F2 - Cln_Code
                        Dim oDtGroup As DataTable = oContract.Table.DefaultView.ToTable(True, "F1", "F2", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                                oContract.RowFilter = "F1='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F2 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "

                                If oContract.Count > 0 Then
                                    Console.WriteLine("Inserting values into AP Details table")
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoAPDetails()", sFuncName)
                                    Dim oDt As DataTable
                                    oDt = oContract.ToTable
                                    Dim oApInvDv As DataView = New DataView(oDt)
                                    If InsertIntoAPDetails(oApInvDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Insertion of values in AP Details table is successful")
                                End If
                            End If
                        Next

                        oContract.RowFilter = Nothing

                        oDtGroup = oContract.Table.DefaultView.ToTable(True, "F2", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oContract.RowFilter = "F2='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "

                                If oContract.Count > 0 Then
                                    Console.WriteLine("Creating A/p Invoice to HMMPD")
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateAPInvoice_Contract()", sFuncName)
                                    Dim oApInvDt As DataTable
                                    oApInvDt = oContract.ToTable
                                    Dim oApInvDv As DataView = New DataView(oApInvDt)
                                    If CreateAPInvoice_Contract_HMMPD(oApInvDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Invoice Creation for another database is successful")

                                End If
                            End If
                        Next

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping Datas to insert into Cash sales", sFuncName)

                        oContract.RowFilter = Nothing

                        'F1 Invoice F2 Clinic Code, F25 Payment Method
                        oDtGroup = oContract.Table.DefaultView.ToTable(True, "F1", "F2", "IncuredMonth", "F25")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                                oContract.RowFilter = "F1='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F2 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "'                  " & _
                                                             " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' and F25 ='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "'"
                                If oContract.Count > 0 Then
                                    Console.WriteLine("Inserting data into Cash table for " & oDtGroup.Rows(i).Item(1).ToString.Trim())
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoCashSales()", sFuncName)
                                    Dim oCashDt As DataTable
                                    oCashDt = oContract.ToTable
                                    Dim oCashDv As DataView = New DataView(oCashDt)
                                    If InsertIntoCashSales(oCashDv, p_oCompany, "LESS_DIS_PAY_CLIENT", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Inserting table into cash table is successful")
                                End If
                            End If
                        Next

                        oContract.RowFilter = Nothing

                        oDtGroup = oContract.Table.DefaultView.ToTable(True, "F25")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "PAYMETHOD") Then

                                Dim sPayMethod As String = oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim()
                                sSql = "SELECT COUNT(""U_PayMethod"") AS ""MNO"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_PAYMETHOD"" WHERE UPPER(""U_PayMethod"") = '" & sPayMethod & "'"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                                Dim iCount As Integer = GetCode(sSql, p_oCompDef.sHMDCSAPDbName)

                                oContract.RowFilter = "F25 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' "

                                If iCount > 0 Then
                                    If oContract.Count > 0 Then
                                        Dim oDtContract_inPay As DataTable
                                        oDtContract_inPay = oContract.ToTable()
                                        Dim oDvContract_inPay As DataView = New DataView(oDtContract_inPay)

                                        Dim oDt_Grouped As DataTable = oDvContract_inPay.Table.DefaultView.ToTable(True, "F2", "IncuredMonth")
                                        For k As Integer = 0 To oDt_Grouped.Rows.Count - 1
                                            If Not (oDt_Grouped.Rows(k).Item(0).ToString.Trim = String.Empty Or oDt_Grouped.Rows(k).Item(1).ToString.ToUpper.Trim() = "CLN_CODE") Then
                                                oDvContract_inPay.RowFilter = "F2 = '" & oDt_Grouped.Rows(k).Item(0).ToString.Trim() & "' AND IncuredMonth = '" & oDt_Grouped.Rows(k).Item(1).ToString.Trim() & "'"

                                                If oDvContract_inPay.Count > 0 Then
                                                    Dim oDtContract_inpay_inv As DataTable
                                                    oDtContract_inpay_inv = oDvContract_inPay.ToTable()
                                                    Dim oDvContract_inpay_inv As DataView = New DataView(oDtContract_inpay_inv)

                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoicePayment_Contract()", sFuncName)
                                                    Console.WriteLine("Creating invoice+payment for contract datas in payment table")
                                                    If CreateARInvoicePayment_Contract(oDvContract_inpay_inv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                                    Console.WriteLine("Creation of Invoice+Payment for contract datas in payment table is successful")
                                                End If
                                            End If
                                        Next
                                    End If
                                Else
                                    If oContract.Count > 0 Then
                                        Dim oDtContract_NotInPay As DataTable
                                        oDtContract_NotInPay = oContract.ToTable()
                                        Dim oDvContract_NotInPay As DataView = New DataView(oDtContract_NotInPay)

                                        Dim oDt_Grouped As DataTable = oDvContract_NotInPay.Table.DefaultView.ToTable(True, "F2", "IncuredMonth")

                                        For k As Integer = 0 To oDt_Grouped.Rows.Count - 1
                                            If Not (oDt_Grouped.Rows(k).Item(0).ToString.Trim = String.Empty Or oDt_Grouped.Rows(k).Item(1).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                                oDvContract_NotInPay.RowFilter = "F2 = '" & oDt_Grouped.Rows(k).Item(0).ToString.Trim() & "' AND IncuredMonth = '" & oDt_Grouped.Rows(k).Item(1).ToString.Trim() & "'"
                                                If oDvContract_NotInPay.Count > 0 Then
                                                    Dim oDtContract_NotInPay_inv As DataTable
                                                    oDtContract_NotInPay_inv = oDvContract_NotInPay.ToTable()
                                                    Dim oDvContract_NotInPay_inv As DataView = New DataView(oDtContract_NotInPay_inv)

                                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateArInvoice()", sFuncName)
                                                    Console.WriteLine("Creating invoice for contract datas which are not in payment table")
                                                    If CreateArInvoice(oDvContract_NotInPay_inv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                                    Console.WriteLine("Invoice creation successful for contract datas which are not in payment table")
                                                End If
                                            End If
                                        Next
                                    End If
                                End If

                            End If
                        Next

                    End If

                    oDvFinalView.RowFilter = Nothing

                    '******************PROCESSING DATAS FOR ROWS WHICH PAY_COMP VALUES GREATER THAN ZERO**************************
                    oDvFinalView.RowFilter = "F34 LIKE 'CONTRACT*' AND F14 <> '0'"
                    Dim odtContract_PayClient As New DataTable
                    odtContract_PayClient = oDvFinalView.ToTable

                    Dim oContract_PayClient As DataView = New DataView(odtContract_PayClient)

                    If oContract_PayClient.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping Datas to insert into Cash sales", sFuncName)

                        'F2 Clinic Code, F25 Payment Method
                        Dim oDtGroup As DataTable = oContract_PayClient.Table.DefaultView.ToTable(True, "F2", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oContract_PayClient.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "
                                If oContract_PayClient.Count > 0 Then
                                    Console.WriteLine("Creating AR invoice")
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoice_HMMPD()", sFuncName)
                                    Dim oDtPayClient_Contract As DataTable
                                    oDtPayClient_Contract = oContract_PayClient.ToTable
                                    Dim oDvPayClient_Contract As DataView = New DataView(oDtPayClient_Contract)
                                    If CreateARInvoice_HMMPD(oDvPayClient_Contract, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("AR Invoice Creation successful")
                                End If
                            End If
                        Next
                    End If

                    '************PROCESSING DATAS FOR NON CONTRACT CUSTOMER TYPE*******************
                    oDvFinalView.RowFilter = Nothing

                    oDvFinalView.RowFilter = "F34 NOT LIKE 'CONTRACT*'"
                    Dim odtNonContract As New DataTable
                    odtNonContract = oDvFinalView.ToTable

                    Dim oNonContractDv As DataView = New DataView(odtNonContract)
                    If oNonContractDv.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping by payment group", sFuncName)

                        Dim oDtGroup As DataTable = oNonContractDv.Table.DefaultView.ToTable(True, "F25")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "PAYMETHOD") Then
                                Dim sPayMethod As String = oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim()
                                sSql = "SELECT COUNT(""U_PayMethod"") AS ""MNO"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_PAYMETHOD"" WHERE UPPER(""U_PayMethod"") = '" & sPayMethod & "'"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                                Dim iCount As Integer = GetCode(sSql, p_oCompDef.sHMDCSAPDbName)

                                oNonContractDv.RowFilter = "F25 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' "
                                If iCount > 0 Then
                                    If oNonContractDv.Count > 0 Then
                                        Console.WriteLine("Invoice Creation for cash table")
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessNonContractDatas()", sFuncName)
                                        Dim oNonContract_PaymentDt As DataTable
                                        oNonContract_PaymentDt = oNonContractDv.ToTable
                                        Dim oNonContract_PaymentDv As DataView = New DataView(oNonContract_PaymentDt)
                                        If ProcessNonContractDatas(oNonContract_PaymentDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Console.WriteLine("Invoice Creation for Cash table successful")
                                    End If
                                Else
                                    If oNonContractDv.Count > 0 Then
                                        Console.WriteLine("Invoice Creation for cash table")
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessNonContract_NonPay()", sFuncName)
                                        Dim oNonContract_PaymentDt As DataTable
                                        oNonContract_PaymentDt = oNonContractDv.ToTable
                                        Dim oNonContract_PaymentDv As DataView = New DataView(oNonContract_PaymentDt)
                                        If ProcessNonContract_NonPay(oNonContract_PaymentDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Console.WriteLine("Invoice Creation for Cash table successful")
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
            If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_SUCCESS)

            'Insert Success Notificaiton into Table..
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtSuccess, file.Name, "Success")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & file.FullName, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessHMDCDatas = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_ERROR)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessHMDCDatas = RTN_ERROR
        End Try
    End Function

    Public Function ProcessHMDCDatas_BACKUP(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessHMDCDatas"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSql = "SELECT ""CardCode"",""VatGroup"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OCRD"""
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtCardCode = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMDCSAPDbName)

            sSql = "SELECT ""ItemCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OITM"""
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtItemCode = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMDCSAPDbName)

            sSql = "SELECT ""U_invoice"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AR_DETAILS"" WHERE IFNULL(""U_Inv_DocEntry"",'') <> ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtInvoice_ARDetails = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMDCSAPDbName)

            sSql = "SELECT ""U_invoice"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AP_DETAILS"" WHERE IFNULL(""U_AP_Inv_DocEntry"",'') <> ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtInvoice_APDetails = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMDCSAPDbName)

            Dim odtDatatable As DataTable
            odtDatatable = oDv.ToTable
            odtDatatable.Columns.Add("IncuredMonth", GetType(Date))
            odtDatatable.Columns.Add("ArCode", GetType(String))
            odtDatatable.Columns.Add("ApCode", GetType(String))
            odtDatatable.Columns.Add("InvoiceDate", GetType(Date))
            odtDatatable.Columns.Add("CostCenter", GetType(String))

            Dim sFileDate As String = file.Name.Substring(11, 8)

            For intRow As Integer = 0 To odtDatatable.Rows.Count - 1
                If Not (odtDatatable.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or odtDatatable.Rows(intRow).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    Console.WriteLine("Processing excel line " & intRow)

                    Dim sCompCode As String = odtDatatable.Rows(intRow).Item(25).ToString
                    If sCompCode = "" Then
                        sErrDesc = "Company Code should not be empty / Check Line " & intRow
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine(sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Dim sInvoice As String = odtDatatable.Rows(intRow).Item(0).ToString.Trim
                    dtInvoice_ARDetails.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtInvoice_ARDetails.DefaultView.Count > 0 Then
                        sErrDesc = "A/R Invoice has been created previously for invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    dtInvoice_APDetails.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtInvoice_APDetails.DefaultView.Count > 0 Then
                        sErrDesc = "A/p Invoice has been created previously for invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Dim sArCode As String = "C" & sCompCode
                    Dim sClinicCode As String = odtDatatable.Rows(intRow).Item(1).ToString
                    Dim sSubCode As String = odtDatatable.Rows(intRow).Item(2).ToString
                    Dim sAPCode As String = "V" & sClinicCode & sSubCode

                    Dim iIndex As Integer = odtDatatable.Rows(intRow).Item(4).ToString.IndexOf(" ")
                    Dim sDate As String = odtDatatable.Rows(intRow).Item(4).ToString.Substring(0, iIndex)
                    Dim dt As Date
                    Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                    Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

                    Dim dIncurMnth As Date = CDate(dt.Date.AddDays(-(dt.Day - 1)).AddMonths(1).AddDays(-1).ToString())

                    Dim dInvoiceDate As Date
                    Date.TryParseExact(sFileDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dInvoiceDate)

                    odtDatatable.Rows(intRow)("IncuredMonth") = dIncurMnth
                    odtDatatable.Rows(intRow)("ArCode") = sArCode
                    odtDatatable.Rows(intRow)("ApCode") = sAPCode
                    odtDatatable.Rows(intRow)("InvoiceDate") = dInvoiceDate
                End If
            Next

            Dim oDvFinalView As DataView
            oDvFinalView = New DataView(odtDatatable)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToCompany(p_oCompany, p_oCompDef.sHMDCSAPDbName, p_oCompDef.sHMDCSAPUserName, p_oCompDef.sHMDCSAPPassword, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then
                Console.WriteLine("Company connection Successful")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction", sFuncName)

                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If oDvFinalView.Count > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoTable_RD001_AR()", sFuncName)

                    Console.WriteLine("Inserting datas in YOT Table")
                    If InsertIntoTable_RD001_AR(oDvFinalView, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Console.WriteLine("Insert into YOT table successful")

                    oDvFinalView.RowFilter = "F34 LIKE 'CONTRACT*' AND F16 > '0'"
                    Dim odtContract As New DataTable
                    odtContract = oDvFinalView.ToTable

                    Dim oContract As DataView = New DataView(odtContract)

                    '*************PROCESSING DATAS IF LESS_DIS_PAY_CLIENT HAS AMOUNT FOR CUSTOMER TYPE IS CONTRACT*****************************
                    If oContract.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping Datas to insert into Cash sales", sFuncName)

                        'F1 Invoice F2 Clinic Code, F25 Payment Method
                        Dim oDtGroup As DataTable = oContract.Table.DefaultView.ToTable(True, "F1", "ArCode", "F2", "IncuredMonth", "F25")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                                oContract.RowFilter = "F1='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and ArCode = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' and F2='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' " & _
                                                             " and IncuredMonth='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "' and F25 ='" & oDtGroup.Rows(i).Item(4).ToString.Trim() & "'"
                                If oContract.Count > 0 Then
                                    Console.WriteLine("Inserting data into Cash table for " & oDtGroup.Rows(i).Item(1).ToString.Trim())
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoCashSales()", sFuncName)
                                    Dim oCashDt As DataTable
                                    oCashDt = oContract.ToTable
                                    Dim oCashDv As DataView = New DataView(oCashDt)
                                    If InsertIntoCashSales(oCashDv, p_oCompany, "LESS_DIS_PAY_CLIENT", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Inserting table into cash table is successful")
                                End If
                            End If
                        Next

                        oContract.RowFilter = Nothing
                        oDtGroup = oContract.Table.DefaultView.ToTable(True, "F2", "F25", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oContract.RowFilter = "F2='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F25 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                             " and IncuredMonth = '" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "
                                If oContract.Count > 0 Then
                                    Console.WriteLine("Creating Invoice for cash table")
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoice_Contract()", sFuncName)
                                    Dim oCashTab_InvoiceDt As DataTable
                                    oCashTab_InvoiceDt = oContract.ToTable
                                    Dim oCashTab_InvoiceDv As DataView = New DataView(oCashTab_InvoiceDt)
                                    If CreateARInvoicePayment_Contract(oCashTab_InvoiceDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Invoice Creation for Cash table successful")
                                End If
                            End If
                        Next

                        oContract.RowFilter = Nothing
                        'F2 - Cln_Code
                        oDtGroup = oContract.Table.DefaultView.ToTable(True, "F1", "F2", "IncuredMonth", "ApCode")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                                oContract.RowFilter = "F1='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F2 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' " & _
                                                      " and ApCode = '" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "' "
                                If oContract.Count > 0 Then
                                    Console.WriteLine("Inserting values into AP Details table")
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoAPDetails()", sFuncName)
                                    Dim oDt As DataTable
                                    oDt = oContract.ToTable
                                    Dim oApInvDv As DataView = New DataView(oDt)
                                    If InsertIntoAPDetails(oApInvDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Insertion of values in AP Details table is successful")
                                End If
                            End If
                        Next

                        oContract.RowFilter = Nothing

                        oDtGroup = oContract.Table.DefaultView.ToTable(True, "F2", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oContract.RowFilter = "F2='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "

                                If oContract.Count > 0 Then
                                    Console.WriteLine("Creating A/p Invoice to HMMPD")
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateAPInvoice_Contract()", sFuncName)
                                    Dim oApInvDt As DataTable
                                    oApInvDt = oContract.ToTable
                                    Dim oApInvDv As DataView = New DataView(oApInvDt)
                                    If CreateAPInvoice_Contract_HMMPD(oApInvDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Invoice Creation for another database is successful")

                                End If
                            End If
                        Next
                    End If

                    oDvFinalView.RowFilter = Nothing

                    '******************PROCESSING DATAS FOR ROWS WHICH PAY_COMP VALUES GREATER THAN ZERO**************************
                    oDvFinalView.RowFilter = "F34 LIKE 'CONTRACT*' AND F14 > '0'"
                    Dim odtContract_PayClient As New DataTable
                    odtContract_PayClient = oDvFinalView.ToTable

                    Dim oContract_PayClient As DataView = New DataView(odtContract_PayClient)

                    If oContract_PayClient.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping Datas to insert into Cash sales", sFuncName)

                        'F2 Clinic Code, F25 Payment Method
                        Dim oDtGroup As DataTable = oContract_PayClient.Table.DefaultView.ToTable(True, "F2", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oContract_PayClient.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "
                                If oContract_PayClient.Count > 0 Then
                                    Console.WriteLine("Creating AR invoice")
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoice_HMMPD()", sFuncName)
                                    Dim oDtPayClient_Contract As DataTable
                                    oDtPayClient_Contract = oContract_PayClient.ToTable
                                    Dim oDvPayClient_Contract As DataView = New DataView(oDtPayClient_Contract)
                                    If CreateARInvoice_HMMPD(oDvPayClient_Contract, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("AR Invoice Creation successful")
                                End If
                            End If
                        Next
                    End If

                    '************PROCESSING DATAS FOR NON CONTRACT CUSTOMER TYPE*******************
                    oDvFinalView.RowFilter = Nothing

                    oDvFinalView.RowFilter = "F34 NOT LIKE 'CONTRACT*'"
                    Dim odtNonContract As New DataTable
                    odtNonContract = oDvFinalView.ToTable

                    Dim oNonContractDv As DataView = New DataView(odtNonContract)
                    If oNonContractDv.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping by payment group", sFuncName)

                        Dim oDtGroup As DataTable = oNonContractDv.Table.DefaultView.ToTable(True, "F25")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "PAYMETHOD") Then
                                Dim sPayMethod As String = oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim()
                                sSql = "SELECT COUNT(""U_PayMethod"") AS ""MNO"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_PAYMETHOD"" WHERE UPPER(""U_PayMethod"") = '" & sPayMethod & "'"
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                                Dim iCount As Integer = GetCode(sSql, p_oCompDef.sHMDCSAPDbName)

                                oNonContractDv.RowFilter = "F25 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' "
                                If iCount > 0 Then
                                    If oNonContractDv.Count > 0 Then
                                        Console.WriteLine("Invoice Creation for cash table")
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessNonContractDatas()", sFuncName)
                                        Dim oNonContract_PaymentDt As DataTable
                                        oNonContract_PaymentDt = oNonContractDv.ToTable
                                        Dim oNonContract_PaymentDv As DataView = New DataView(oNonContract_PaymentDt)
                                        If ProcessNonContractDatas(oNonContract_PaymentDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Console.WriteLine("Invoice Creation for Cash table successful")
                                    End If
                                Else
                                    If oNonContractDv.Count > 0 Then
                                        Console.WriteLine("Invoice Creation for cash table")
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessNonContract_NonPay()", sFuncName)
                                        Dim oNonContract_PaymentDt As DataTable
                                        oNonContract_PaymentDt = oNonContractDv.ToTable
                                        Dim oNonContract_PaymentDv As DataView = New DataView(oNonContract_PaymentDt)
                                        If ProcessNonContract_NonPay(oNonContractDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Console.WriteLine("Invoice Creation for Cash table successful")
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
            If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_SUCCESS)

            'Insert Success Notificaiton into Table..
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtSuccess, file.Name, "Success")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & file.FullName, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessHMDCDatas_BACKUP = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_ERROR)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessHMDCDatas_BACKUP = RTN_ERROR
        End Try
    End Function

    Private Function InsertIntoTable_RD001_AR(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "InsertIntoTable_RD001_AR"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For i As Integer = 1 To oDv.Count - 1
                If Not (oDv(i)(0).ToString.Trim = String.Empty) Then
                    Console.WriteLine("Inserting Line Num : " & i)
                    sSql = String.Empty

                    sSql = "INSERT INTO " & p_oCompDef.sHMDCSAPDbName & ".""@AE_RD001_AR"" (""Code"",""Name"",""U_invoice"",""U_cln_code"",""U_subcode"",""U_cln_name"", " & _
                            " ""U_txn_date"",""U_id_type"",""U_id"",""U_lastname"",""U_given_name"",""U_christian"",""U_treat_code"",""U_treatment"",""U_cost"", " & _
                            " ""U_pay_comp"",""U_pay_client"",""U_les_dis_pay_client"",""U_admin"",""U_reimburse"",""U_cmoney"",""U_treat_charge"",""U_less_dis_treat_chg"",""U_surface"", " & _
                            " ""U_tooth_no"",""U_discount"",""U_paymethod"",""U_company"",""U_scheme"",""U_is_referral"",""U_Office_Invoice"",""U_date"", " & _
                            " ""U_amt"",""U_issued_by"",""U_is_refund"",""U_Customer_type"",""U_incurred_month"",""U_ar_code"",""U_ap_code"",""U_invoice_date"") " & _
                            " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sHMDCSAPDbName & """.""@AE_RD001_AR""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sHMDCSAPDbName & """.""@AE_RD001_AR""), " & _
                            " '" & oDv(i)(0).ToString.Trim & "','" & oDv(i)(1).ToString.Trim & "','" & oDv(i)(2).ToString.Trim & "','" & oDv(i)(3).ToString.Trim & "'," & _
                            " '" & oDv(i)(4).ToString.Trim & "','" & oDv(i)(5).ToString.Trim & "','" & oDv(i)(6).ToString.Trim & "','" & oDv(i)(7).ToString.Trim & "'," & _
                            " '" & oDv(i)(8).ToString.Trim & "','" & oDv(i)(9).ToString.Trim & "','" & oDv(i)(10).ToString.Trim & "','" & oDv(i)(11).ToString.Trim & "'," & _
                            " '" & oDv(i)(12).ToString.Trim & "','" & oDv(i)(13).ToString.Trim & "','" & oDv(i)(14).ToString.Trim & "','" & oDv(i)(15).ToString.Trim & "'," & _
                            " '" & oDv(i)(16).ToString.Trim & "','" & oDv(i)(17).ToString.Trim & "','" & oDv(i)(18).ToString.Trim & "','" & oDv(i)(19).ToString.Trim & "'," & _
                            " '" & oDv(i)(20).ToString.Trim & "','" & oDv(i)(21).ToString.Trim & "','" & oDv(i)(22).ToString.Trim & "','" & oDv(i)(23).ToString.Trim & "'," & _
                            " '" & oDv(i)(24).ToString.Trim & "','" & oDv(i)(25).ToString.Trim & "','" & oDv(i)(26).ToString.Trim & "','" & oDv(i)(27).ToString.Trim & "'," & _
                            " '" & oDv(i)(28).ToString.Trim & "','" & oDv(i)(29).ToString.Trim & "','" & oDv(i)(30).ToString.Trim & "','" & oDv(i)(31).ToString.Trim & "'," & _
                            " '" & oDv(i)(32).ToString.Trim & "','" & oDv(i)(33).ToString.Trim & "','" & oDv(i)(34).ToString.Trim & "','" & oDv(i)(35).ToString.Trim & "'," & _
                            " '" & oDv(i)(36).ToString.Trim & "','" & oDv(i)(37).ToString.Trim & "' )"
                    oRecSet.DoQuery(sSql)
                End If
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoTable_RD001_AR = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            InsertIntoTable_RD001_AR = RTN_ERROR
        End Try
    End Function

    Private Function InsertIntoCashSales(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal sType As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "InsertIntoCashSales"
        Dim sInvoice, sClinicCode, sPayMethod, sArCode, sIncuredMnth, sSQL, sItemCode, sBank As String
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            'F1 Invoice F2 Clinic Code, F25 Payment Method ar code and incurred month

            sInvoice = oDv(0)(0).ToString.Trim
            sClinicCode = oDv(0)(1).ToString.Trim
            sPayMethod = oDv(0)(24).ToString.Trim
            sArCode = oDv(0)(35).ToString.Trim
            sIncuredMnth = oDv(0)(34).ToString.Trim

            sSQL = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_ITEMCODE"" " & _
                  " WHERE ""U_FileCode"" = 'YOT' AND UPPER(""U_Field"") = 'PAY_CLIENT' AND UPPER(""U_DocType"") = 'A/R' AND UPPER(""U_CustType"") = 'CASH' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            sItemCode = GetStringValue(sSQL, p_oCompDef.sHMDCSAPDbName)

            sSQL = "SELECT ""U_BankGL"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_PAYMETHOD"" WHERE UPPER(""U_PayMethod"") = '" & sPayMethod.ToUpper & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            sBank = GetStringValue(sSQL, p_oCompDef.sHMDCSAPDbName)

            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dInvoiceDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dInvoiceDate)

            Dim dAmount As Double = 0

            For i As Integer = 0 To oDv.Count - 1
                If sType = "LESS_DIS_PAY_CLIENT" Then
                    Try
                        If Not (oDv(i)(15).ToString.Trim = String.Empty) Then
                            dAmount = CDbl(oDv(i)(15).ToString.Trim)
                        End If
                    Catch ex As Exception
                        dAmount = 0.0
                    End Try
                ElseIf sType = "PAY_COMP" Then
                    Try
                        If Not (oDv(i)(13).ToString.Trim = String.Empty) Then
                            dAmount = CDbl(oDv(i)(13).ToString.Trim)
                        End If
                    Catch ex As Exception
                        dAmount = 0.0
                    End Try
                End If
            Next
            sSQL = "INSERT INTO " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AR_DETAILS""(""Code"",""Name"",""U_incurred_month"", ""U_invoice_date"", " & _
                       " ""U_invoice"",""U_amount"",""U_cln_code"",""U_subcode"",""U_ItemCode"",""U_paymethod"",""U_bank"",""U_invoice_type"",""U_ar_code"",""U_CostCenter"") " & _
                       " VALUES ((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AR_DETAILS""), " & _
                       " (SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AR_DETAILS""), " & _
                       " '" & sIncuredMnth & "','" & dInvoiceDate.ToString("yyyy-MM-dd") & "','" & oDv(0)(0).ToString & "','" & dAmount & "', " & _
                       " '" & oDv(0)(1).ToString & "','" & oDv(0)(2).ToString & "','" & sItemCode & "','" & sPayMethod & "','" & sBank & "','Cash Sales','" & sArCode & "','" & sClinicCode & "') "

            oRecordSet.DoQuery(sSQL)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoCashSales = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERORR", sFuncName)
            InsertIntoCashSales = RTN_ERROR
        End Try
    End Function

    Private Function CreateARInvoicePayment_Contract(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateARInvoice_Contract"
        Dim sItemCode As String = String.Empty
        Dim sSql As String = String.Empty
        Dim dAmount As Double = 0.0
        Dim sClinicCode, sPayMethod, sArCode, sIncuredMnth, sBank, sVatGroup, sCardCode As String
        Dim iCount, iErrCode As Integer

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sClinicCode = oDv(0)(1).ToString.Trim
            sPayMethod = oDv(0)(24).ToString.Trim
            sArCode = oDv(0)(35).ToString.Trim
            sIncuredMnth = oDv(0)(34).ToString.Trim

            'sCardCode = p_oCompDef.sHMDCARInvPayCardCode

            sSql = "SELECT ""U_ar_code"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_PAYMETHOD"" " & _
                  " WHERE UPPER(""U_PayMethod"") = '" & sPayMethod.ToUpper & "' AND ""U_ClinicCode"" = '" & sClinicCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sCardCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            dtCardCode.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
            If dtCardCode.DefaultView.Count = 0 Then
                sErrDesc = "CardCode not found in SAP/CardCode :: " & sCardCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else
                sCardCode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                sVatGroup = dtCardCode.DefaultView.Item(0)(1).ToString().Trim()
            End If

            'sSql = "SELECT ""U_DefBank"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OPRC"" WHERE ""PrcCode"" = '" & sClinicCode & "'"
            sSql = "SELECT ""U_BankGL"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_PAYMETHOD"" " & _
                  " WHERE UPPER(""U_PayMethod"") = '" & sPayMethod.ToUpper & "' AND ""U_ClinicCode"" = '" & sClinicCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sBank = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            If sBank = "" Then
                sErrDesc = "Check the Bank in Cost Center table for the clinic code : " & sClinicCode
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_ITEMCODE"" " & _
                   " WHERE UPPER(""U_FileCode"") = 'HMDC' AND UPPER(""U_Field"") = 'PAY_CLIENT' AND UPPER(""U_DocType"") = 'A/R' AND UPPER(""U_CustType"") = 'CASH' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            If sItemCode = "" Then
                sErrDesc = "Check ItemCode in configuration table"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dtItemCode.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
            If dtItemCode.DefaultView.Count = 0 Then
                sErrDesc = "ItemCode not found in SAP/Item Code :: " & sItemCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim dTotal As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                Try
                    If Not (oDv(i)(15).ToString.Trim = String.Empty) Then
                        dAmount = CDbl(oDv(i)(15).ToString.Trim)
                    End If
                Catch ex As Exception
                    dAmount = 0.0
                End Try

                dTotal = dTotal + dAmount
            Next

            If dTotal > 0 Then
                Dim oARInvoice As SAPbobsCOM.Documents
                oARInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                oARInvoice.CardCode = sCardCode
                oARInvoice.DocDate = dDocDate
                iCount = 1

                If iCount > 1 Then
                    oARInvoice.Lines.Add()
                End If

                oARInvoice.Lines.ItemCode = sItemCode
                oARInvoice.Lines.Quantity = 1
                oARInvoice.Lines.Price = dTotal
                If Not (sVatGroup = String.Empty) Then
                    oARInvoice.Lines.VatGroup = sVatGroup
                End If
                If Not (sClinicCode = String.Empty) Then
                    oARInvoice.Lines.CostingCode2 = sClinicCode
                    oARInvoice.Lines.COGSCostingCode2 = sClinicCode
                End If
                iCount = iCount + 1

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                If oARInvoice.Add() <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo, iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)

                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim sQuery As String

                    Dim oPayments As SAPbobsCOM.Payments
                    oPayments = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                    oPayments.CardCode = sCardCode
                    oPayments.DocDate = dDocDate
                    oPayments.Invoices.DocEntry = iDocEntry

                    oPayments.TransferAccount = sBank
                    oPayments.TransferDate = Date.Now.ToString()
                    oPayments.TransferReference = sPayMethod
                    oPayments.TransferSum = dTotal

                    If oPayments.Add() <> 0 Then
                        sErrDesc = "ERROR DURING PAYMENT AFTER INVOICE / " & oCompany.GetLastErrorDescription
                        Throw New ArgumentException(sErrDesc)
                    Else
                        'sSql = "UPDATE " & p_oCompDef.sHMDCSAPDbName & ".""OINV"" SET ""IsICT"" = 'Y' WHERE ""DocEntry"" = '" & iDocEntry & "'"
                        'objRS.DoQuery(sSql)

                        sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OINV"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                        objRS.DoQuery(sSql)
                        If objRS.RecordCount > 0 Then
                            iDocNo = objRS.Fields.Item("DocNum").Value
                        End If
                        Console.WriteLine("Document Created successfully :: " & iDocNo)

                        Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1")
                        For k As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                                Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                                sQuery = "UPDATE " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AR_DETAILS"" SET ""U_Inv_DocNo"" = '" & iDocNo & "',""U_Inv_DocEntry"" = '" & iDocEntry & "'" & _
                                 " WHERE ""U_cln_code"" = '" & sClinicCode & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' " & _
                                 " AND ""U_paymethod"" = '" & sPayMethod & "' AND ""U_invoice"" = '" & sInvoice & "' AND IFNULL(""U_Inv_DocEntry"",'') = '' "

                                objRS.DoQuery(sQuery)
                            End If
                        Next


                    End If

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateARInvoicePayment_Contract = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateARInvoicePayment_Contract = RTN_ERROR
        End Try
    End Function

    Private Function InsertIntoAPDetails(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "InsertIntoApDetails"
        Dim sSql As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sItemCode As String = String.Empty
        Dim sApCode As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_ITEMCODE"" WHERE UPPER(""U_FileCode"") = 'HMDC' AND UPPER(""U_Field"") = 'PAY_CLIENT' " & _
                   " AND UPPER(""U_DocType"") = 'A/P' AND UPPER(""U_CustType"") = 'CONTRACT'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            sSql = "SELECT ""U_CardCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_CONTRACT_OWNER"" WHERE UPPER(""U_Type"") = 'A/P'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sApCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            Dim sInvoiceDt As String = oDv(0)(37).ToString.Trim
            Dim iIndex As Integer = sInvoiceDt.IndexOf(" ")
            Dim sInvoiceDate_Trimed As String = sInvoiceDt.Substring(0, iIndex)
            Dim dt As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sInvoiceDate_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim dAmount As Double = 0
            For i As Integer = 0 To oDv.Count - 1
                sSql = String.Empty

                Try
                    If Not (oDv(i)(15).ToString.Trim = String.Empty) Then
                        dAmount = CDbl(oDv(i)(15).ToString.Trim)
                    End If
                Catch ex As Exception
                    dAmount = 0.0
                End Try

            Next

            sSql = "INSERT INTO " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AP_DETAILS""(""Code"",""Name"",""U_company_code"",""U_incurred_month"",""U_invoice_date"",""U_invoice"", " & _
                       " ""U_amount"",""U_cln_code"",""U_subcode"",""U_ItemCode"",""U_invoice_type"",""U_ap_code"",""U_CostCenter"") " & _
                       " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AP_DETAILS""), " & _
                       " (SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AP_DETAILS""), " & _
                       " '" & oDv(0)(25).ToString & "','" & oDv(0)(34).ToString & "', '" & dt.ToString("yyyy-MM-dd") & "','" & oDv(0)(0).ToString & "','" & dAmount & "','" & oDv(0)(1).ToString & "', " & _
                       " '" & oDv(0)(2).ToString & "','" & sItemCode & "','CONTRACT','" & sApCode & "','" & oDv(0)(1).ToString & "') "

            oRecordSet.DoQuery(sSql)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoAPDetails = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            InsertIntoAPDetails = RTN_ERROR
        End Try
    End Function

    Private Function CreateAPInvoice_Contract_HMMPD(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateAPInvoice_Contract_HMMPD"
        Dim sItemCode As String = String.Empty
        Dim sSql As String = String.Empty
        Dim dAmount As Double = 0.0
        Dim sClinicCode, sApCode, sIncuredMnth, sCardcode, sVatGroup As String
        Dim iCount, iErrCode As Integer

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sClinicCode = oDv(0)(1).ToString.Trim
            sApCode = oDv(0)(36).ToString.Trim
            sIncuredMnth = oDv(0)(34).ToString.Trim

            sSql = "SELECT ""U_CardCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_CONTRACT_OWNER"" WHERE UPPER(""U_Type"") = 'A/P' AND UPPER(""U_CustomerType"") = 'CONTRACT' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sCardcode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            dtCardCode.DefaultView.RowFilter = "CardCode = '" & sCardcode & "'"
            If dtCardCode.DefaultView.Count = 0 Then
                sErrDesc = "CardCode not found in SAP/CardCode :: " & sCardcode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else
                sCardcode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                sVatGroup = dtCardCode.DefaultView.Item(0)(1).ToString().Trim()
            End If

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_ITEMCODE"" " & _
                  " WHERE UPPER(""U_FileCode"") = 'HMDC' AND UPPER(""U_Field"") = 'PAY_CLIENT' AND UPPER(""U_DocType"") = 'A/P' AND UPPER(""U_CustType"") = 'CONTRACT' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            If sItemCode = "" Then
                sErrDesc = "Check ItemCode in configuration table"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dtItemCode.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
            If dtItemCode.DefaultView.Count = 0 Then
                sErrDesc = "ItemCode not found in SAP/Item Code :: " & sItemCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            Dim dTotal As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                Try
                    If Not (oDv(i)(15).ToString.Trim = String.Empty) Then
                        dAmount = CDbl(oDv(i)(15).ToString.Trim)
                    End If
                Catch ex As Exception
                    dAmount = 0.0
                End Try

                dTotal = dTotal + dAmount
            Next

            Dim dPercent As Double = 0.0
            sSql = "SELECT ""U_Percentage"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_PERCENTAGE"" WHERE UPPER(""U_Type"") = 'A/R'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dPercent = getAmt(sSql, p_oCompDef.sHMMPDSAPDbName, p_oCompDef.sHMMPDSAPUserName, p_oCompDef.sHMMPDSAPPassword)

            dTotal = dTotal * (dPercent / 100)

            If dTotal > 0 Then
                Dim oApInvoice As SAPbobsCOM.Documents
                oApInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

                oApInvoice.CardCode = sCardcode
                oApInvoice.DocDate = dDocDate
                iCount = 1

                If iCount > 1 Then
                    oApInvoice.Lines.Add()
                End If

                oApInvoice.Lines.ItemCode = sItemCode
                oApInvoice.Lines.Quantity = 1
                oApInvoice.Lines.Price = dTotal
                If Not (sVatGroup = String.Empty) Then
                    oApInvoice.Lines.VatGroup = sVatGroup
                End If
                If Not (sClinicCode = String.Empty) Then
                    oApInvoice.Lines.CostingCode2 = sClinicCode
                    oApInvoice.Lines.COGSCostingCode2 = sClinicCode
                End If
                iCount = iCount + 1

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                If oApInvoice.Add() <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo, iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oApInvoice)

                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim sQuery As String

                    sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OPCH"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSql)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AP_DETAILS"" SET ""U_AP_Inv_DocNo"" = '" & iDocNo & "',""U_AP_Inv_DocEntry"" = '" & iDocEntry & "'" & _
                                     " WHERE ""U_cln_code"" = '" & sClinicCode & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' AND ""U_invoice"" = '" & sInvoice & "'"

                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateAPInvoice_Contract_HMMPD = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateAPInvoice_Contract_HMMPD = RTN_ERROR
        End Try
    End Function

    Private Function ProcessNonContractDatas(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessNonContractDatas"
        Dim sSql As String = String.Empty
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'F1 Invoice F2 Clinic Code, F25 Payment Method
            Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1", "ArCode", "F2", "IncuredMonth")
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    oDv.RowFilter = "F1='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and ArCode = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                    " and F2='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'  and IncuredMonth='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "' "
                    If oDv.Count > 0 Then
                        Console.WriteLine("Inserting data into Cash table for " & oDtGroup.Rows(i).Item(1).ToString.Trim())
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoCashSales()", sFuncName)
                        Dim oCashDt As DataTable
                        oCashDt = oDv.ToTable
                        Dim oCashDv As DataView = New DataView(oCashDt)
                        If InsertIntoCashSales(oCashDv, p_oCompany, "LESS_DIS_PAY_CLIENT", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Console.WriteLine("Inserting table into cash table is successful")
                    End If
                End If
            Next

            oDv.RowFilter = Nothing

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas to create A/R invoice + Payment", sFuncName)
            oDtGroup = oDv.Table.DefaultView.ToTable(True, "F2", "ArCode", "IncuredMonth")
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    oDv.RowFilter = "F2='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and ArCode = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                    " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'  "
                    If oDv.Count > 0 Then
                        Console.WriteLine("Creating Invoice for cash table")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoicePayment_NonContract()", sFuncName)
                        Dim oCashTab_InvoiceDt As DataTable
                        oCashTab_InvoiceDt = oDv.ToTable
                        Dim oCashTab_InvoiceDv As DataView = New DataView(oCashTab_InvoiceDt)
                        If CreateARInvoicePayment_NonContract(oCashTab_InvoiceDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Console.WriteLine("Invoice Creation for Cash table successful")
                    End If
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessNonContractDatas = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessNonContractDatas = RTN_ERROR
        End Try
    End Function

    Private Function CreateARInvoicePayment_NonContract(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateARInvoicePayment_NonContract"
        Dim sItemCode As String = String.Empty
        Dim sSql As String = String.Empty
        Dim dAmount As Double = 0.0
        Dim sClinicCode, sPayMethod, sArCode, sIncuredMnth, sBank, sVatGroup As String
        Dim iCount, iErrCode As Integer

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sClinicCode = oDv(0)(1).ToString.Trim
            sPayMethod = oDv(0)(24).ToString.Trim
            sArCode = oDv(0)(35).ToString.Trim
            sIncuredMnth = oDv(0)(34).ToString.Trim

            dtCardCode.DefaultView.RowFilter = "CardCode = '" & sArCode & "'"
            If dtCardCode.DefaultView.Count = 0 Then
                sErrDesc = "CardCode not found in SAP/CardCode :: " & sArCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else
                sArCode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                sVatGroup = dtCardCode.DefaultView.Item(0)(1).ToString().Trim()
            End If

            sSql = "SELECT ""U_DefBank"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OPRC"" WHERE ""PrcCode"" = '" & sClinicCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sBank = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            If sBank = "" Then
                sErrDesc = "Check the Bank in cost center table for the clinic code : " & sClinicCode
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Default bank in Cost center table is null / Check Cost center " & sClinicCode, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            sSql = "SELECT ""U_Type"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OCRD"" WHERE ""CardCode"" = '" & sArCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            If sItemCode = "" Then
                sErrDesc = "Check ItemCode is mandatory/Check U_Type column in Business partner master/BP Code : " & sArCode
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dtItemCode.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
            If dtItemCode.DefaultView.Count = 0 Then
                sErrDesc = "ItemCode not found in SAP/Item Code :: " & sItemCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim dTotal As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                Try
                    If Not (oDv(i)(15).ToString.Trim = String.Empty) Then
                        dAmount = CDbl(oDv(i)(15).ToString.Trim)
                    End If
                Catch ex As Exception
                    dAmount = 0.0
                End Try

                dTotal = dTotal + dAmount
            Next

            If dTotal > 0 Then
                Dim oARInvoice As SAPbobsCOM.Documents
                oARInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                oARInvoice.CardCode = sArCode
                oARInvoice.DocDate = dDocDate
                iCount = 1

                If iCount > 1 Then
                    oARInvoice.Lines.Add()
                End If

                oARInvoice.Lines.ItemCode = sItemCode
                oARInvoice.Lines.Quantity = 1
                oARInvoice.Lines.Price = dTotal
                If Not (sVatGroup = String.Empty) Then
                    oARInvoice.Lines.VatGroup = sVatGroup
                End If
                If Not (sClinicCode = String.Empty) Then
                    oARInvoice.Lines.CostingCode2 = sClinicCode
                    oARInvoice.Lines.COGSCostingCode2 = sClinicCode
                End If
                iCount = iCount + 1

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                If oARInvoice.Add() <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo, iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)

                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim sQuery As String

                    Dim oPayments As SAPbobsCOM.Payments
                    oPayments = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

                    oPayments.CardCode = sArCode
                    oPayments.DocDate = dDocDate
                    oPayments.Invoices.DocEntry = iDocEntry

                    oPayments.TransferAccount = sBank
                    oPayments.TransferDate = Date.Now.ToString()
                    oPayments.TransferReference = sPayMethod
                    oPayments.TransferSum = dTotal

                    If oPayments.Add() <> 0 Then
                        sErrDesc = "ERROR DURING PAYMENT AFTER INVOICE / " & oCompany.GetLastErrorDescription
                        Throw New ArgumentException(sErrDesc)
                    Else
                        'sSql = "UPDATE " & p_oCompDef.sHMDCSAPDbName & ".""OINV"" SET ""IsICT"" = 'Y' WHERE ""DocEntry"" = '" & iDocEntry & "'"
                        'objRS.DoQuery(sSql)

                        sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OINV"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                        objRS.DoQuery(sSql)
                        If objRS.RecordCount > 0 Then
                            iDocNo = objRS.Fields.Item("DocNum").Value
                        End If
                        Console.WriteLine("Document Created successfully :: " & iDocNo)

                        Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1")
                        For k As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                                Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                                sQuery = "UPDATE " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AR_DETAILS"" SET ""U_Inv_DocNo"" = '" & iDocNo & "',""U_Inv_DocEntry"" = '" & iDocEntry & "'" & _
                                         " WHERE ""U_ar_code"" = '" & sArCode & "' AND ""U_cln_code"" = '" & sClinicCode & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' " & _
                                         " AND ""U_paymethod"" = '" & sPayMethod & "' AND ""U_invoice"" = '" & sInvoice & "'"

                                objRS.DoQuery(sQuery)
                            End If
                        Next

                    End If

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateARInvoicePayment_NonContract = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateARInvoicePayment_NonContract = RTN_ERROR
        End Try
    End Function

    Private Function ProcessNonContract_NonPay(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessNonContract_NonPay"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1", "ArCode", "F2", "IncuredMonth")
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    oDv.RowFilter = "F1='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and ArCode = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                    " and F2='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'  and IncuredMonth='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "' "
                    If oDv.Count > 0 Then
                        Console.WriteLine("Inserting data into Cash table for " & oDtGroup.Rows(i).Item(1).ToString.Trim())
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoCashSales()", sFuncName)
                        Dim oCashDt As DataTable
                        oCashDt = oDv.ToTable
                        Dim oCashDv As DataView = New DataView(oCashDt)
                        If InsertIntoCashSales(oCashDv, p_oCompany, "LESS_DIS_PAY_CLIENT", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Console.WriteLine("Inserting table into cash table is successful")
                    End If
                End If
            Next

            oDv.RowFilter = Nothing

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas to create A/R invoice", sFuncName)
            oDtGroup = oDv.Table.DefaultView.ToTable(True, "F2", "ArCode", "IncuredMonth")
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    oDv.RowFilter = "F2='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and ArCode = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                    " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'  "
                    If oDv.Count > 0 Then
                        Console.WriteLine("Creating Invoice for cash table")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoice_Contract()", sFuncName)
                        Dim oInvoiceDt As DataTable
                        oInvoiceDt = oDv.ToTable
                        Dim oInvoiceDv As DataView = New DataView(oInvoiceDt)
                        If CreateArInvoice_NonContract(oInvoiceDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Console.WriteLine("Invoice Creation for Cash table successful")
                    End If
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessNonContract_NonPay = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessNonContract_NonPay = RTN_ERROR
        End Try
    End Function

    Private Function CreateArInvoice(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateArInvoice"
        Dim sItemCode As String = String.Empty
        Dim sSql As String = String.Empty
        Dim dAmount As Double = 0.0
        Dim sClinicCode, sPayMethod, sArCode, sIncuredMnth, sVatGroup, sCardCode As String
        Dim iCount, iErrCode As Integer

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sClinicCode = oDv(0)(1).ToString.Trim
            sPayMethod = oDv(0)(24).ToString.Trim
            sArCode = oDv(0)(35).ToString.Trim
            sIncuredMnth = oDv(0)(34).ToString.Trim

            sCardCode = p_oCompDef.sHMDCARInvPayCardCode

            dtCardCode.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
            If dtCardCode.DefaultView.Count = 0 Then
                sErrDesc = "CardCode not found in SAP/CardCode :: " & sCardCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else
                sCardCode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                sVatGroup = dtCardCode.DefaultView.Item(0)(1).ToString().Trim()
            End If

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            sSql = "SELECT ""U_Type"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OCRD"" WHERE ""CardCode"" = '" & sCardCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            If sItemCode = "" Then
                sErrDesc = "Check ItemCode is mandatory/Check U_Type column in Business partner master/BP Code : " & sCardCode
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dtItemCode.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
            If dtItemCode.DefaultView.Count = 0 Then
                sErrDesc = "ItemCode not found in SAP/Item Code :: " & sItemCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim dTotal As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                Try
                    If Not (oDv(i)(15).ToString.Trim = String.Empty) Then
                        dAmount = CDbl(oDv(i)(15).ToString.Trim)
                    End If
                Catch ex As Exception
                    dAmount = 0.0
                End Try

                dTotal = dTotal + dAmount
            Next

            If dTotal > 0 Then
                Dim oARInvoice As SAPbobsCOM.Documents
                oARInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                oARInvoice.CardCode = sCardCode
                oARInvoice.DocDate = dDocDate
                iCount = 1

                If iCount > 1 Then
                    oARInvoice.Lines.Add()
                End If

                oARInvoice.Lines.ItemCode = sItemCode
                oARInvoice.Lines.Quantity = 1
                oARInvoice.Lines.Price = dTotal
                If Not (sVatGroup = String.Empty) Then
                    oARInvoice.Lines.VatGroup = sVatGroup
                End If
                If Not (sClinicCode = String.Empty) Then
                    oARInvoice.Lines.CostingCode2 = sClinicCode
                    oARInvoice.Lines.COGSCostingCode2 = sClinicCode
                End If
                iCount = iCount + 1

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                If oARInvoice.Add() <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo, iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)

                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim sQuery As String

                    sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OINV"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSql)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AR_DETAILS"" SET ""U_Inv_DocNo"" = '" & iDocNo & "',""U_Inv_DocEntry"" = '" & iDocEntry & "'" & _
                                     " WHERE ""U_cln_code"" = '" & sClinicCode & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' " & _
                                     " AND ""U_paymethod"" = '" & sPayMethod & "' AND ""U_invoice"" = '" & sInvoice & "' AND IFNULL(""U_Inv_DocEntry"",'') = '' "

                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateArInvoice = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateArInvoice = RTN_ERROR
        End Try
    End Function

    Private Function CreateArInvoice_NonContract(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateArInvoice_NonContract"
        Dim sItemCode As String = String.Empty
        Dim sSql As String = String.Empty
        Dim dAmount As Double = 0.0
        Dim sClinicCode, sPayMethod, sArCode, sIncuredMnth, sVatGroup, sCardCode As String
        Dim iCount, iErrCode As Integer

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sClinicCode = oDv(0)(1).ToString.Trim
            sPayMethod = oDv(0)(24).ToString.Trim
            sArCode = oDv(0)(35).ToString.Trim
            sIncuredMnth = oDv(0)(34).ToString.Trim

            sCardCode = sArCode

            dtCardCode.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
            If dtCardCode.DefaultView.Count = 0 Then
                sErrDesc = "CardCode not found in SAP/CardCode :: " & sCardCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else
                sCardCode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                sVatGroup = dtCardCode.DefaultView.Item(0)(1).ToString().Trim()
            End If

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            sSql = "SELECT ""U_Type"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OCRD"" WHERE ""CardCode"" = '" & sCardCode & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            If sItemCode = "" Then
                sErrDesc = "Check ItemCode is mandatory/Check U_Type column in Business partner master/BP Code : " & sCardCode
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dtItemCode.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
            If dtItemCode.DefaultView.Count = 0 Then
                sErrDesc = "ItemCode not found in SAP/Item Code :: " & sItemCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim dTotal As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                Try
                    If Not (oDv(i)(15).ToString.Trim = String.Empty) Then
                        dAmount = CDbl(oDv(i)(15).ToString.Trim)
                    End If
                Catch ex As Exception
                    dAmount = 0.0
                End Try

                dTotal = dTotal + dAmount
            Next

            If dTotal > 0 Then
                Dim oARInvoice As SAPbobsCOM.Documents
                oARInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                oARInvoice.CardCode = sCardCode
                oARInvoice.DocDate = dDocDate
                iCount = 1

                If iCount > 1 Then
                    oARInvoice.Lines.Add()
                End If

                oARInvoice.Lines.ItemCode = sItemCode
                oARInvoice.Lines.Quantity = 1
                oARInvoice.Lines.Price = dTotal
                If Not (sVatGroup = String.Empty) Then
                    oARInvoice.Lines.VatGroup = sVatGroup
                End If
                If Not (sClinicCode = String.Empty) Then
                    oARInvoice.Lines.CostingCode2 = sClinicCode
                    oARInvoice.Lines.COGSCostingCode2 = sClinicCode
                End If
                iCount = iCount + 1

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                If oARInvoice.Add() <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo, iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)

                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim sQuery As String

                    sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OINV"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSql)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sHMDCSAPDbName & ".""@AE_AR_DETAILS"" SET ""U_Inv_DocNo"" = '" & iDocNo & "',""U_Inv_DocEntry"" = '" & iDocEntry & "'" & _
                                     " WHERE ""U_cln_code"" = '" & sClinicCode & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' " & _
                                     " AND ""U_paymethod"" = '" & sPayMethod & "' AND ""U_invoice"" = '" & sInvoice & "' AND IFNULL(""U_Inv_DocEntry"",'') = '' "

                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateArInvoice_NonContract = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateArInvoice_NonContract = RTN_ERROR
        End Try
    End Function

    Private Function CreateARInvoice_HMMPD(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateARInvoice_HMMPD"
        Dim sItemCode As String = String.Empty
        Dim sSql As String = String.Empty
        Dim dAmount As Double = 0.0
        Dim sClinicCode, sArCode, sIncuredMnth, sCardCode, sVatGroup As String
        Dim iCount, iErrCode As Integer

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sClinicCode = oDv(0)(1).ToString.Trim
            sArCode = oDv(0)(35).ToString.Trim
            sIncuredMnth = oDv(0)(34).ToString.Trim

            sSql = "SELECT ""U_CardCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_CONTRACT_OWNER"" WHERE UPPER(""U_Type"") = 'A/R' AND UPPER(""U_CustomerType"") = 'CONTRACT'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sCardCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            dtCardCode.DefaultView.RowFilter = "CardCode = '" & sCardCode & "'"
            If dtCardCode.DefaultView.Count = 0 Then
                sErrDesc = "CardCode not found in SAP/CardCode :: " & sCardCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else
                sCardCode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                sVatGroup = dtCardCode.DefaultView.Item(0)(1).ToString().Trim()
            End If

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""@AE_ITEMCODE"" " & _
                  " WHERE UPPER(""U_FileCode"") = 'HMDC' AND UPPER(""U_Field"") = 'PAY_COMP' AND UPPER(""U_DocType"") = 'A/R' AND UPPER(""U_CustType"") = 'CONTRACT' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMDCSAPDbName)

            If sItemCode = "" Then
                sErrDesc = "Check ItemCode in configuration table"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dtItemCode.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
            If dtItemCode.DefaultView.Count = 0 Then
                sErrDesc = "ItemCode not found in SAP/Item Code :: " & sItemCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            Dim dTotal As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                Try
                    If Not (oDv(i)(13).ToString.Trim = String.Empty) Then
                        dAmount = CDbl(oDv(i)(13).ToString.Trim)
                    End If
                Catch ex As Exception
                    dAmount = 0.0
                End Try

                dTotal = dTotal + dAmount
            Next

            Dim dPercent As Double
            sSql = "SELECT ""U_Percentage"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_PERCENTAGE"" WHERE UPPER(""U_Type"") = 'A/P'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dPercent = getAmt(sSql, p_oCompDef.sHMMPDSAPDbName, p_oCompDef.sHMMPDSAPUserName, p_oCompDef.sHMMPDSAPPassword)

            dTotal = dTotal * (dPercent / 100)

            If dTotal > 0 Then
                Dim oARInvoice As SAPbobsCOM.Documents
                oARInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                oARInvoice.CardCode = sCardCode
                oARInvoice.DocDate = dDocDate
                iCount = 1

                If iCount > 1 Then
                    oARInvoice.Lines.Add()
                End If

                oARInvoice.Lines.ItemCode = sItemCode
                oARInvoice.Lines.Quantity = 1
                oARInvoice.Lines.Price = dTotal
                If Not (sVatGroup = String.Empty) Then
                    oARInvoice.Lines.VatGroup = sVatGroup
                End If
                If Not (sClinicCode = String.Empty) Then
                    oARInvoice.Lines.CostingCode2 = sClinicCode
                    oARInvoice.Lines.COGSCostingCode2 = sClinicCode
                End If
                iCount = iCount + 1

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                If oARInvoice.Add() <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocNo, iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)

                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim sQuery As String

                    sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sHMDCSAPDbName & ".""OINV"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSql)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sHMDCSAPDbName & ".""@AE_RD001_AR"" SET ""U_HMMPD_ARInvNo"" = '" & iDocNo & "', ""U_HMMPD_ARInvEntry"" = '" & iDocEntry & "'" & _
                                     " WHERE ""U_cln_code"" = '" & sClinicCode & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' " & _
                                     " AND ""U_invoice"" = '" & sInvoice & "' "

                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateARInvoice_HMMPD = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateARInvoice_HMMPD = RTN_ERROR
        End Try
    End Function

End Module
