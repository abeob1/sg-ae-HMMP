Imports System.Globalization

Module modSalesOrder

    Private dtSoAcrlInvList As DataTable
    Private dtCostAcrlInvList As DataTable
    Private dtInsurerList As DataTable
    Private dtMBMSList As DataTable

    Public Function ProcessSalesOrders(ByVal file As System.IO.FileInfo, ByVal oDv As DataView, ByRef sErrDesc As String) As Long
        'Function created to take backup before changing to code to check clinic code exists in TPA table
        Dim sFuncName As String = "ProcessSalesOrders"
        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE IFNULL(""U_JrnlEntry"",'') <> '' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtSoAcrlInvList = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

            'sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_JrnlEntry"",'') <> '' "
            sSQL = "SELECT DISTINCT ""U_invoice"" FROM( " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_JrnlEntry"",'') <> '' UNION ALL " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_Glar_NC_DocEntry"",'') <> '' UNION ALL " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_Glap_NC_DocEntry"",'') <> '' )T1 "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtCostAcrlInvList = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

            Dim odtDatatable As DataTable
            odtDatatable = oDv.ToTable

            odtDatatable.Columns.Add("CostCenter", GetType(String))
            odtDatatable.Columns.Add("Insurer", GetType(String))
            odtDatatable.Columns.Add("IncuredMonth", GetType(Date))
            odtDatatable.Columns.Add("Type", GetType(String))

            For intRow As Integer = 0 To odtDatatable.Rows.Count - 1
                If Not (odtDatatable.Rows(intRow).Item(1).ToString.Trim() = String.Empty Or odtDatatable.Rows(intRow).Item(1).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
                    Console.WriteLine("Processing excel line " & intRow & " to get MBMS and Insurer from config table")

                    Dim sCompCode As String = odtDatatable.Rows(intRow).Item(1).ToString
                    Dim sCompName As String = odtDatatable.Rows(intRow).Item(0).ToString
                    sCompName = sCompName.Replace("'", " ")
                    Dim sSchemeCode As String = odtDatatable.Rows(intRow).Item(3).ToString
                    Dim sClinicCode As String = odtDatatable.Rows(intRow).Item(4).ToString
                    Dim sRemarks As String = odtDatatable.Rows(intRow).Item(29).ToString
                    sRemarks = sRemarks.Replace("'", " ")
                    Dim sDiagDesc As String = odtDatatable.Rows(intRow).Item(23).ToString
                    sDiagDesc = sDiagDesc.Replace("'", " ")

                    If sCompCode = "" Then
                        sErrDesc = "Company Code should not be empty / Check Line " & intRow
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine(sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Dim sType As String
                    Dim sArCode As String = "C" & sCompCode
                    sSQL = "SELECT ""U_Type"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" WHERE ""CardCode"" = '" & sArCode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    sType = GetStringValue(sSQL, p_oCompDef.sSAPDBName)

                    If sType = "" Then
                        sType = p_oCompDef.sType
                    End If

                    Dim sInvoice As String = odtDatatable.Rows(intRow).Item(17).ToString.Trim
                    dtSoAcrlInvList.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtSoAcrlInvList.DefaultView.Count > 0 Then
                        sErrDesc = "Invoice has been created previously for invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    dtCostAcrlInvList.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtCostAcrlInvList.DefaultView.Count > 0 Then
                        sErrDesc = "Invoice has been created previously for invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Dim iIndex As Integer = odtDatatable.Rows(intRow).Item(16).ToString.IndexOf(" ")
                    Dim sDate As String = odtDatatable.Rows(intRow).Item(16).ToString.Substring(0, iIndex)
                    Dim dt As Date
                    Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                    Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

                    Dim dIncurMnth As Date = CDate(dt.Date.AddDays(-(dt.Day - 1)).AddMonths(1).AddDays(-1).ToString())
                    Dim sCostCenter As String = GetCostCenter(sCompCode, dt, sSchemeCode, p_oCompDef.sSAPDBName)
                    Dim sInsurer As String = GetInsurer(sCompCode, dt, sSchemeCode, p_oCompDef.sSAPDBName)

                    If sCostCenter = "" Then
                        sErrDesc = "MBMS column cannot be null / Check Cost Center for respective company code in config table/Check line " & intRow
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine(sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    End If
                    If sInsurer = "" Then
                        sErrDesc = "Insurer column cannot be null / Check Insurer for the respective company code in config table /Check line " & intRow
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine(sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    odtDatatable.Rows(intRow)("F1") = sCompName
                    odtDatatable.Rows(intRow)("F5") = sClinicCode.ToUpper()
                    odtDatatable.Rows(intRow)("F24") = sDiagDesc
                    odtDatatable.Rows(intRow)("F30") = sRemarks
                    odtDatatable.Rows(intRow)("CostCenter") = sCostCenter
                    odtDatatable.Rows(intRow)("Insurer") = sInsurer
                    odtDatatable.Rows(intRow)("IncuredMonth") = dIncurMnth
                    odtDatatable.Rows(intRow)("Type") = sType.ToUpper
                End If
            Next

            Dim oDvFinalView As DataView
            oDvFinalView = New DataView(odtDatatable)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToCompany(p_oCompany, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then
                Console.WriteLine("Company connection Successful")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction", sFuncName)

                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoSOTable()", sFuncName)

                Console.WriteLine("Inserting datas in SO Table")
                If InsertIntoSOTable(oDvFinalView, file.Name, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Console.WriteLine("Data insert into SO Table Successful")

                'F4 - Clinic Code
                Dim oDtGroup As DataTable = oDvFinalView.Table.DefaultView.ToTable(True, "F5")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                        Dim sCliniCode As String = oDtGroup.Rows(i).Item(0).ToString.Trim()

                        sSQL = "SELECT COUNT(""U_cln_code"") AS ""MNO"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_TPA_APCODE"" WHERE UPPER(""U_cln_code"") = '" & sCliniCode.ToUpper() & "'"
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                        Dim iCount As Integer = GetCode(sSQL, p_oCompDef.sSAPDBName)

                        oDvFinalView.RowFilter = "F5 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' "
                        If oDvFinalView.Count > 0 Then
                            If iCount > 0 Then
                                Dim odtTpaList As DataTable
                                odtTpaList = oDvFinalView.ToTable
                                Dim oDvTpaList As DataView = New DataView(odtTpaList)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessDatas_TPAListings()", sFuncName)
                                If ProcessDatas_TPAListings(oDvTpaList, p_oCompany, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            Else
                                'SIMILAR TO CLINIC CODE NOT LIKE OUT
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessDatas_NonTPAListings()", sFuncName)
                                Dim odtNonTpaList As DataTable
                                odtNonTpaList = oDvFinalView.ToTable
                                Dim oDvNonTpaList As DataView = New DataView(odtNonTpaList)
                                If ProcessDatas_NonTPAListings(oDvNonTpaList, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    End If
                Next

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
            ProcessSalesOrders = RTN_SUCCESS

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
            ProcessSalesOrders = RTN_ERROR

        End Try
    End Function

    Private Function InsertIntoSOTable(ByVal oDV As DataView, ByVal sFileName As String, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "InsertIntoSOTable"
        'Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName

        'Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        'Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()
        Dim sSql As String = String.Empty

        Try
            'Con.ConnectionString = sConstr
            'Con.Open()

            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'LINE BY LINE
            For i As Integer = 1 To oDV.Count - 1
                If Not (oDV(i)(1).ToString.Trim = String.Empty) Then
                    Console.WriteLine("Inserting Line Num : " & i)
                    sSql = String.Empty

                    Dim sCompCode As String = oDV(i)(1).ToString.Trim
                    Dim sApCode As String = "V" & oDV(i)(4).ToString.Trim

                    sSql = "INSERT INTO " & p_oCompDef.sSAPDBName & ".""@AE_MS007_SO"" (""Code"",""Name"",""U_company"",""U_company_code"",""U_C"",""U_scheme_code"",""U_cln_code"",""U_m_id_type"",""U_m_id"",""U_m_lastname"",""U_m_given_name"",""U_m_christian""," & _
                            " ""U_relation"",""U_id_type"",""U_id"",""U_lastname"",""U_given_name"",""U_christian"",""U_txn_date"",""U_invoice"",""U_treatment"",""U_charge"",""U_pay_comp"",""U_pay_client"",""U_diag"",""U_diag_desc"", " & _
                            " ""U_refer_from_name"",""U_policy_num"",""U_cert_num"",""U_treat_code"",""U_remark_fg"",""U_remark1"",""U_paiddate"",""U_status"",""U_status_code"",""U_cust_no"",""U_scheme_remark"",""U_dept1"",""U_dept2""," & _
                            " ""U_dept3"",""U_ds1"",""U_ds2"",""U_ds3"",""U_in_time"",""U_insco"",""U_sl_fr"",""U_sl_to"",""U_CompTotRecCnt"",""U_CompTotBillAmt"",""U_scheme_desc"",""U_CostCenter"",""U_Insurer"",""U_Incurred_month"",""U_ar_code"",""U_ap_code"",""U_Type"",""U_FileName"")" & _
                            " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS007_SO""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS007_SO"")," & _
                            " '" & oDV(i)(0).ToString.Trim & "','" & oDV(i)(1).ToString.Trim & "','" & oDV(i)(2).ToString.Trim & "','" & oDV(i)(3).ToString.Trim & "'," & _
                            " '" & oDV(i)(4).ToString.Trim & "','" & oDV(i)(5).ToString.Trim & "','" & oDV(i)(6).ToString.Trim & "','" & oDV(i)(7).ToString.Trim & "'," & _
                            " '" & oDV(i)(8).ToString.Trim & "','" & oDV(i)(9).ToString.Trim & "','" & oDV(i)(10).ToString.Trim & "','" & oDV(i)(11).ToString.Trim & "'," & _
                            " '" & oDV(i)(12).ToString.Trim & "','" & oDV(i)(13).ToString.Trim & "','" & oDV(i)(14).ToString.Trim & "','" & oDV(i)(15).ToString.Trim & "'," & _
                            " '" & oDV(i)(16).ToString.Trim & "','" & oDV(i)(17).ToString.Trim & "','" & oDV(i)(18).ToString.Trim & "','" & oDV(i)(19).ToString.Trim & "'," & _
                            " '" & oDV(i)(20).ToString.Trim & "','" & oDV(i)(21).ToString.Trim & "','" & oDV(i)(22).ToString.Trim & "','" & oDV(i)(23).ToString.Trim & "'," & _
                            " '" & oDV(i)(24).ToString.Trim & "','" & oDV(i)(25).ToString.Trim & "','" & oDV(i)(26).ToString.Trim & "','" & oDV(i)(27).ToString.Trim & "'," & _
                            " '" & oDV(i)(28).ToString.Trim & "','" & oDV(i)(29).ToString.Trim & "','" & oDV(i)(30).ToString.Trim & "','" & oDV(i)(31).ToString.Trim & "'," & _
                            " '" & oDV(i)(32).ToString.Trim & "','" & oDV(i)(33).ToString.Trim & "','" & oDV(i)(34).ToString.Trim & "','" & oDV(i)(35).ToString.Trim & "'," & _
                            " '" & oDV(i)(36).ToString.Trim & "','" & oDV(i)(37).ToString.Trim & "','" & oDV(i)(38).ToString.Trim & "','" & oDV(i)(39).ToString.Trim & "'," & _
                            " '" & oDV(i)(40).ToString.Trim & "','" & oDV(i)(41).ToString.Trim & "','" & oDV(i)(42).ToString.Trim & "','" & oDV(i)(43).ToString.Trim & "'," & _
                            " '" & oDV(i)(44).ToString.Trim & "','" & oDV(i)(45).ToString.Trim & "','" & oDV(i)(46).ToString.Trim & "','" & oDV(i)(47).ToString.Trim & "'," & _
                            " '" & oDV(i)(48).ToString.Trim & "','" & oDV(i)(49).ToString.Trim & "','" & oDV(i)(50).ToString.Trim & "', " & _
                            " (SELECT TOP 1 ""CardCode"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" WHERE ""CardFName"" = '" & sCompCode & "'),'" & sApCode & "', '" & oDV(i)(51).ToString.Trim & "','" & sFileName & "') "

                    oRecSet.DoQuery(sSql)

                    'Dim oCmd As New Odbc.OdbcCommand
                    'oCmd.CommandType = CommandType.Text
                    'oCmd.CommandText = sSql
                    'oCmd.Connection = Con
                    'oCmd.CommandTimeout = 0
                    'oCmd.ExecuteNonQuery()
                End If
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoSOTable = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            InsertIntoSOTable = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)

            'Finally
            '    Con.Dispose()
        End Try

    End Function

    Private Function ProcessDatas_TPAListings(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessDatas_TPAListings"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oDv.RowFilter = "Type LIKE 'CAPITATION*'"
            Dim odt As New DataTable
            odt = oDv.ToTable
            Dim oNewDv As DataView = New DataView(odt)

            If oNewDv.Count > 0 Then
                'F5 - Clinic code F18 - Invoice
                Dim oDtGroup As DataTable = oNewDv.Table.DefaultView.ToTable(True, "F5", "F18", "CostCenter", "IncuredMonth")
                Console.WriteLine("Grouping datas for insert into Cost accrual table")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                        oNewDv.RowFilter = "F5 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F18 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                 " and CostCenter='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' and IncuredMonth='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "'"

                        If oNewDv.Count > 0 Then
                            Dim odtCapData As DataTable
                            odtCapData = oNewDv.ToTable
                            Dim oDvCapData As DataView = New DataView(odtCapData)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoCostAccrual()", sFuncName)
                            If InsertIntoCostAccrual(oDvCapData, p_oCompany, file.Name, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Next
                Console.WriteLine("Cost Accrual Data insert successful")

                oNewDv.RowFilter = Nothing

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for journal entry creation", sFuncName)
                'F5 Clinic Code, F3 Sub Code
                oDtGroup = oNewDv.Table.DefaultView.ToTable(True, "F5", "F3", "CostCenter", "IncuredMonth")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                        oNewDv.RowFilter = "F5 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F3 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                            " and CostCenter = '" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "' "

                        If oNewDv.Count > 0 Then
                            Dim odtCapData_JE As DataTable
                            odtCapData_JE = oNewDv.ToTable
                            Dim oDvCapData_JE As DataView = New DataView(odtCapData_JE)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateCostAccrualJE()", sFuncName)
                            If CreateCostAccrualJE(oDvCapData_JE, file, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Next
                Console.WriteLine("Journal Entry creation for cost accrual datas successful")
            End If

            oDv.RowFilter = Nothing

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Filtering Non capitation rows", sFuncName)

            oDv.RowFilter = "Type NOT LIKE 'CAPITATION*'"
            Dim odt_NonCap As New DataTable
            odt_NonCap = oDv.ToTable
            Dim oDv_NonCap As DataView = New DataView(odt_NonCap)

            If oDv_NonCap.Count > 0 Then
                Dim oDtGroup As DataTable = oDv_NonCap.Table.DefaultView.ToTable(True, "F5", "F18", "CostCenter", "IncuredMonth")
                Console.WriteLine("Grouping datas for insert into Cost accrual table Non Capitation")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                        oDv_NonCap.RowFilter = "F5 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F18 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                 " and CostCenter='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' and IncuredMonth='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "'"

                        If oDv_NonCap.Count > 0 Then
                            Dim odtNonCapData As DataTable
                            odtNonCapData = oDv_NonCap.ToTable
                            Dim oDvNonCapData As DataView = New DataView(odtNonCapData)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoCostAccrual()", sFuncName)
                            If InsertIntoCostAccrual(oDvNonCapData, p_oCompany, file.Name, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Next
                Console.WriteLine("Insertion of Non capitation Cost Accrual datas is successful")

                oDv_NonCap.RowFilter = Nothing

                oDtGroup = oDv_NonCap.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COSTCENTER") Then
                        oDv_NonCap.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                " and IncuredMonth = '" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "

                        If oDv_NonCap.Count > 0 Then
                            Dim odtCapData_JE As DataTable
                            odtCapData_JE = oDv_NonCap.ToTable
                            Dim oDvCapData_JE As DataView = New DataView(odtCapData_JE)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateJE_NonCapitation_GLAR()", sFuncName)
                            If CreateJE_NonCapitation_GLAR(oDvCapData_JE, file, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateJE_NonCapitation_GLAP()", sFuncName)
                            If CreateJE_NonCapitation_GLAP(oDvCapData_JE, file, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Next
                Console.WriteLine("Journal Entry creation for cost accrual datas successful")
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessDatas_TPAListings = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessDatas_TPAListings = RTN_ERROR
        End Try
    End Function

    Private Function ProcessDatas_NonTPAListings(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessDatas_NonTPAListings"
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            oDv.RowFilter = "Type NOT LIKE 'CAPITATION*'"
            Dim odt As New DataTable
            odt = oDv.ToTable
            Dim oNewDv As DataView = New DataView(odt)

            If oNewDv.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping databased based on company code,incurred month and MBMS", sFuncName)
                'F2 - Company Code F18 Invoice
                Dim oDtGroup As DataTable = oNewDv.Table.DefaultView.ToTable(True, "F2", "F18", "CostCenter", "IncuredMonth")

                Console.WriteLine("Processing Datas for inserting into Accrual table")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
                        oNewDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F18 ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                           " and CostCenter='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' and IncuredMonth='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "'"

                        If oNewDv.Count > 0 Then
                            Dim odtAcrualDts As DataTable
                            odtAcrualDts = oNewDv.ToTable
                            Dim oDvAcrlDatas As DataView = New DataView(odtAcrualDts)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoSOAccural()", sFuncName)
                            If InsertIntoSOAccural(oDvAcrlDatas, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Next
                Console.WriteLine("Inserting into Accrual table is Successful")

                oNewDv.RowFilter = Nothing

                Console.WriteLine("Processing Datas for creating Journal entry for accrual table")
                oDtGroup = oNewDv.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
                        oNewDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                           " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'"

                        If oNewDv.Count > 0 Then
                            Dim odtAcrualJrnlDts As DataTable
                            odtAcrualJrnlDts = oNewDv.ToTable
                            Dim oDvAcrlJrnlDatas As DataView = New DataView(odtAcrualJrnlDts)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateJournalEntry()", sFuncName)
                            If CreateJournalEntry(oDvAcrlJrnlDatas, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Next
                Console.WriteLine("Journal entry creation for accrual table is successful")
            End If

            oDv.RowFilter = Nothing

            oDv.RowFilter = "Type LIKE 'CAPITATION*'"
            Dim oCapdt As New DataTable
            oCapdt = oDv.ToTable
            Dim oCapDv As DataView = New DataView(oCapdt)

            If oCapDv.Count > 0 Then
                oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oDtGroup As DataTable = oCapDv.Table.DefaultView.ToTable(True, "F18")
                For k As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                        Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()
                        Dim sQuery As String = String.Empty

                        sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" SET ""U_status"" = 'C' " & _
                                 " WHERE ""U_invoice"" = '" & sInvoice & "' AND ""U_status"" = 'O' "

                        oRecordSet.DoQuery(sQuery)

                    End If
                Next
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessDatas_NonTPAListings = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessDatas_NonTPAListings = RTN_ERROR
        End Try
    End Function

    Private Function CreateJournalEntry(ByVal odv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateJournalEntry"
        Dim sSql As String = String.Empty
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCreditAct As String = String.Empty
        Dim sDebitAct As String = String.Empty
        Dim sCompCode As String = String.Empty
        Dim dTotalSal As Double = 0.0
        Dim dPayComp As Double = 0.0
        Dim iErrCode As Long
        Dim sCostCenter As String = String.Empty
        Dim sIncuredMnth As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS007_GL"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" WHERE A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sCreditAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Credit Account is " & sCreditAct, sFuncName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS007_GL"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" WHERE A.""U_ActType"" = 'D'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sDebitAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Debit Account is " & sDebitAct, sFuncName)

            If sCreditAct = "" Then
                sErrDesc = "Credit account should not be empty"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sDebitAct = "" Then
                sErrDesc = "Debit Account should not be empty"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dTotalSal = 0
            For i As Integer = 0 To odv.Count - 1
                dPayComp = CDbl(odv(i)(20).ToString.Trim)
                dTotalSal = dTotalSal + dPayComp
            Next

            If dTotalSal > 0 Then
                sCompCode = odv(0)(1).ToString.Trim
                sCostCenter = odv(0)(48).ToString.Trim
                sIncuredMnth = odv(0)(50).ToString.Trim
                Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
                Dim dt As Date = CDate(sIncuredMnth.Substring(0, iIndex))

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                oJournalEntry.TaxDate = dt
                oJournalEntry.ReferenceDate = dt
                oJournalEntry.Lines.ShortName = sCreditAct
                oJournalEntry.Lines.Credit = dTotalSal
                If Not (sCostCenter = String.Empty) Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    ' oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If
                If Not (sCostCenter = String.Empty) Then
                    oJournalEntry.Memo = "Estimated Sales for " & sCompCode & " and MBMS " & sCostCenter
                Else
                    oJournalEntry.Memo = "Estimated Sales for " & sCompCode
                End If

                oJournalEntry.Lines.Add()

                oJournalEntry.Lines.AccountCode = sDebitAct
                oJournalEntry.Lines.Debit = dTotalSal

                If oJournalEntry.Add() <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim sJournalEntryNo, sTransId As Integer
                    p_oCompany.GetNewObjectCode(sTransId)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

                    Dim oRecordSet As SAPbobsCOM.Recordset
                    Dim sQuery As String
                    oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    sQuery = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & sTransId & "'"
                    oRecordSet.DoQuery(sQuery)
                    If oRecordSet.RecordCount > 0 Then
                        sJournalEntryNo = oRecordSet.Fields.Item("Number").Value
                    End If

                    Console.WriteLine("Document Created Successfully :: " & sJournalEntryNo)

                    Dim oDtGroup As DataTable = odv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" SET ""U_Journal"" = '" & sJournalEntryNo & "',""U_JrnlEntry"" = '" & sTransId & "'" & _
                                     " WHERE ""U_company_code"" = '" & sCompCode & "' AND ""U_OcrCode"" = '" & sCostCenter & "' AND ""U_invoice"" = '" & sInvoice & "'" & _
                                     " AND ""U_Incurred_month"" = '" & sIncuredMnth & "' AND IFNULL(""U_Journal"",'') = '' AND IFNULL(""U_JrnlEntry"",'') = ''"
                            oRecordSet.DoQuery(sQuery)

                        End If
                    Next
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateJournalEntry = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateJournalEntry = RTN_ERROR
        End Try
    End Function

    Private Function CreateCostAccrualJE(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateCostAccrualJE"
        Dim sSql As String = String.Empty
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCreditAct As String = String.Empty
        Dim sDebitAct As String = String.Empty
        Dim sClincCode As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim dTotvalue As Double = 0.0
        Dim dPayComp As Double = 0.0
        Dim iErrCode As Long
        Dim sIncurMnth As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)


            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_REV_GL"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE UPPER(A.""U_Filecode"") = 'MS007' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sCreditAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_REV_GL"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE UPPER(A.""U_Filecode"") = 'MS007' AND A.""U_ActType"" = 'D'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sDebitAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            If sCreditAct = "" Then
                sErrDesc = "Credit account should not be empty"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sDebitAct = "" Then
                sErrDesc = "Debit Account should not be empty"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dTotvalue = 0
            For i As Integer = 0 To oDv.Count - 1
                dPayComp = CDbl(oDv(i)(20).ToString.Trim)
                dTotvalue = dTotvalue + dPayComp
            Next

            If dTotvalue > 0 Then
                Dim sSubCode As String = String.Empty
                sSubCode = oDv(0)(2).ToString.Trim
                sClincCode = oDv(0)(4).ToString.Trim
                'sClincCode = p_oCompDef.sCOAcrlCardCode
                sCostCenter = oDv(0)(48).ToString.Trim
                sIncurMnth = oDv(0)(50).ToString.Trim
                Dim iIndex As Integer = sIncurMnth.IndexOf(" ")
                Dim dt As Date = CDate(sIncurMnth.Substring(0, iIndex))

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                oJournalEntry.TaxDate = dt
                oJournalEntry.ReferenceDate = dt
                If sCostCenter <> String.Empty Then
                    oJournalEntry.Memo = "Estimated Cost for " & sClincCode & sSubCode & " and MBMS " & sCostCenter
                Else
                    oJournalEntry.Memo = "Estimated Cost for " & sClincCode & sSubCode
                End If

                oJournalEntry.Lines.ShortName = sCreditAct
                oJournalEntry.Lines.Credit = dTotvalue
                If Not sCostCenter = String.Empty Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                oJournalEntry.Lines.Add()

                oJournalEntry.Lines.AccountCode = sDebitAct
                oJournalEntry.Lines.Debit = dTotvalue
                If Not sCostCenter = String.Empty Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                If oJournalEntry.Add() <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim sJournalEntryNo, sTransId As Integer
                    p_oCompany.GetNewObjectCode(sTransId)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

                    Dim oRs As SAPbobsCOM.Recordset
                    Dim sQuery As String

                    sQuery = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & sTransId & "'"
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery(sQuery)
                    If oRs.RecordCount > 0 Then
                        sJournalEntryNo = oRs.Fields.Item("Number").Value
                    End If

                    Console.WriteLine("Document Created Successfully :: " & sJournalEntryNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_Journal"" = '" & sJournalEntryNo & "',""U_JrnlEntry"" = '" & sTransId & "' " & _
                                     " WHERE ""U_cln_code"" = '" & sClincCode & "' AND ""U_OcrCode"" = '" & sCostCenter & "' AND ""U_invoice"" = '" & sInvoice & "'" & _
                                     " AND ""U_incurred_month"" = '" & sIncurMnth & "' AND IFNULL(""U_Journal"",'') = '' AND IFNULL(""U_JrnlEntry"",'') = ''"
                            oRs.DoQuery(sQuery)
                        End If
                    Next


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateCostAccrualJE = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateCostAccrualJE = RTN_ERROR
        End Try
    End Function

    Private Function CreateJE_NonCapitation_GLAR(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateJE_NonCapitation_GLAR"
        Dim sSql As String = String.Empty
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCreditAct As String = String.Empty
        Dim sDebitAct As String = String.Empty
        Dim sClincCode As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim dTotvalue As Double = 0.0
        Dim dPayComp As Double = 0.0
        Dim iErrCode As Long
        Dim sIncurMnth As String = String.Empty
        Dim sSource As String = file.Name.Substring(0, 5)

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAR_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE UPPER(A.""U_FileCode"") = 'MS007' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sCreditAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAR_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE UPPER(A.""U_FileCode"") = 'MS007' AND A.""U_ActType"" = 'D'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sDebitAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            If sCreditAct = "" Then
                sErrDesc = "Credit account should not be empty"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sDebitAct = "" Then
                sErrDesc = "Debit Account should not be empty"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dTotvalue = 0
            For i As Integer = 0 To oDv.Count - 1
                dPayComp = CDbl(oDv(i)(20).ToString.Trim)
                dTotvalue = dTotvalue + dPayComp
            Next

            If dTotvalue > 0 Then
                Dim sCompCode As String = String.Empty
                sCompCode = oDv(0)(1).ToString.Trim
                sClincCode = oDv(0)(4).ToString.Trim
                'sClincCode = p_oCompDef.sCOAcrlCardCode
                sCostCenter = oDv(0)(48).ToString.Trim
                sIncurMnth = oDv(0)(50).ToString.Trim
                Dim iIndex As Integer = sIncurMnth.IndexOf(" ")
                Dim dt As Date = CDate(sIncurMnth.Substring(0, iIndex))

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                oJournalEntry.TaxDate = dt
                oJournalEntry.ReferenceDate = dt
                If sCostCenter <> String.Empty Then
                    oJournalEntry.Memo = "Est TPA Claim:" & sCompCode & " and MBMS " & sCostCenter
                Else
                    oJournalEntry.Memo = "Est TPA Claim:" & sCompCode
                End If

                oJournalEntry.Lines.ShortName = sCreditAct
                oJournalEntry.Lines.Credit = dTotvalue
                If Not sCostCenter = String.Empty Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                oJournalEntry.Lines.Add()

                oJournalEntry.Lines.AccountCode = sDebitAct
                oJournalEntry.Lines.Debit = dTotvalue
                If Not sCostCenter = String.Empty Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                If oJournalEntry.Add() <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim sJournalEntryNo, sTransId As Integer
                    p_oCompany.GetNewObjectCode(sTransId)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

                    Dim oRs As SAPbobsCOM.Recordset
                    Dim sQuery As String

                    sQuery = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & sTransId & "'"
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery(sQuery)
                    If oRs.RecordCount > 0 Then
                        sJournalEntryNo = oRs.Fields.Item("Number").Value
                    End If

                    Console.WriteLine("Document Created Successfully :: " & sJournalEntryNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_Glar_NC_DocNum"" = '" & sJournalEntryNo & "',""U_Glar_NC_DocEntry"" = '" & sTransId & "' " & _
                                     " WHERE ""U_OcrCode"" = '" & sCostCenter & "' AND ""U_invoice"" = '" & sInvoice & "'" & _
                                     " AND ""U_incurred_month"" = '" & sIncurMnth & "' AND IFNULL(""U_Glar_NC_DocNum"",'') = '' AND IFNULL(""U_Glar_NC_DocEntry"",'') = ''"
                            oRs.DoQuery(sQuery)
                        End If
                    Next


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateJE_NonCapitation_GLAR = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateJE_NonCapitation_GLAR = RTN_ERROR
        End Try
    End Function

    Private Function CreateJE_NonCapitation_GLAP(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateJE_NonCapitation_GLAP"
        Dim sSql As String = String.Empty
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCreditAct As String = String.Empty
        Dim sDebitAct As String = String.Empty
        Dim sClincCode As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim dTotvalue As Double = 0.0
        Dim dPayComp As Double = 0.0
        Dim iErrCode As Long
        Dim sIncurMnth As String = String.Empty
        Dim sSource As String = file.Name.Substring(0, 5)

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAP_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE UPPER(A.""U_FileCode"") = 'MS007' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sCreditAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAP_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE UPPER(A.""U_FileCode"") = 'MS007' AND A.""U_ActType"" = 'D'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sDebitAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            If sCreditAct = "" Then
                sErrDesc = "Credit account should not be empty"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sDebitAct = "" Then
                sErrDesc = "Debit Account should not be empty"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dTotvalue = 0
            For i As Integer = 0 To oDv.Count - 1
                dPayComp = CDbl(oDv(i)(20).ToString.Trim)
                dTotvalue = dTotvalue + dPayComp
            Next

            If dTotvalue > 0 Then
                Dim sSubCode As String = String.Empty
                sSubCode = oDv(0)(2).ToString.Trim
                sClincCode = oDv(0)(4).ToString.Trim
                'sClincCode = p_oCompDef.sCOAcrlCardCode
                sCostCenter = oDv(0)(48).ToString.Trim
                sIncurMnth = oDv(0)(50).ToString.Trim
                Dim iIndex As Integer = sIncurMnth.IndexOf(" ")
                Dim dt As Date = CDate(sIncurMnth.Substring(0, iIndex))

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                oJournalEntry.TaxDate = dt
                oJournalEntry.ReferenceDate = dt
                If sCostCenter <> String.Empty Then
                    oJournalEntry.Memo = "Est TPA Reimbuse:" & sClincCode & sSubCode & " and MBMS " & sCostCenter
                Else
                    oJournalEntry.Memo = "Est TPA Reimbuse:" & sClincCode & sSubCode
                End If

                oJournalEntry.Lines.ShortName = sCreditAct
                oJournalEntry.Lines.Credit = dTotvalue
                If Not sCostCenter = String.Empty Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                oJournalEntry.Lines.Add()

                oJournalEntry.Lines.AccountCode = sDebitAct
                oJournalEntry.Lines.Debit = dTotvalue
                If Not sCostCenter = String.Empty Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                If oJournalEntry.Add() <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim sJournalEntryNo, sTransId As Integer
                    p_oCompany.GetNewObjectCode(sTransId)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

                    Dim oRs As SAPbobsCOM.Recordset
                    Dim sQuery As String

                    sQuery = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & sTransId & "'"
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery(sQuery)
                    If oRs.RecordCount > 0 Then
                        sJournalEntryNo = oRs.Fields.Item("Number").Value
                    End If

                    Console.WriteLine("Document Created Successfully :: " & sJournalEntryNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_Glap_NC_DocNum"" = '" & sJournalEntryNo & "',""U_Glap_NC_DocEntry"" = '" & sTransId & "' " & _
                                     " WHERE ""U_OcrCode"" = '" & sCostCenter & "' AND ""U_invoice"" = '" & sInvoice & "'" & _
                                     " AND ""U_incurred_month"" = '" & sIncurMnth & "' AND IFNULL(""U_Glap_NC_DocNum"",'') = '' AND IFNULL(""U_Glap_NC_DocEntry"",'') = ''"
                            oRs.DoQuery(sQuery)
                        End If
                    Next


                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateJE_NonCapitation_GLAP = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateJE_NonCapitation_GLAP = RTN_ERROR
        End Try
    End Function

End Module
