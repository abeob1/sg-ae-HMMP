Module modARInvoice_CardCode_Writeoff

    Private dtSoAcrlInvList As DataTable
    Private dtCostAcrlInvList As DataTable

    Public Function ProcessARInvoice_CardCode_Writeoff(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessARInvoice_CardCode_Writeoff"
        Dim sSQL As String

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE IFNULL(""U_status"",'O') = 'C' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtSoAcrlInvList = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

            sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_New_Status"",'O') = 'C' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtCostAcrlInvList = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

            Dim odtDatatable As DataTable
            odtDatatable = oDv.ToTable

            odtDatatable.Columns.Add("CostCenter", GetType(String))
            odtDatatable.Columns.Add("Insurer", GetType(String))
            odtDatatable.Columns.Add("IncuredMonth", GetType(Date))
            odtDatatable.Columns.Add("Type", GetType(String))
            odtDatatable.Columns.Add("AcrlType", GetType(String))

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

                    Dim sInvoice As String = odtDatatable.Rows(intRow).Item(17).ToString.Trim
                    dtSoAcrlInvList.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtSoAcrlInvList.DefaultView.Count > 0 Then
                        sErrDesc = "Status already closed for invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    dtCostAcrlInvList.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtCostAcrlInvList.DefaultView.Count > 0 Then
                        sErrDesc = "Status already closed for invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Dim sNewType As String = String.Empty
                    sSQL = "SELECT ""U_Type"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE ""U_invoice"" = '" & sInvoice & "' "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    sNewType = GetStringValue(sSQL, p_oCompDef.sSAPDBName)

                    If sNewType = "" Then
                        sSQL = "SELECT ""U_Type"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE ""U_invoice"" = '" & sInvoice & "' "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                        sNewType = GetStringValue(sSQL, p_oCompDef.sSAPDBName)
                    End If

                    Dim sType As String
                    Dim sArCode As String = "C" & sCompCode
                    sSQL = "SELECT ""U_Type"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" WHERE ""CardCode"" = '" & sArCode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    sType = GetStringValue(sSQL, p_oCompDef.sSAPDBName)

                    If sType = "" Then
                        sType = p_oCompDef.sType
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
                    odtDatatable.Rows(intRow)("AcrlType") = sNewType.ToUpper

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

                If oDvFinalView.Count > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoTable()", sFuncName)

                    Console.WriteLine("Inserting datas in AR Table")
                    If InsertIntoARTable_Writeoff(oDvFinalView, file.Name, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Console.WriteLine("Data insert into AR Table Successful")

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Group by clinic code to check in TPA Listing table", sFuncName)
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
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessDatas_NonTPAListings()", sFuncName)
                                    Dim odtTpaList As DataTable
                                    odtTpaList = oDvFinalView.ToTable
                                    Dim oDvTpaList As DataView = New DataView(odtTpaList)
                                    If ProcessDatas_TPAListings(oDvTpaList, p_oCompany, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                Else
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessDatas_NonTPAListings()", sFuncName)
                                    Dim odtNonTpaList As DataTable
                                    odtNonTpaList = oDvFinalView.ToTable
                                    Dim oDvNonTpaList As DataView = New DataView(odtNonTpaList)
                                    If ProcessDatas_NonTPAListings(oDvNonTpaList, p_oCompany, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                            End If
                        End If
                    Next

                    'oDvFinalView.RowFilter = Nothing

                    'Console.WriteLine("Updating the Accural tables")
                    'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Updating the Accural tables", sFuncName)
                    'If UpdateSOAccrualTable(p_oCompany, oDvFinalView, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Accural table updation is success", sFuncName)
                    'Console.WriteLine("Status updation in Accrual table is success")
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

            Console.WriteLine("AR Writeoff file processed successfully")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessARInvoice_CardCode_Writeoff = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Console.WriteLine(sErrDesc)
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
            ProcessARInvoice_CardCode_Writeoff = RTN_ERROR
        End Try
    End Function

    Private Function InsertIntoARTable_Writeoff(ByVal oDv As DataView, ByVal sFileName As String, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "InsertIntoARTable_Writeoff"
        Dim sSql As String = String.Empty
        Dim sCompCode As String = String.Empty
        Dim sClinicCode As String = String.Empty

        Try

            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            For i As Integer = 1 To oDv.Count - 1
                If Not (oDv(i)(1).ToString.Trim = String.Empty) Then
                    Console.WriteLine("Inserting Line Num : " & i)
                    sSql = String.Empty

                    sCompCode = oDv(i)(1).ToString
                    sCompCode = "C" & oDv(i)(1).ToString

                    sClinicCode = oDv(i)(4).ToString
                    sClinicCode = "V" & oDv(i)(4).ToString

                    sSql = " INSERT INTO " & p_oCompDef.sSAPDBName & ".""@AE_MS007_ARWRITEOF""(""Code"",""Name"",""U_company"",""U_company_code"",""U_C"",""U_scheme_code"",""U_cln_code"",""U_m_id_type"",""U_m_id"",""U_m_lastname"",""U_m_given_name"",""U_m_christian""," & _
                            " ""U_relation"",""U_id_type"",""U_id"",""U_lastname"",""U_given_name"",""U_christian"",""U_txn_date"",""U_invoice"",""U_treatment"",""U_charge"",""U_pay_comp"",""U_pay_client"",""U_diag"",""U_diag_desc"", " & _
                            " ""U_refer_from_name"",""U_policy_num"",""U_cert_num"",""U_treat_code"",""U_remark_fg"",""U_remark1"",""U_paiddate"",""U_status"",""U_status_code"",""U_cust_no"",""U_scheme_remark"",""U_dept1"",""U_dept2""," & _
                            " ""U_dept3"",""U_ds1"",""U_ds2"",""U_ds3"",""U_in_time"",""U_insco"",""U_sl_fr"",""U_sl_to"",""U_CompTotRecCnt"",""U_CompTotBillAmt"",""U_scheme_desc"",""U_OcrCode"",""U_Insurer"",""U_Incurred_month"",""U_ar_code"",""U_ap_code"",""U_Type"",""U_FileName"")" & _
                            " Values ((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS007_ARWRITEOF""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS007_ARWRITEOF""), " & _
                            " '" & oDv(i)(0).ToString & "','" & oDv(i)(1).ToString & "','" & oDv(i)(2).ToString & "', " & _
                            " '" & oDv(i)(3).ToString & "','" & oDv(i)(4).ToString & "','" & oDv(i)(5).ToString & "', '" & oDv(i)(6).ToString & "','" & oDv(i)(7).ToString & "', " & _
                            " '" & oDv(i)(8).ToString & "','" & oDv(i)(9).ToString & "','" & oDv(i)(10).ToString & "', '" & oDv(i)(11).ToString & "','" & oDv(i)(12).ToString & "', " & _
                            " '" & oDv(i)(13).ToString & "','" & oDv(i)(14).ToString & "','" & oDv(i)(15).ToString & "', '" & oDv(i)(16).ToString & "','" & oDv(i)(17).ToString & "', " & _
                            " '" & oDv(i)(18).ToString & "','" & oDv(i)(19).ToString & "','" & oDv(i)(20).ToString & "', '" & oDv(i)(21).ToString & "','" & oDv(i)(22).ToString & "', " & _
                            " '" & oDv(i)(23).ToString & "','" & oDv(i)(24).ToString & "','" & oDv(i)(25).ToString & "', '" & oDv(i)(26).ToString & "','" & oDv(i)(27).ToString & "', " & _
                            " '" & oDv(i)(28).ToString & "','" & oDv(i)(29).ToString & "','" & oDv(i)(30).ToString & "', '" & oDv(i)(31).ToString & "','" & oDv(i)(32).ToString & "', " & _
                            " '" & oDv(i)(33).ToString & "','" & oDv(i)(34).ToString & "','" & oDv(i)(35).ToString & "', '" & oDv(i)(36).ToString & "','" & oDv(i)(37).ToString & "', " & _
                            " '" & oDv(i)(38).ToString & "','" & oDv(i)(39).ToString & "','" & oDv(i)(40).ToString & "', '" & oDv(i)(41).ToString & "','" & oDv(i)(42).ToString & "', " & _
                            " '" & oDv(i)(43).ToString & "','" & oDv(i)(44).ToString & "','" & oDv(i)(45).ToString & "', '" & oDv(i)(46).ToString & "','" & oDv(i)(47).ToString & "', " & _
                            " '" & oDv(i)(48).ToString & "','" & oDv(i)(49).ToString & "','" & oDv(i)(50).ToString & "','" & sCompCode & "','" & sClinicCode & "','" & oDv(i)(51).ToString & "','" & sFileName & "' )"

                    oRecSet.DoQuery(sSql)
                End If
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoARTable_Writeoff = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while executing query", sFuncName)
            InsertIntoARTable_Writeoff = RTN_ERROR
            Throw New Exception(ex.Message)
        End Try

    End Function

    Private Function ProcessDatas_TPAListings(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessDatas_TPAListings"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oDv.RowFilter = "AcrlType NOT LIKE 'CAPITATION*'"
            Dim odt As New DataTable
            odt = oDv.ToTable
            Dim oNewDv As DataView = New DataView(odt)

            If oNewDv.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping databased based on company code,incurred month and MBMS", sFuncName)
                'F2 - Company Code
                Dim oDtGroup As DataTable = oNewDv.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")

                Console.WriteLine("Processing Datas for A/R invoice Creation and Reversal Journal")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
                        oNewDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                           " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "

                        If oNewDv.Count > 0 Then
                            Dim odtARInvDts As DataTable
                            odtARInvDts = oNewDv.ToTable
                            Dim oDvARInvDts As DataView = New DataView(odtARInvDts)
                            Console.WriteLine("Creating Reverse Journal/Line : " & i)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateRevJournal_GLAR()", sFuncName)
                            If CreateRevJournal_GLAR(p_oCompany, oDvARInvDts, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            Console.WriteLine("Reverse journal created successfully")
                        End If
                    End If
                Next
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

    Private Function CreateRevJournal_GLAR(ByVal oCompany As SAPbobsCOM.Company, ByVal odv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateRevJournal_GLAR"
        Dim sCostCenter As String = String.Empty
        Dim dPayCompAmt, dTotPayCompAmt As Double
        Dim sDocDate As String = String.Empty
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCompCode As String = String.Empty
        Dim sCreditAct As String = String.Empty
        Dim sDebitAct As String = String.Empty
        Dim iErrCode As Integer
        Dim sXcelInvNo As String
        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sDocDate = file.Name.Substring(9, 8)

            Dim dt As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd"}
            Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

            sSQL = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAR_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSQL = sSQL & " WHERE A.""U_FileCode"" = 'MS007' AND A.""U_ActType"" = 'D'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            sCreditAct = GetStringValue(sSQL, p_oCompDef.sSAPDBName)

            sSQL = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAR_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSQL = sSQL & " WHERE A.""U_FileCode"" = 'MS007' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            sDebitAct = GetStringValue(sSQL, p_oCompDef.sSAPDBName)

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

            dTotPayCompAmt = 0
            dPayCompAmt = 0.0


            Dim oDtGroup As DataTable = odv.Table.DefaultView.ToTable(True, "F18")
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    sXcelInvNo = oDtGroup.Rows(i).Item(0).ToString.Trim()

                    Dim oRs As SAPbobsCOM.Recordset
                    Dim sQuery As String = String.Empty
                    oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    sQuery = "SELECT IFNULL(SUM(""U_pay_comp""),0) AS ""U_pay_comp"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" " & _
                           " WHERE ""U_invoice"" = '" & sXcelInvNo & "'  " & _
                           " AND IFNULL(""U_Glar_NC_Rev_DocNum"",'') = '' AND IFNULL(""U_Glar_NC_Rev_Entry"",'') = '' "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)

                    oRs.DoQuery(sQuery)
                    If oRs.RecordCount > 0 Then
                        dPayCompAmt = oRs.Fields.Item("U_pay_comp").Value
                        dTotPayCompAmt = dTotPayCompAmt + dPayCompAmt
                    End If
                End If
            Next

            If dTotPayCompAmt > 0 Then
                sCompCode = odv(0)(1).ToString.Trim
                sCostCenter = odv(0)(48).ToString.Trim

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                oJournalEntry.TaxDate = dt
                oJournalEntry.ReferenceDate = dt
                oJournalEntry.Memo = "Reversal of Estimated TPA Claim for " & sCompCode

                oJournalEntry.Lines.ShortName = sCreditAct
                oJournalEntry.Lines.Credit = dTotPayCompAmt

                oJournalEntry.Lines.Add()

                oJournalEntry.Lines.AccountCode = sDebitAct
                oJournalEntry.Lines.Debit = dTotPayCompAmt
                If Not (sCostCenter = String.Empty) Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                If oJournalEntry.Add() <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim sJournalEntryNo, sTransId As Integer
                    p_oCompany.GetNewObjectCode(sTransId)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

                    Dim sQuery As String
                    Dim oRecordSet As SAPbobsCOM.Recordset
                    oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    sQuery = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & sTransId & "'"
                    oRecordSet.DoQuery(sQuery)
                    If oRecordSet.RecordCount > 0 Then
                        sJournalEntryNo = oRecordSet.Fields.Item("Number").Value
                    End If

                    Console.WriteLine("Document Created Successfully :: " & sJournalEntryNo)

                    odv.RowFilter = Nothing
                    oDtGroup = odv.Table.DefaultView.ToTable(True, "F18")
                    For i As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            sXcelInvNo = oDtGroup.Rows(i).Item(0).ToString.Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_New_Status"" = 'C',""U_Glar_NC_Rev_DocNum"" = '" & sJournalEntryNo & "',""U_Glar_NC_Rev_Entry"" = '" & sTransId & "' " & _
                                     " WHERE ""U_OcrCode"" = '" & sCostCenter & "' AND ""U_invoice"" = '" & sXcelInvNo & "' " & _
                                     " AND IFNULL(""U_Glar_NC_Rev_DocNum"",'') = '' AND IFNULL(""U_Glar_NC_Rev_Entry"",'') = ''"
                            oRecordSet.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateRevJournal_GLAR = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateRevJournal_GLAR = RTN_ERROR
        End Try
    End Function

    Private Function ProcessDatas_NonTPAListings(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessDatas_NonTPAListings"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oDv.RowFilter = "AcrlType NOT LIKE 'CAPITATION*'"
            Dim odt As New DataTable
            odt = oDv.ToTable
            Dim oNewDv As DataView = New DataView(odt)

            If oNewDv.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping databased based on company code,incurred month and MBMS", sFuncName)
                'F2 - Company Code 
                Dim oDtGroup As DataTable = oNewDv.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")

                Console.WriteLine("Processing Datas for creating Reversal Journal entry")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
                        oNewDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                           " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "

                        If oNewDv.Count > 0 Then
                            Dim odtAcrualDts As DataTable
                            odtAcrualDts = oNewDv.ToTable
                            Dim oDvAcrlDatas As DataView = New DataView(odtAcrualDts)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateRevJournal_NonTpaListing()", sFuncName)
                            If CreateRevJournal_NonTpaListing(p_oCompany, oDvAcrlDatas, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            Console.WriteLine("Reverse journal entry created succesfully for grouped data line " & i)

                        End If
                    End If
                Next
                Console.WriteLine("Creation of Reverse Journal is Successful")
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

    Private Function CreateRevJournal_NonTpaListing(ByVal oCompany As SAPbobsCOM.Company, ByVal odv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateRevJournal_NonTpaListing"
        Dim sSql As String = String.Empty
        Dim sCostCenter As String = String.Empty
        Dim dPayCompAmt, dTotPayCompAmt As Double
        Dim sDocDate As String = String.Empty
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCompCode As String = String.Empty
        Dim sCreditAct As String = String.Empty
        Dim sDebitAct As String = String.Empty
        Dim iErrCode As Integer
        Dim sXcelInvNo As String
        Dim sCompCode1 As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sDocDate = file.Name.Substring(9, 8)

            Dim dt As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS007_GL_REV"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE A.""U_FileCode"" = 'MS007' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sCreditAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS007_GL_REV"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE A.""U_FileCode"" = 'MS007' AND A.""U_ActType"" = 'D'"
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

            dTotPayCompAmt = 0
            dPayCompAmt = 0.0

            Dim oDtGroup As DataTable = odv.Table.DefaultView.ToTable(True, "F18")
            For k As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                    Dim oRs As SAPbobsCOM.Recordset
                    Dim sQuery As String = String.Empty
                    oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    sQuery = "SELECT IFNULL(SUM(""U_total_sales""),0) AS ""U_total_sales"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL""  " & _
                           " WHERE ""U_invoice"" = '" & sInvoice & "' " & _
                           " AND IFNULL(""U_RevJournalEntry"",'') = '' AND IFNULL(""U_RevJrnlNo"",'') = '' "
                    oRs.DoQuery(sQuery)
                    If oRs.RecordCount > 0 Then
                        dPayCompAmt = oRs.Fields.Item("U_total_sales").Value
                        dTotPayCompAmt = dTotPayCompAmt + dPayCompAmt
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)

                End If
            Next

            If dTotPayCompAmt > 0 Then
                sCompCode = odv(0)(1).ToString.Trim
                sCostCenter = odv(0)(48).ToString.Trim

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                oJournalEntry.TaxDate = dt
                oJournalEntry.ReferenceDate = dt
                oJournalEntry.Memo = "Reversal of Estimated sales for " & sCompCode

                oJournalEntry.Lines.ShortName = sCreditAct
                oJournalEntry.Lines.Credit = dTotPayCompAmt

                oJournalEntry.Lines.Add()

                oJournalEntry.Lines.AccountCode = sDebitAct
                oJournalEntry.Lines.Debit = dTotPayCompAmt
                If Not (sCostCenter = String.Empty) Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                If oJournalEntry.Add() <> 0 Then
                    oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim sJournalEntryNo, sTransId As Integer
                    p_oCompany.GetNewObjectCode(sTransId)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

                    Dim sQuery As String
                    Dim oRecordSet As SAPbobsCOM.Recordset
                    oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    sQuery = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & sTransId & "'"
                    oRecordSet.DoQuery(sQuery)
                    If oRecordSet.RecordCount > 0 Then
                        sJournalEntryNo = oRecordSet.Fields.Item("Number").Value
                    End If

                    Console.WriteLine("Document Created Successfully :: " & sJournalEntryNo)

                    odv.RowFilter = Nothing
                    oDtGroup = odv.Table.DefaultView.ToTable(True, "F18")
                    For i As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            sXcelInvNo = oDtGroup.Rows(i).Item(0).ToString.Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" SET ""U_status"" = 'C',""U_RevJrnlNo"" = '" & sJournalEntryNo & "',""U_RevJournalEntry"" = '" & sTransId & "'" & _
                                     " WHERE ""U_company_code"" = '" & sCompCode & "' AND ""U_OcrCode"" = '" & sCostCenter & "' AND IFNULL(""U_RevJournalEntry"",'') = '' AND IFNULL(""U_RevJrnlNo"",'') = '' " & _
                                     " AND ""U_invoice"" = '" & sXcelInvNo & "' "

                            oRecordSet.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateRevJournal_NonTpaListing = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateRevJournal_NonTpaListing = RTN_ERROR
        End Try
    End Function

End Module
