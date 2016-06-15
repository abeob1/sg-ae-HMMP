Module modARInvoice

    Private dtSoAcrlInvList As DataTable
    Private dtCostAcrlInvList As DataTable
    Private dtInsurerList As DataTable
    Private dtMBMSList As DataTable
    Private dtItemCode As DataTable
    Private dtCardCode As DataTable

    Public Function ProcessARInvoice(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessARInvoice"
        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE ""U_status"" = 'C' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtSoAcrlInvList = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

            'sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE ""U_status"" = 'C' "
            sSQL = "SELECT DISTINCT ""U_invoice"" FROM( " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_AP_Inv_DocEntry"",'') <> '' UNION ALL " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_Glap_NC_Rev_Entry"",'') <> '' UNION ALL " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_RevJournalEntry"",'') <> '' )T1 "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtCostAcrlInvList = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

            sSQL = "SELECT ""ItemCode"" FROM " & p_oCompDef.sSAPDBName & ".""OITM"" "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtItemCode = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

            sSQL = "SELECT ""CardCode"",""VatGroup"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"""
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            dtCardCode = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

            Dim odtDatatable As DataTable
            odtDatatable = oDv.ToTable

            odtDatatable.Columns.Add("CostCenter", GetType(String))
            odtDatatable.Columns.Add("Insurer", GetType(String))
            odtDatatable.Columns.Add("IncuredMonth", GetType(Date))
            odtDatatable.Columns.Add("Type", GetType(String))
            odtDatatable.Columns.Add("ApCode", GetType(String))
            odtDatatable.Columns.Add("AcrlType", GetType(String))

            For intRow As Integer = 0 To odtDatatable.Rows.Count - 1
                If Not (odtDatatable.Rows(intRow).Item(1).ToString.Trim() = String.Empty Or odtDatatable.Rows(intRow).Item(1).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
                    Console.WriteLine("Processing excel line " & intRow & " to get MBMS and Insurer from config table")

                    Dim sCompCode As String = odtDatatable.Rows(intRow).Item(1).ToString
                    Dim sCompName As String = odtDatatable.Rows(intRow).Item(0).ToString
                    sCompName = sCompName.Replace("'", " ")
                    Dim sSchemeCode As String = odtDatatable.Rows(intRow).Item(3).ToString
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

                    Dim sClinicCode As String = odtDatatable.Rows(intRow).Item(4).ToString
                    sSQL = "SELECT ""U_ap_code"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_TPA_APCODE"" WHERE UPPER(""U_cln_code"") = '" & sClinicCode.ToUpper() & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
                    Dim sApCode As String = GetStringValue(sSQL, p_oCompDef.sSAPDBName)

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
                    odtDatatable.Rows(intRow)("F24") = sDiagDesc
                    odtDatatable.Rows(intRow)("F30") = sRemarks
                    odtDatatable.Rows(intRow)("CostCenter") = sCostCenter
                    odtDatatable.Rows(intRow)("Insurer") = sInsurer
                    odtDatatable.Rows(intRow)("IncuredMonth") = dIncurMnth
                    odtDatatable.Rows(intRow)("Type") = sType.ToUpper
                    odtDatatable.Rows(intRow)("ApCode") = sApCode
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
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoARTable()", sFuncName)

                    Console.WriteLine("Inserting datas in AR Table")
                    If InsertIntoARTable(oDvFinalView, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Console.WriteLine("Data insert into AR Table Successful")

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Group by ApCode", sFuncName)

                    oDvFinalView.RowFilter = "ISNULL(ApCode,'') <> ''"

                    Dim oDtTpaDatas As New DataTable
                    oDtTpaDatas = oDvFinalView.ToTable
                    Dim oDvTpaDatas As DataView = New DataView(oDtTpaDatas)

                    If oDvTpaDatas.Count > 0 Then

                        oDvTpaDatas.RowFilter = "Type NOT LIKE 'CAPITATION*'"
                        Dim oNonCapDt As New DataTable
                        oNonCapDt = oDvTpaDatas.ToTable
                        Dim oNonCapDv As DataView = New DataView(oNonCapDt)

                        If oNonCapDv.Count > 0 Then
                            Dim oDtGroup As DataTable = oNonCapDv.Table.DefaultView.ToTable(True, "ApCode")
                            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "APCODE") Then
                                    oNonCapDv.RowFilter = "ApCode = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' "

                                    If oNonCapDv.Count > 0 Then
                                        Dim odtApDatas As New DataTable
                                        odtApDatas = oNonCapDv.ToTable
                                        Dim oDvApDatas As DataView = New DataView(odtApDatas)

                                        Console.WriteLine("Creating A/p invoice for " & oDtGroup.Rows(i).Item(0).ToString.Trim())
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateApInvoice_NonCapitation()", sFuncName)
                                        If CreateApInvoice_NonCapitation(oDvApDatas, file, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Console.WriteLine("A/p Invoice Created successfully")
                                    End If
                                End If
                            Next
                        End If

                        oDvTpaDatas.RowFilter = Nothing

                        oDvTpaDatas.RowFilter = "AcrlType NOT LIKE 'CAPITATION*'"
                        Dim oDtNonCap_AcrlType As New DataTable
                        oDtNonCap_AcrlType = oDvTpaDatas.ToTable
                        Dim oDvNonCap_AcrlType As DataView = New DataView(oDtNonCap_AcrlType)

                        If oDvNonCap_AcrlType.Count > 0 Then
                            Dim oDtGroup As DataTable = oDvNonCap_AcrlType.Table.DefaultView.ToTable(True, "ApCode", "CostCenter", "IncuredMonth")
                            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "APCODE") Then
                                    oDvNonCap_AcrlType.RowFilter = "ApCode = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                          " and IncuredMonth ='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "

                                    If oDvNonCap_AcrlType.Count > 0 Then
                                        Dim oDt_NonCapInvDts As DataTable
                                        oDt_NonCapInvDts = oDvNonCap_AcrlType.ToTable
                                        Dim oDv_NonCapInvDts As DataView = New DataView(oDt_NonCapInvDts)

                                        Console.WriteLine("Creating Reversal journal")
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateRevJournal_GLAP_NonCap()", sFuncName)
                                        If CreateRevJournal_GLAP_NonCap(oDv_NonCapInvDts, file, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Console.WriteLine("Reversal journal created successfully")
                                    End If
                                End If
                            Next
                        End If

                        oDvTpaDatas.RowFilter = Nothing

                        oDvTpaDatas.RowFilter = "Type LIKE 'CAPITATION*'"
                        Dim oDtCap As DataTable
                        oDtCap = oDvTpaDatas.ToTable
                        Dim oDvCap As DataView = New DataView(oDtCap)

                        If oDvCap.Count > 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Group capitation datas by ap code", sFuncName)
                            Dim oDtGroup As DataTable = oDvCap.Table.DefaultView.ToTable(True, "ApCode")
                            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "APCODE") Then
                                    oDvCap.RowFilter = "ApCode = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' "
                                    If oDvCap.Count > 0 Then
                                        Dim odtApDatas As New DataTable
                                        odtApDatas = oDvCap.ToTable
                                        Dim oDvApDatas As DataView = New DataView(odtApDatas)
                                        Console.WriteLine("Creating A/p invoice for " & oDtGroup.Rows(i).Item(0).ToString.Trim())
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateApInvoice_Capitation()", sFuncName)
                                        If CreateApInvoice_Capitation(oDvApDatas, file, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Console.WriteLine("A/p Invoice Created successfully")
                                    End If
                                End If
                            Next

                        End If

                        oDvTpaDatas.RowFilter = Nothing

                        oDvTpaDatas.RowFilter = "AcrlType LIKE 'CAPITATION*'"
                        Dim oDtCap_AcrlType As DataTable
                        oDtCap_AcrlType = oDvTpaDatas.ToTable
                        Dim oDvCap_AcrlType As DataView = New DataView(oDtCap_AcrlType)

                        If oDvCap_AcrlType.Count > 0 Then

                            Dim oDtGroup As DataTable = oDvCap_AcrlType.Table.DefaultView.ToTable(True, "ApCode", "CostCenter", "IncuredMonth")
                            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "APCODE") Then
                                    oDvCap_AcrlType.RowFilter = "ApCode = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                       " and IncuredMonth ='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "

                                    If oDvCap_AcrlType.Count > 0 Then
                                        Dim oDt_CapInvDts As DataTable
                                        oDt_CapInvDts = oDvCap_AcrlType.ToTable
                                        Dim oDv_CapInvDts As DataView = New DataView(oDt_CapInvDts)
                                        Console.WriteLine("Creating Reversal journal")
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateRevJouranl_CostAccrualJE()", sFuncName)
                                        If CreateRevJouranl_CostAccrualJE(oDv_CapInvDts, file, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        Console.WriteLine("Reversal journal created successfully")
                                    End If
                                End If
                            Next
                        End If
                    End If

                    'oDvFinalView.RowFilter = Nothing

                    'Console.WriteLine("Updating the Cost accrual datas")
                    'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling UpdateCostAccrual()", sFuncName)
                    'If UpdateCostAccrual(p_oCompany, oDvFinalView, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    'Console.WriteLine("Updation of Cost accrual datas is successful")

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
            ProcessARInvoice = RTN_SUCCESS
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
            ProcessARInvoice = RTN_ERROR

        End Try
    End Function

    Private Function InsertIntoARTable(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "InsertIntoARTable"
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
                    'sClinicCode = "V" & oDv(i)(4).ToString
                    sSql = "SELECT ""U_ap_code"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_TPA_APCODE"" WHERE UPPER(""U_cln_code"") = '" & sClinicCode.ToUpper() & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                    Dim sApCode As String = GetStringValue(sSql, p_oCompDef.sSAPDBName)

                    sSql = " INSERT INTO " & p_oCompDef.sSAPDBName & ".""@AE_MS007_AR""(""Code"",""Name"",""U_company"",""U_company_code"",""U_C"",""U_scheme_code"",""U_cln_code"",""U_m_id_type"",""U_m_id"",""U_m_lastname"",""U_m_given_name"",""U_m_christian""," & _
                            " ""U_relation"",""U_id_type"",""U_id"",""U_lastname"",""U_given_name"",""U_christian"",""U_txn_date"",""U_invoice"",""U_treatment"",""U_charge"",""U_pay_comp"",""U_pay_client"",""U_diag"",""U_diag_desc"", " & _
                            " ""U_refer_from_name"",""U_policy_num"",""U_cert_num"",""U_treat_code"",""U_remark_fg"",""U_remark1"",""U_paiddate"",""U_status"",""U_status_code"",""U_cust_no"",""U_scheme_remark"",""U_dept1"",""U_dept2""," & _
                            " ""U_dept3"",""U_ds1"",""U_ds2"",""U_ds3"",""U_in_time"",""U_insco"",""U_sl_fr"",""U_sl_to"",""U_CompTotRecCnt"",""U_CompTotBillAmt"",""U_scheme_desc"",""U_OcrCode"",""U_Insurer"",""U_Incurred_month"",""U_ar_code"",""U_ap_code"",""U_Type"")" & _
                            " Values ((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS007_AR""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS007_AR""), " & _
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
                            " '" & oDv(i)(48).ToString & "','" & oDv(i)(49).ToString & "','" & oDv(i)(50).ToString & "','" & sCompCode & "','" & sApCode & "','" & oDv(i)(51).ToString & "' )"

                    oRecSet.DoQuery(sSql)
                End If
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoARTable = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while executing query", sFuncName)
            InsertIntoARTable = RTN_ERROR
            Throw New Exception(ex.Message)
        End Try

    End Function

    Private Function CreateApInvoice_NonCapitation(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateApInvoice_NonCapitation"
        Dim sItemCode As String = String.Empty
        Dim sSql As String = String.Empty
        Dim dAmount As Double = 0.0
        Dim sClinicCode, sIncuredMnth, sCardCode, sVatGroup, sCostCenter As String
        Dim iCount, iErrCode As Integer
        Dim bLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'Dim iPos As Integer = file.Name.LastIndexOf("_")
            'Dim sCCode As String = file.Name.Substring(iPos + 1, (file.Name.Length - iPos - 1))
            'sCCode = sCCode.Replace(".xls", "")

            sClinicCode = oDv(0)(4).ToString.Trim
            sIncuredMnth = oDv(0)(50).ToString.Trim

            'sSql = "SELECT ""U_ap_code"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_TPA_APCODE"" WHERE UPPER(""U_cln_code"") = '" & sClinicCode.ToUpper() & "'"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            'sCardCode = GetStringValue(sSql, p_oCompDef.sSAPDBName)
            sCardCode = oDv(0)(52).ToString.Trim

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

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_ITEMCODE"" WHERE UPPER(""U_FileCode"") = 'MS007' AND UPPER(""U_Outnetwork"") = 'Y' " & _
                   " AND UPPER(IFNULL(""U_Type"",'')) <> 'CAPITATION'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sSAPDBName)

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

            Dim sDocDate As String
            sDocDate = file.Name.Substring(12, 8)

            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            Dim dPayComp, dTotPayComp As Double

            Dim oApInvoice As SAPbobsCOM.Documents
            oApInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            oApInvoice.CardCode = sCardCode
            oApInvoice.DocDate = dDocDate
            iCount = 1

            Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "CostCenter", "IncuredMonth")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas based on MBMS for creating journal entry", sFuncName)
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COSTCENTER") Then
                    oDv.RowFilter = "CostCenter = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' AND IncuredMonth = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "'"

                    If oDv.Count > 0 Then
                        sCostCenter = oDv(0)(48).ToString.Trim

                        dPayComp = 0
                        dTotPayComp = 0

                        For K As Integer = 0 To oDv.Count - 1
                            dPayComp = CDbl(oDv(K)(20).ToString.Trim)
                            dTotPayComp = dTotPayComp + dPayComp
                        Next

                        If dTotPayComp > 0 Then
                            If iCount > 1 Then
                                oApInvoice.Lines.Add()
                            End If

                            oApInvoice.Lines.ItemCode = sItemCode
                            oApInvoice.Lines.Quantity = 1
                            oApInvoice.Lines.Price = dTotPayComp
                            If Not (sVatGroup = String.Empty) Then
                                oApInvoice.Lines.VatGroup = sVatGroup
                            End If
                            If Not (sCostCenter = String.Empty) Then
                                oApInvoice.Lines.CostingCode = sCostCenter
                                oApInvoice.Lines.COGSCostingCode = sCostCenter
                            End If

                            iCount = iCount + 1
                            bLineAdded = True
                        End If

                    End If
                End If
            Next

            If bLineAdded = True Then
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

                    sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sSAPDBName & ".""OPCH"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSql)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    oDv.RowFilter = Nothing
                    oDtGroup = oDv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = String.Empty
                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_AP_Inv_DocNo"" = '" & iDocNo & "', ""U_AP_Inv_DocEntry"" = '" & iDocEntry & "'" & _
                                     " WHERE ""U_incurred_month"" = '" & sIncuredMnth & "' AND ""U_invoice"" = '" & sInvoice & "' AND IFNULL(""U_AP_Inv_DocEntry"",'') = '' "

                            objRS.DoQuery(sQuery)

                            sQuery = String.Empty
                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_MS007_AR"" SET ""U_APInvoiceNo"" = '" & iDocNo & "' " & _
                                     " WHERE ""U_invoice"" = '" & sInvoice & "' AND IFNULL(""U_APInvoiceNo"",'') = '' "
                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateApInvoice_NonCapitation = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateApInvoice_NonCapitation = RTN_ERROR
        End Try
    End Function

    Private Function CreateRevJournal_GLAP_NonCap(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateRevJournal_GLAP_NonCap"
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
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAP_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE A.""U_FileCode"" = 'MS007' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sCreditAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAP_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
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

            dTotvalue = 0
            dPayComp = 0

            Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F18")
            For k As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                    sSql = "SELECT IFNULL(SUM(""U_pay_comp""),0) AS ""U_pay_comp""  FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE ""U_invoice"" = '" & sInvoice & "' " & _
                           " AND IFNULL(""U_Glap_NC_Rev_DocNum"",'') = '' AND IFNULL(""U_Glap_NC_Rev_Entry"",'') = '' " 'AND ""U_status"" = 'O'
                    oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(sSql)
                    If oRecordSet.RecordCount > 0 Then
                        dPayComp = oRecordSet.Fields.Item("U_pay_comp").Value
                        dTotvalue = dTotvalue + dPayComp
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                End If
            Next

            'For i As Integer = 0 To oDv.Count - 1
            '    dPayComp = CDbl(oDv(i)(20).ToString.Trim)
            '    dTotvalue = dTotvalue + dPayComp
            'Next

            If dTotvalue > 0 Then
                sClincCode = oDv(0)(4).ToString.Trim
                sCostCenter = oDv(0)(48).ToString.Trim
                sIncurMnth = oDv(0)(50).ToString.Trim

                Dim sApCode As String = oDv(0)(52).ToString.Trim()

                Dim sDocDate As String
                sDocDate = file.Name.Substring(12, 8)
                Dim dt As Date
                Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                oJournalEntry.TaxDate = dt
                oJournalEntry.ReferenceDate = dt
                If sCostCenter <> "" Then
                    oJournalEntry.Memo = "Reversal of Est TPA Reimbuse for " & sClincCode & " and " & sCostCenter
                Else
                    oJournalEntry.Memo = "Reversal of Est TPA Reimbuse for " & sClincCode
                End If
                'oJournalEntry.Memo = "Reversal of Est TPA Reimbuse for " & sApCode

                oJournalEntry.Lines.ShortName = sDebitAct
                oJournalEntry.Lines.Credit = dTotvalue
                If Not sCostCenter = String.Empty Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                oJournalEntry.Lines.Add()

                oJournalEntry.Lines.AccountCode = sCreditAct
                oJournalEntry.Lines.Debit = dTotvalue
                If Not sCostCenter = String.Empty Then
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

                    Dim oRs As SAPbobsCOM.Recordset
                    Dim sQuery As String

                    sQuery = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & sTransId & "'"
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery(sQuery)
                    If oRs.RecordCount > 0 Then
                        sJournalEntryNo = oRs.Fields.Item("Number").Value
                    End If

                    Console.WriteLine("Document Created Successfully :: " & sJournalEntryNo)

                    oDv.RowFilter = Nothing

                    oDtGroup = oDv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_status"" = 'C',""U_Glap_NC_Rev_DocNum"" = '" & sJournalEntryNo & "',""U_Glap_NC_Rev_Entry"" = '" & sTransId & "' " & _
                                     " WHERE ""U_OcrCode"" = '" & sCostCenter & "' AND ""U_invoice"" = '" & sInvoice & "' AND ""U_incurred_month"" = '" & sIncurMnth & "' " & _
                                     " AND IFNULL(""U_Glap_NC_Rev_DocNum"",'') = '' AND IFNULL(""U_Glap_NC_Rev_Entry"",'') = ''"
                            oRs.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateRevJournal_GLAP_NonCap = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateRevJournal_GLAP_NonCap = RTN_ERROR
        End Try
    End Function

    Private Function CreateApInvoice_Capitation(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateApInvoice_Capitation"
        Dim sItemCode As String = String.Empty
        Dim sSql As String = String.Empty
        Dim dAmount As Double = 0.0
        Dim sClinicCode, sIncuredMnth, sCardCode, sVatGroup, sCostCenter As String
        Dim iCount, iErrCode As Integer
        Dim bLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sClinicCode = oDv(0)(4).ToString.Trim
            sIncuredMnth = oDv(0)(50).ToString.Trim

            'sSql = "SELECT ""U_ap_code"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_TPA_APCODE"" WHERE UPPER(""U_cln_code"") = '" & sClinicCode.ToUpper() & "'"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            'sCardCode = GetStringValue(sSql, p_oCompDef.sSAPDBName)
            sCardCode = oDv(0)(52).ToString.Trim

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

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_ITEMCODE"" WHERE UPPER(""U_FileCode"") = 'MS007' AND UPPER(""U_Outnetwork"") = 'Y' " & _
                   " AND UPPER(IFNULL(""U_Type"",'')) = 'CAPITATION'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sSAPDBName)

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

            Dim sDocDate As String
            sDocDate = file.Name.Substring(12, 8)
            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            Dim dPayComp, dTotPayComp As Double

            Dim oApInvoice As SAPbobsCOM.Documents
            oApInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            oApInvoice.CardCode = sCardCode
            oApInvoice.DocDate = dDocDate
            iCount = 1

            Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "CostCenter", "IncuredMonth")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas based on MBMS for creating journal entry", sFuncName)
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COSTCENTER") Then
                    oDv.RowFilter = "CostCenter = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and IncuredMonth ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "

                    If oDv.Count > 0 Then
                        sCostCenter = oDv(0)(48).ToString.Trim

                        dPayComp = 0
                        dTotPayComp = 0

                        For k As Integer = 0 To oDv.Count - 1
                            dPayComp = CDbl(oDv(k)(20).ToString.Trim)
                            dTotPayComp = dTotPayComp + dPayComp
                        Next

                        If dTotPayComp > 0 Then
                            If iCount > 1 Then
                                oApInvoice.Lines.Add()
                            End If

                            oApInvoice.Lines.ItemCode = sItemCode
                            oApInvoice.Lines.Quantity = 1
                            oApInvoice.Lines.Price = dTotPayComp
                            If Not (sVatGroup = String.Empty) Then
                                oApInvoice.Lines.VatGroup = sVatGroup
                            End If
                            If Not (sCostCenter = String.Empty) Then
                                oApInvoice.Lines.CostingCode = sCostCenter
                                oApInvoice.Lines.COGSCostingCode = sCostCenter
                            End If

                            iCount = iCount + 1
                            bLineAdded = True
                        End If

                    End If
                End If
            Next

            If bLineAdded = True Then

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

                    sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sSAPDBName & ".""OPCH"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSql)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    oDv.RowFilter = Nothing
                    oDtGroup = oDv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = String.Empty
                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_AP_Inv_DocNo"" = '" & iDocNo & "', ""U_AP_Inv_DocEntry"" = '" & iDocEntry & "'" & _
                                     " WHERE  ""U_incurred_month"" = '" & sIncuredMnth & "' AND ""U_invoice"" = '" & sInvoice & "' " & _
                                     " AND IFNULL(""U_AP_Inv_DocEntry"",'') = '' "

                            objRS.DoQuery(sQuery)

                            sQuery = String.Empty
                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_MS007_AR"" SET ""U_APInvoiceNo"" = '" & iDocNo & "' " & _
                                     " WHERE ""U_invoice"" = '" & sInvoice & "' AND IFNULL(""U_APInvoiceNo"",'') = '' "
                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateApInvoice_Capitation = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateApInvoice_Capitation = RTN_ERROR
        End Try
    End Function

    Private Function CreateRevJouranl_CostAccrualJE(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateRevJouranl_CostAccrualJE"
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
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_REV_GL"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE A.""U_Filecode"" = 'MS007' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            sCreditAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_REV_GL"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE A.""U_Filecode"" = 'MS007' AND A.""U_ActType"" = 'D'"
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
            dPayComp = 0

            Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F18")
            For k As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                    sSql = "SELECT IFNULL(SUM(""U_pay_comp""),0) AS ""U_pay_comp"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE ""U_invoice"" = '" & sInvoice & "' " & _
                           " AND IFNULL(""U_RevJrnlNo"",'') = '' AND IFNULL(""U_RevJournalEntry"",'') = '' AND ""U_status"" = 'O'"
                    oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(sSql)
                    If oRecordSet.RecordCount > 0 Then
                        dPayComp = oRecordSet.Fields.Item("U_pay_comp").Value
                        dTotvalue = dTotvalue + dPayComp
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                End If
            Next

            'For i As Integer = 0 To oDv.Count - 1
            '    dPayComp = CDbl(oDv(i)(20).ToString.Trim)
            '    dTotvalue = dTotvalue + dPayComp
            'Next

            If dTotvalue > 0 Then
                sClincCode = oDv(0)(4).ToString.Trim
                'sClincCode = p_oCompDef.sCOAcrlCardCode
                sCostCenter = oDv(0)(48).ToString.Trim
                sIncurMnth = oDv(0)(50).ToString.Trim

                Dim sApCode As String = oDv(0)(52).ToString.Trim()

                Dim sDocDate As String
                sDocDate = file.Name.Substring(12, 8)
                Dim dt As Date
                Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                oJournalEntry.TaxDate = dt
                oJournalEntry.ReferenceDate = dt
                If sCostCenter <> "" Then
                    oJournalEntry.Memo = "Reversal of Estimated cost for " & sClincCode & " " & sCostCenter
                Else
                    oJournalEntry.Memo = "Reversal of Estimated cost for " & sClincCode
                End If
                'oJournalEntry.Memo = "Reversal of Estimated cost for " & sApCode

                oJournalEntry.Lines.ShortName = sDebitAct
                oJournalEntry.Lines.Credit = dTotvalue
                If Not sCostCenter = String.Empty Then
                    oJournalEntry.Lines.CostingCode = sCostCenter
                    oJournalEntry.Lines.CostingCode2 = sCostCenter
                End If

                oJournalEntry.Lines.Add()

                oJournalEntry.Lines.AccountCode = sCreditAct
                oJournalEntry.Lines.Debit = dTotvalue
                If Not sCostCenter = String.Empty Then
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

                    Dim oRs As SAPbobsCOM.Recordset
                    Dim sQuery As String

                    sQuery = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & sTransId & "'"
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRs.DoQuery(sQuery)
                    If oRs.RecordCount > 0 Then
                        sJournalEntryNo = oRs.Fields.Item("Number").Value
                    End If

                    Console.WriteLine("Document Created Successfully :: " & sJournalEntryNo)

                    oDv.RowFilter = Nothing
                    oDtGroup = oDv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_status"" = 'C',""U_RevJrnlNo"" = '" & sJournalEntryNo & "',""U_RevJournalEntry"" = '" & sTransId & "' " & _
                                     " WHERE ""U_OcrCode"" = '" & sCostCenter & "' AND ""U_invoice"" = '" & sInvoice & "'" & _
                                     " AND ""U_incurred_month"" = '" & sIncurMnth & "' AND IFNULL(""U_RevJrnlNo"",'') = '' AND IFNULL(""U_RevJournalEntry"",'') = ''"
                            oRs.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateRevJouranl_CostAccrualJE = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message.ToString()
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateRevJouranl_CostAccrualJE = RTN_ERROR
        End Try
    End Function

End Module
