Module modCMSSHMMPD

    Private dtMBMSList As DataTable
    Private dtCardCode As DataTable
    Private dtItemCode As DataTable
    Private dtInvoice_ARDetails As DataTable
    Private dtInvoice_APDetails As DataTable

    Public Function ProcessHMMPDDatas(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessHMMPDDatas"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim sFile As String = file.Name
            Dim sFileName As String = Mid(sFile, 7)
            Dim sSource As String = String.Empty

            For Each c As Char In sFileName
                If Char.IsLetter(c) Then
                    If sSource = "" Then
                        sSource = c
                    Else
                        sSource = sSource & c
                    End If
                Else
                    Exit For
                End If
            Next

            sSql = "SELECT ""CardCode"",""VatGroup"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""OCRD"" "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtCardCode = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMMPDSAPDbName)

            sSql = "SELECT ""ItemCode"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""OITM"""
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtItemCode = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMMPDSAPDbName)

            sSql = "SELECT ""U_invoice"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AR_DETAILS"" WHERE IFNULL(""U_Inv_DocEntry"",'') <> ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtInvoice_ARDetails = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMMPDSAPDbName)

            sSql = "SELECT ""U_invoice"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AP_DETAILS"" WHERE IFNULL(""U_AP_Inv_DocEntry"",'') <> ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dtInvoice_APDetails = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sHMMPDSAPDbName)

            Dim odtDatatable As DataTable
            odtDatatable = oDv.ToTable
            odtDatatable.Columns.Add("IncuredMonth", GetType(Date))
            odtDatatable.Columns.Add("ArCode", GetType(String))
            odtDatatable.Columns.Add("ApCode", GetType(String))
            odtDatatable.Columns.Add("InvoiceDate", GetType(Date))
            odtDatatable.Columns.Add("CostCenter", GetType(String))

            Dim iLindex As Integer = file.Name.LastIndexOf("_")
            Dim sFileDate As String = file.Name.Substring(iLindex, (Len(file.Name) - iLindex))
            sFileDate = sFileDate.Replace("_", "")
            sFileDate = sFileDate.Replace(".xls", "")

            For intRow As Integer = 0 To odtDatatable.Rows.Count - 1
                If Not (odtDatatable.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or odtDatatable.Rows(intRow).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    Console.WriteLine("Processing excel line " & intRow)

                    Dim sCustomerType As String = String.Empty
                    sCustomerType = odtDatatable.Rows(intRow).Item(33).ToString

                    If sCustomerType.ToUpper() = "CONTRACT" Then
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
                        'Dim sSubCode As String = odtDatatable.Rows(intRow).Item(2).ToString
                        Dim sAPCode As String = "V" & sClinicCode '& sSubCode

                        dtCardCode.DefaultView.RowFilter = "CardCode = '" & sAPCode & "'"
                        If dtCardCode.DefaultView.Count = 0 Then
                            sErrDesc = "CardCode not found in SAP. Ap Code :: " & sAPCode
                            Console.WriteLine(sErrDesc)
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Throw New ArgumentException(sErrDesc)
                        End If

                        Dim iIndex As Integer = odtDatatable.Rows(intRow).Item(4).ToString.IndexOf(" ")
                        Dim sDate As String = odtDatatable.Rows(intRow).Item(4).ToString.Substring(0, iIndex)
                        Dim dt As Date
                        Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
                        Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

                        Dim sSchemeCode As String = odtDatatable.Rows(intRow).Item(26).ToString
                        Dim dIncurMnth As Date = CDate(dt.Date.AddDays(-(dt.Day - 1)).AddMonths(1).AddDays(-1).ToString())

                        Dim sCostCenter As String '= GetCostCenter(sCompCode, dt, sSchemeCode, p_oCompDef.sHMMPDSAPDbName)
                        'sSql = "SELECT TOP 1 ""U_MBMS"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_MBMS"" WHERE ""U_company_code"" = '" & sCompCode & "' AND ""U_Effective_Date"" <= '" & dt.ToString("yyyy-MM-dd") & "' " & _
                        '       " AND ""U_Scheme"" = '" & sSchemeCode & "' ORDER BY ""U_Effective_Date"" DESC "

                        sSql = "SELECT B.""PrcCode"" AS ""U_MBMS"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_MBMS"" A " & _
                               " INNER JOIN " & p_oCompDef.sHMMPDSAPDbName & ".""OPRC"" B ON UPPER(B.""PrcCode"") = UPPER(A.""U_MBMS"") " & _
                               " WHERE A.""U_company_code"" = '" & sCompCode & "' AND A.""U_Effective_Date"" <= '" & dt.ToString("yyyy-MM-dd") & "' " & _
                               " AND A.""U_Scheme"" = '" & sSchemeCode & "' ORDER BY A.""U_Effective_Date"" DESC "

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
                        sCostCenter = GetStringValue(sSql, p_oCompDef.sHMMPDSAPDbName)

                        Dim dInvoiceDate As Date
                        Date.TryParseExact(sFileDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dInvoiceDate)

                        If sSchemeCode = "" Then
                            sErrDesc = "Scheme Code is mandatory/Check in Excel line " & intRow
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Console.WriteLine(sErrDesc)
                            Throw New ArgumentException(sErrDesc)
                        End If

                        If sCostCenter = "" Then
                            sErrDesc = "MBMS column cannot be null / Check Cost Center for respective company code in config table/Check line " & intRow
                            Call WriteToLogFile(sErrDesc, sFuncName)
                            Console.WriteLine(sErrDesc)
                            Throw New ArgumentException(sErrDesc)
                        End If

                        odtDatatable.Rows(intRow)("F2") = sClinicCode.ToUpper()
                        odtDatatable.Rows(intRow)("F12") = sTreatment
                        odtDatatable.Rows(intRow)("IncuredMonth") = dIncurMnth
                        odtDatatable.Rows(intRow)("ArCode") = sArCode
                        odtDatatable.Rows(intRow)("ApCode") = sAPCode
                        odtDatatable.Rows(intRow)("InvoiceDate") = dInvoiceDate
                        odtDatatable.Rows(intRow)("CostCenter") = sCostCenter.ToUpper()
                    End If
                End If
            Next

            Dim oDvFinalView As DataView
            oDvFinalView = New DataView(odtDatatable)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToCompany(p_oCompany, p_oCompDef.sHMMPDSAPDbName, p_oCompDef.sHMMPDSAPUserName, p_oCompDef.sHMMPDSAPPassword, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then
                Console.WriteLine("Company connection Successful")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction", sFuncName)

                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If oDvFinalView.Count > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoTable_RD001_AR()", sFuncName)

                    Console.WriteLine("Inserting datas in YOT Table")
                    If InsertIntoTable_RD001_AR(oDvFinalView, p_oCompany, sSource, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Console.WriteLine("Insert into YOT table successful")

                    '**************PROCESSING DATAS WHICH LESS_DIS_PAY_CLIENT HAS AMOUNT***************
                    oDvFinalView.RowFilter = "F34 LIKE 'CONTRACT*' AND F16 <> '0'"
                    Dim odtpayClient As New DataTable
                    odtpayClient = oDvFinalView.ToTable

                    Dim oPayClientDv As DataView = New DataView(odtpayClient)

                    If oPayClientDv.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for AR invoice creation", sFuncName)

                        'F1 Invoice F2 Clinic Code, F25 Payment Method
                        Dim oDtGroup As DataTable = oPayClientDv.Table.DefaultView.ToTable(True, "F1", "F2", "CostCenter", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                                oPayClientDv.RowFilter = "F1='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F2 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                             " and CostCenter='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' and IncuredMonth ='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "'"
                                If oPayClientDv.Count > 0 Then
                                    Console.WriteLine("Inserting data into Cash table for " & oDtGroup.Rows(i).Item(1).ToString.Trim())
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoCashSales()", sFuncName)
                                    Dim oCashDt As DataTable
                                    oCashDt = oPayClientDv.ToTable
                                    Dim oCashDv As DataView = New DataView(oCashDt)
                                    If InsertIntoCashSales(oCashDv, p_oCompany, "LESS_DIS_PAY_CLIENT", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Inserting table into cash table is successful")
                                End If
                            End If
                        Next

                        oPayClientDv.RowFilter = Nothing

                        oDtGroup = oPayClientDv.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "ARCODE") Then
                                oPayClientDv.RowFilter = "F2 ='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                         " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "
                                If oPayClientDv.Count > 0 Then
                                    Dim oPClient_invDt As DataTable
                                    oPClient_invDt = oPayClientDv.ToTable
                                    Dim oPClient_InvDv As DataView = New DataView(oPClient_invDt)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoice()", sFuncName)
                                    If CreateARInvoice(oPClient_InvDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                            End If
                        Next
                    End If

                    '**************PROCESSING DATAS WHICH PAY_COMP HAS AMOUNT***************
                    oDvFinalView.RowFilter = Nothing
                    oDvFinalView.RowFilter = "F34 LIKE 'CONTRACT*' AND F14 <> '0'"
                    Dim odtPayComp As New DataTable
                    odtPayComp = oDvFinalView.ToTable

                    Dim oDvPayComp As DataView = New DataView(odtPayComp)
                    If oDvPayComp.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for insert into AP table", sFuncName)

                        oDvPayComp.RowFilter = Nothing
                        'F2 - Cln_Code
                        Dim oDtGroup As DataTable = oDvPayComp.Table.DefaultView.ToTable(True, "F1", "F2", "CostCenter", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                                oDvPayComp.RowFilter = "F1='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F2 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                       " and CostCenter = '" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "' "
                                If oDvPayComp.Count > 0 Then
                                    Console.WriteLine("Inserting values into AP Details table")
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoAPDetails()", sFuncName)
                                    Dim oDt As DataTable
                                    oDt = oDvPayComp.ToTable
                                    Dim oApInvDv As DataView = New DataView(oDt)
                                    If InsertIntoAPDetails(oApInvDv, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Insertion of values in AP Details table is successful")
                                End If
                            End If
                        Next

                        oDvPayComp.RowFilter = Nothing

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for AP invoice creation", sFuncName)

                        oDtGroup = oDvPayComp.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oDvPayComp.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' AND CostCenter = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                        " AND IncuredMonth = '" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "
                                If oDvPayComp.Count > 0 Then
                                    Dim oNewDtPayComp As DataTable
                                    oNewDtPayComp = oDvPayComp.ToTable
                                    Dim oNewDvPayComp As DataView = New DataView(oNewDtPayComp)
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateAPInvoice()", sFuncName)
                                    If CreateAPInvoice(oNewDvPayComp, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
            ProcessHMMPDDatas = RTN_SUCCESS
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
            ProcessHMMPDDatas = RTN_ERROR
        End Try
    End Function

    Private Function InsertIntoTable_RD001_AR(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal sSource As String, ByRef sErrDesc As String) As Long
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

                    sSql = "INSERT INTO " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_RD001_AR"" (""Code"",""Name"",""U_invoice"",""U_cln_code"",""U_subcode"",""U_cln_name"", " & _
                            " ""U_txn_date"",""U_id_type"",""U_id"",""U_lastname"",""U_given_name"",""U_christian"",""U_treat_code"",""U_treatment"",""U_cost"", " & _
                            " ""U_pay_comp"",""U_pay_client"",""U_les_dis_pay_client"",""U_admin"",""U_reimburse"",""U_cmoney"",""U_treat_charge"",""U_less_dis_treat_chg"",""U_surface"", " & _
                            " ""U_tooth_no"",""U_discount"",""U_paymethod"",""U_company"",""U_scheme"",""U_is_referral"",""U_Office_Invoice"",""U_date"", " & _
                            " ""U_amt"",""U_issued_by"",""U_is_refund"",""U_Customer_type"",""U_incurred_month"",""U_ar_code"",""U_ap_code"",""U_invoice_date"",""U_CostCenter"",""U_source"") " & _
                            " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sHMMPDSAPDbName & """.""@AE_RD001_AR""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sHMMPDSAPDbName & """.""@AE_RD001_AR""), " & _
                            " '" & oDv(i)(0).ToString.Trim & "','" & oDv(i)(1).ToString.Trim & "','" & oDv(i)(2).ToString.Trim & "','" & oDv(i)(3).ToString.Trim & "'," & _
                            " '" & oDv(i)(4).ToString.Trim & "','" & oDv(i)(5).ToString.Trim & "','" & oDv(i)(6).ToString.Trim & "','" & oDv(i)(7).ToString.Trim & "'," & _
                            " '" & oDv(i)(8).ToString.Trim & "','" & oDv(i)(9).ToString.Trim & "','" & oDv(i)(10).ToString.Trim & "','" & oDv(i)(11).ToString.Trim & "'," & _
                            " '" & oDv(i)(12).ToString.Trim & "','" & oDv(i)(13).ToString.Trim & "','" & oDv(i)(14).ToString.Trim & "','" & oDv(i)(15).ToString.Trim & "'," & _
                            " '" & oDv(i)(16).ToString.Trim & "','" & oDv(i)(17).ToString.Trim & "','" & oDv(i)(18).ToString.Trim & "','" & oDv(i)(19).ToString.Trim & "'," & _
                            " '" & oDv(i)(20).ToString.Trim & "','" & oDv(i)(21).ToString.Trim & "','" & oDv(i)(22).ToString.Trim & "','" & oDv(i)(23).ToString.Trim & "'," & _
                            " '" & oDv(i)(24).ToString.Trim & "','" & oDv(i)(25).ToString.Trim & "','" & oDv(i)(26).ToString.Trim & "','" & oDv(i)(27).ToString.Trim & "'," & _
                            " '" & oDv(i)(28).ToString.Trim & "','" & oDv(i)(29).ToString.Trim & "','" & oDv(i)(30).ToString.Trim & "','" & oDv(i)(31).ToString.Trim & "'," & _
                            " '" & oDv(i)(32).ToString.Trim & "','" & oDv(i)(33).ToString.Trim & "','" & oDv(i)(34).ToString.Trim & "','" & oDv(i)(35).ToString.Trim & "'," & _
                            " '" & oDv(i)(36).ToString.Trim & "','" & oDv(i)(37).ToString.Trim & "','" & oDv(i)(38).ToString.Trim & "', '" & sSource & "' )"
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
            sArCode = "C" & sClinicCode 'oDv(0)(35).ToString.Trim
            sIncuredMnth = oDv(0)(34).ToString.Trim

            sSQL = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_ITEMCODE"" " & _
                  " WHERE ""U_FileCode"" = 'YOT' AND UPPER(""U_Field"") = 'PAY_CLIENT' AND UPPER(""U_DocType"") = 'A/R' AND UPPER(""U_CustType"") = 'CASH' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            sItemCode = GetStringValue(sSQL, p_oCompDef.sHMMPDSAPDbName)

            sSQL = "SELECT ""U_BankGL"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_PAYMETHOD"" WHERE UPPER(""U_PayMethod"") = '" & sPayMethod.ToUpper & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            sBank = GetStringValue(sSQL, p_oCompDef.sHMMPDSAPDbName)

            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dInvoiceDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dInvoiceDate)

            Dim dTotal As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                sSQL = String.Empty
                Dim dAmount As Double

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

                dTotal = dTotal + dAmount
            Next
            sSQL = "INSERT INTO " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AR_DETAILS""(""Code"",""Name"",""U_incurred_month"", ""U_invoice_date"", " & _
                       " ""U_invoice"",""U_amount"",""U_cln_code"",""U_subcode"",""U_ItemCode"",""U_paymethod"",""U_bank"",""U_invoice_type"",""U_ar_code"",""U_CostCenter"") " & _
                       " VALUES ((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AR_DETAILS""), " & _
                       " (SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AR_DETAILS""), " & _
                       " '" & sIncuredMnth & "','" & dInvoiceDate.ToString("yyyy-MM-dd") & "','" & oDv(0)(0).ToString & "','" & dTotal & "', " & _
                       " '" & oDv(0)(1).ToString & "','" & oDv(0)(2).ToString & "','" & sItemCode & "','" & sPayMethod & "','" & sBank & "','Cash Sales','" & sArCode & "','" & oDv(0)(38).ToString & "') "

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

    Private Function CreateARInvoice(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateARInvoice"
        Dim sSql As String = String.Empty
        Dim sItemCode, sArCode, sCostCenter, sIncuredMnth, sVatGroup, sClinicCode As String
        Dim dAmount As Double = 0.0
        Dim iCount, iErrCode As Integer

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sArCode = "C" & oDv(0)(1).ToString.Trim()
            sCostCenter = oDv(0)(38).ToString.Trim()
            sIncuredMnth = oDv(0)(34).ToString.Trim()
            sClinicCode = oDv(0)(1).ToString.Trim()

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

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_ITEMCODE"" " & _
                   " WHERE UPPER(""U_FileCode"") = 'RD001' AND UPPER(""U_Field"") = 'LESS_DIS_PAY_CLIENT' AND UPPER(""U_DocType"") = 'A/R' AND UPPER(""U_CustType"") = 'CONTRACT' " & _
                   " AND ""U_ClinicCode"" = '" & sClinicCode & "' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMMPDSAPDbName)

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

            Dim dPercent As Double
            sSql = "SELECT ""U_Percentage"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_PERCENTAGE"" WHERE UPPER(""U_Type"") = 'A/R'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dPercent = getAmt(sSql, p_oCompDef.sHMMPDSAPDbName, p_oCompDef.sHMMPDSAPUserName, p_oCompDef.sHMMPDSAPPassword)

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

            dTotal = dTotal * (dPercent / 100)

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
                If Not (sCostCenter = String.Empty) Then
                    oARInvoice.Lines.CostingCode = sCostCenter
                    oARInvoice.Lines.COGSCostingCode = sCostCenter
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

                    sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""OINV"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSql)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AR_DETAILS"" SET ""U_Inv_DocNo"" = '" & iDocNo & "',""U_Inv_DocEntry"" = '" & iDocEntry & "'" & _
                                     " WHERE ""U_ar_code"" = '" & sArCode & "' AND ""U_CostCenter"" = '" & sCostCenter & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' " & _
                                     " AND ""U_invoice"" = '" & sInvoice & "'"

                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateARInvoice = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateARInvoice = RTN_ERROR
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

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_ITEMCODE"" WHERE ""U_FileCode"" = 'YOT'	AND UPPER(""U_Field"") = 'PAY_CLIENT' " & _
                   " AND UPPER(""U_DocType"") = 'A/P' AND UPPER(""U_CustType"") = 'CONTRACT'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMMPDSAPDbName)

            'sSql = "SELECT ""U_CardCode"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_CONTRACT_OWNER"" WHERE UPPER(""U_Type"") = 'A/P'"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            'sApCode = GetStringValue(sSql, p_oCompDef.sHMMPDSAPDbName)

            sApCode = "V" & oDv(0)(1).ToString.Trim '& oDv(0)(2).ToString.Trim()

            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim dTotal As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                sSql = String.Empty
                Dim dAmount As Double
                Try
                    If Not (oDv(i)(13).ToString.Trim = String.Empty) Then
                        dAmount = CDbl(oDv(i)(13).ToString.Trim)
                    End If
                Catch ex As Exception
                    dAmount = 0.0
                End Try

                dTotal = dTotal + dAmount
            Next

            Dim sInvoiceDate As String = oDv(0)(37).ToString.Trim()
            Dim iIndex As Integer = sInvoiceDate.IndexOf(" ")
            Dim sInvoiceDt_Trimed As String = sInvoiceDate.Substring(0, iIndex)
            Dim dInvoiceDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sInvoiceDt_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dInvoiceDate)

            sSql = "INSERT INTO " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AP_DETAILS""(""Code"",""Name"",""U_company_code"",""U_incurred_month"",""U_invoice_date"",""U_invoice"", " & _
                      " ""U_amount"",""U_cln_code"",""U_subcode"",""U_ItemCode"",""U_invoice_type"",""U_ap_code"",""U_CostCenter"") " & _
                      " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AP_DETAILS""), " & _
                      " (SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AP_DETAILS""), " & _
                      " '" & oDv(0)(25).ToString & "','" & oDv(0)(34).ToString & "', '" & dInvoiceDate.ToString("yyyy-MM-dd") & "','" & oDv(0)(0).ToString & "','" & dTotal & "','" & oDv(0)(1).ToString & "', " & _
                      " '" & oDv(0)(2).ToString & "','" & sItemCode & "','CONTRACT','" & sApCode & "','" & oDv(0)(38).ToString & "' ) "

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

    Private Function CreateAPInvoice(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateAPInvoice"
        Dim sSql, sItemCode, sAPCode, sClinicCode, sIncuredMnth, sVatGroup, sCostCenter As String
        Dim dAmount As Double = 0.0
        Dim iCount, iErrCode As Integer

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sClinicCode = oDv(0)(1).ToString.Trim()
            sAPCode = "V" & sClinicCode '& oDv(0)(2).ToString.Trim()
            sIncuredMnth = oDv(0)(34).ToString.Trim()
            sCostCenter = oDv(0)(38).ToString.Trim()

            dtCardCode.DefaultView.RowFilter = "CardCode = '" & sAPCode & "'"
            If dtCardCode.DefaultView.Count = 0 Then
                sErrDesc = "CardCode not found in SAP/CardCode :: " & sAPCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else
                sAPCode = dtCardCode.DefaultView.Item(0)(0).ToString().Trim()
                sVatGroup = dtCardCode.DefaultView.Item(0)(1).ToString().Trim()
            End If

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim sIncruMnth_Trimed As String = sIncuredMnth.Substring(0, iIndex)
            Dim dDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sIncruMnth_Trimed, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dDocDate)

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_ITEMCODE"" " & _
                   " WHERE UPPER(""U_FileCode"") = 'RD001' AND UPPER(""U_Field"") = 'LESS_DIS_PAY_CLIENT' AND UPPER(""U_DocType"") = 'A/P' AND UPPER(""U_CustType"") = 'CONTRACT' " & _
                   " AND ""U_ClinicCode"" = '" & sClinicCode & "' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sHMMPDSAPDbName)

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

            Dim dPercent As Double
            sSql = "SELECT ""U_Percentage"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_PERCENTAGE"" WHERE UPPER(""U_Type"") = 'A/P'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            dPercent = getAmt(sSql, p_oCompDef.sHMMPDSAPDbName, p_oCompDef.sHMMPDSAPUserName, p_oCompDef.sHMMPDSAPPassword)

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

            dTotal = dTotal * (dPercent / 100)

            If dTotal > 0 Then
                Dim oApInvoice As SAPbobsCOM.Documents
                oApInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

                oApInvoice.CardCode = sAPCode
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
                    oApInvoice.Lines.CostingCode = sCostCenter
                    oApInvoice.Lines.COGSCostingCode = sCostCenter
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

                    sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sHMMPDSAPDbName & ".""OPCH"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSql)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F1")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sHMMPDSAPDbName & ".""@AE_AP_DETAILS"" SET ""U_AP_Inv_DocNo"" = '" & iDocNo & "',""U_AP_Inv_DocEntry"" = '" & iDocEntry & "'" & _
                                     " WHERE ""U_cln_code"" = '" & sClinicCode & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' AND ""U_invoice"" = '" & sInvoice & "'"

                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateAPInvoice = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateAPInvoice = RTN_ERROR
        End Try

    End Function

End Module
