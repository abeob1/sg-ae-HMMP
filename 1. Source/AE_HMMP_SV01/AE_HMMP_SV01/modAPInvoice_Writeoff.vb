﻿Module modAPInvoice_Writeoff

    Private dtSoAcrlInvList As DataTable
    Private dtCostAcrlInvList As DataTable
    Private dtFileDate As Date

    Public Function ProcessAPInvoice_WriteOff(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessAPInvoice_WriteOff"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim sFileDate As String = file.Name.Substring(9, 8)
            Dim dformat() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sFileDate, dformat, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dtFileDate)

            sSql = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE IFNULL(""U_status"",'O') = 'C' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtSoAcrlInvList = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_status"",'O') = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtCostAcrlInvList = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

            Dim odtDatatable As DataTable
            odtDatatable = oDv.ToTable

            odtDatatable.Columns.Add("CostCenter", GetType(String))
            odtDatatable.Columns.Add("Insurer", GetType(String))
            odtDatatable.Columns.Add("IncuredMonth", GetType(Date))
            odtDatatable.Columns.Add("Type", GetType(String))
            odtDatatable.Columns.Add("AcrlType", GetType(String))

            For intRow As Integer = 0 To odtDatatable.Rows.Count - 1
                If Not (odtDatatable.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or odtDatatable.Rows(intRow).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    Console.WriteLine("Processing excel line " & intRow & " to get MBMS and Insurer from config table")

                    Dim sCliniCode As String = odtDatatable.Rows(intRow).Item(1).ToString
                    Dim sCompCode As String = odtDatatable.Rows(intRow).Item(6).ToString
                    Dim sCompName As String = odtDatatable.Rows(intRow).Item(5).ToString
                    sCompName = sCompName.Replace("'", " ")
                    Dim sSchemeCode As String = odtDatatable.Rows(intRow).Item(7).ToString

                    If sCompCode = "" Then
                        sErrDesc = "Company Code should not be empty / Check Line " & intRow
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine(sErrDesc)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Dim sType, sArCode As String
                    sArCode = "C" & sCompCode
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting type for the cardcode " & sArCode, sFuncName)
                    sSql = "SELECT ""U_Type"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" WHERE ""CardCode"" = '" & sArCode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                    sType = GetStringValue(sSql, p_oCompDef.sSAPDBName)

                    If sType = "" Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Type is empty.Taking default value from config file", sFuncName)
                        sType = p_oCompDef.sType
                    End If

                    Dim sInvoice As String = odtDatatable.Rows(intRow).Item(0).ToString.Trim
                    dtSoAcrlInvList.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtSoAcrlInvList.DefaultView.Count > 0 Then
                        sErrDesc = "Invoice already closed. invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    dtCostAcrlInvList.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
                    If dtCostAcrlInvList.DefaultView.Count > 0 Then
                        sErrDesc = "Invoice already closed. invoice no :: " & sInvoice
                        Console.WriteLine(sErrDesc)
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If

                    Dim sNewType As String = String.Empty
                    sSql = "SELECT ""U_Type"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE ""U_invoice"" = '" & sInvoice & "' "
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                    sNewType = GetStringValue(sSql, p_oCompDef.sSAPDBName)

                    If sNewType = "" Then
                        sSql = "SELECT ""U_Type"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE ""U_invoice"" = '" & sInvoice & "' "
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                        sNewType = GetStringValue(sSql, p_oCompDef.sSAPDBName)
                    End If

                    Dim iIndex As Integer = odtDatatable.Rows(intRow).Item(4).ToString.IndexOf(" ")
                    Dim sDate As String = odtDatatable.Rows(intRow).Item(4).ToString.Substring(0, iIndex)
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

                    odtDatatable.Rows(intRow)("F2") = sCliniCode.ToUpper()
                    odtDatatable.Rows(intRow)("F6") = sCompName
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

                    Console.WriteLine("Inserting datas into PO Table")
                    If InsertIntoAPTable_WriteOff(oDvFinalView, file.Name, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Console.WriteLine("Data insert into PO Table Successful")

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas based on Type column - Non Capitation", sFuncName)

                    oDvFinalView.RowFilter = "AcrlType NOT LIKE 'CAPITATION*'"
                    Dim odtNonCap As New DataTable
                    odtNonCap = oDvFinalView.ToTable

                    Dim oNonCapDv As DataView = New DataView(odtNonCap)

                    If oNonCapDv.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas to create reverse journal entry Non Capitation", sFuncName)
                        Dim oDtGroup As DataTable = oNonCapDv.Table.DefaultView.ToTable(True, "F2", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oNonCapDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "

                                If oNonCapDv.Count > 0 Then
                                    Console.WriteLine("Creating Reverse journal - Non Capitation for clinic code " & oDtGroup.Rows(i).Item(0).ToString.Trim())
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateReverseJournal_CostAccrual", sFuncName)
                                    Dim dtCostAcrlDatas As DataTable
                                    dtCostAcrlDatas = oNonCapDv.ToTable
                                    Dim oDVCostAcrlDatas As DataView = New DataView(dtCostAcrlDatas)
                                    If CreateReverseJournal_CostAccrual_NonCap(p_oCompany, oDVCostAcrlDatas, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Reverse Journal created for clinic code " & oDtGroup.Rows(i).Item(0).ToString.Trim())
                                End If

                            End If
                        Next
                    End If

                    oDvFinalView.RowFilter = Nothing
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas based on Type column - Capitation", sFuncName)

                    oDvFinalView.RowFilter = "AcrlType LIKE 'CAPITATION*'"
                    Dim odtCap As New DataTable
                    odtCap = oDvFinalView.ToTable

                    Dim oCapDv As DataView = New DataView(odtCap)
                    If oCapDv.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas to create reverse journal entry Capitation", sFuncName)

                        Dim oDtGroup As DataTable = oCapDv.Table.DefaultView.ToTable(True, "F2", "IncuredMonth")
                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oCapDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and IncuredMonth = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "

                                If oCapDv.Count > 0 Then
                                    Console.WriteLine("Creating Reverse journal - Capitation for Clinic Code " & oDtGroup.Rows(i).Item(0).ToString.Trim())
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateReverseJournal_CostAccrual_Capitation", sFuncName)
                                    Dim dtCostAcrlDatas As DataTable
                                    dtCostAcrlDatas = oCapDv.ToTable
                                    Dim oDVCostAcrlDatas As DataView = New DataView(dtCostAcrlDatas)
                                    If CreateReverseJournal_CostAccrual_Capitation(p_oCompany, oDVCostAcrlDatas, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Console.WriteLine("Reverse Journal created for Clinic code " & oDtGroup.Rows(i).Item(0).ToString.Trim())
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
            ProcessAPInvoice_WriteOff = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessAPInvoice_WriteOff = RTN_ERROR
        End Try
    End Function

    Private Function InsertIntoAPTable_WriteOff(ByVal oDv As DataView, ByVal sFileName As String, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "InsertIntoAPTable_WriteOff"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'LINE BY LINE
            For i As Integer = 1 To oDv.Count - 1
                If Not (oDv(i)(0).ToString.Trim = String.Empty) Then
                    Console.WriteLine("Inserting Line Num : " & i)
                    sSql = String.Empty

                    Dim sCompCode As String = oDv(i)(6).ToString.Trim
                    If sCompCode <> "" Then
                        sCompCode = "C" & sCompCode
                    End If
                    Dim sClinicCode As String = oDv(i)(1).ToString.Trim
                    If sClinicCode <> "" Then
                        sClinicCode = "Z" & sClinicCode
                    End If


                    sSql = "INSERT INTO " & p_oCompDef.sSAPDBName & ".""@AE_MS002_APWRITEOF"" (""Code"",""Name"",""U_invoice"",""U_cln_code"",""U_subcode"",""U_cln_name"",""U_txn_date""," & _
                            " ""U_company"",""U_company_code"",""U_scheme_code"",""U_m_id_type"",""U_m_id"",""U_id_type"",""U_id"",""U_treat_code"", " & _
                            " ""U_treatment"",""U_charge"",""U_pay_comp"",""U_pay_client"",""U_oper"",""U_ds"",""U_reimburse"",""U_cmoney"",""U_diag_desc"", " & _
                            " ""U_refer_from_name"",""U_lastname"",""U_given_name"",""U_christian"",""U_remark_fg"",""U_manualfee"",""U_in_time"",""U_status"", " & _
                            " ""U_sl_fr"",""U_sl_to"",""U_txn_remark_type"",""U_txn_remark"",""U_txn_remark_userid"",""U_create_datetime"",""U_create_userid"", " & _
                            " ""U_OcrCode"",""U_Insurer"",""U_incurred_month"",""U_ar_code"",""U_ap_code"",""U_Type"",""U_FileName"" ) " & _
                            " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS002_APWRITEOF""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS002_APWRITEOF"")," & _
                            " '" & oDv(i)(0).ToString.Trim & "','" & oDv(i)(1).ToString.Trim & "','" & oDv(i)(2).ToString.Trim & "','" & oDv(i)(3).ToString.Trim & "'," & _
                            " '" & oDv(i)(4).ToString.Trim & "','" & oDv(i)(5).ToString.Trim & "','" & oDv(i)(6).ToString.Trim & "','" & oDv(i)(7).ToString.Trim & "'," & _
                            " '" & oDv(i)(8).ToString.Trim & "','" & oDv(i)(9).ToString.Trim & "','" & oDv(i)(10).ToString.Trim & "','" & oDv(i)(11).ToString.Trim & "'," & _
                            " '" & oDv(i)(12).ToString.Trim & "','" & oDv(i)(13).ToString.Trim & "','" & oDv(i)(14).ToString.Trim & "','" & oDv(i)(15).ToString.Trim & "'," & _
                            " '" & oDv(i)(16).ToString.Trim & "','" & oDv(i)(17).ToString.Trim & "','" & oDv(i)(18).ToString.Trim & "','" & oDv(i)(19).ToString.Trim & "'," & _
                            " '" & oDv(i)(20).ToString.Trim & "','" & oDv(i)(21).ToString.Trim & "','" & oDv(i)(22).ToString.Trim & "','" & oDv(i)(23).ToString.Trim & "'," & _
                            " '" & oDv(i)(24).ToString.Trim & "','" & oDv(i)(25).ToString.Trim & "','" & oDv(i)(26).ToString.Trim & "','" & oDv(i)(27).ToString.Trim & "'," & _
                            " '" & oDv(i)(28).ToString.Trim & "','" & oDv(i)(29).ToString.Trim & "','" & oDv(i)(30).ToString.Trim & "','" & oDv(i)(31).ToString.Trim & "'," & _
                            " '" & oDv(i)(32).ToString.Trim & "','" & oDv(i)(33).ToString.Trim & "','" & oDv(i)(34).ToString.Trim & "','" & oDv(i)(35).ToString.Trim & "'," & _
                            " '" & oDv(i)(36).ToString.Trim & "','" & oDv(i)(37).ToString.Trim & "','" & oDv(i)(38).ToString.Trim & "','" & oDv(i)(39).ToString.Trim & "'," & _
                            " '" & sCompCode & "','" & sClinicCode & "',(SELECT ""U_Type"" FROM ""OCRD"" WHERE ""CardCode"" = '" & sCompCode & "'),'" & sFileName & "')"


                    oRecSet.DoQuery(sSql)
                End If
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoAPTable_WriteOff = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoAPTable_WriteOff = RTN_ERROR
        End Try
    End Function

    Private Function CreateReverseJournal_CostAccrual_NonCap(ByVal oCompany As SAPbobsCOM.Company, ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateReverseJournal_CostAccrual_NonCap"
        Dim sSql, sQuery As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCostCenter As String = String.Empty
        Dim sClinicCod As String = String.Empty
        Dim dCmoney, dPayClient, dOper As Double
        Dim sOperAct, sCMoneyAct, sPayClntAct, sActCode As String
        Dim iErrCode, iCount As Integer
        Dim bIsLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim sInvoice As String = oDv(0)(0).ToString.Trim
            Dim sSubCode As String = oDv(0)(2).ToString.Trim
            sClinicCod = oDv(0)(1).ToString.Trim
            Dim sIncurMonth As String = oDv(0)(39).ToString.Trim
            Dim sCompCode As String = oDv(0)(6).ToString.Trim
            Dim sDate As String = file.Name.Substring(9, 8)
            Dim dt As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                     " WHERE UPPER(A.""U_Field"") = 'CMONEY' AND IFNULL(A.""U_Type"",'') = ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            sCMoneyAct = GetActCode(sQuery, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                     " WHERE UPPER(A.""U_Field"") = 'OPER' AND IFNULL(A.""U_Type"",'') = ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            sOperAct = GetActCode(sQuery, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                     " WHERE UPPER(A.""U_Field"") = 'PAY_CLIENT' AND IFNULL(A.""U_Type"",'') = ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            sPayClntAct = GetActCode(sQuery, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                     " WHERE IFNULL(A.""U_Field"",'') = '' AND IFNULL(A.""U_Type"",'') = ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            sActCode = GetActCode(sQuery, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

            If sCMoneyAct = "" Then
                sErrDesc = "Account code for CMoney column cannot be null/Check the account code in config table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sOperAct = "" Then
                sErrDesc = "Account code for Oper column cannot be null/Check the account code in config table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sPayClntAct = "" Then
                sErrDesc = "Account code for Pay_Client column cannot be null/Check the account code in config table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sActCode = "" Then
                sErrDesc = "Account code cannot be null/Check the account code in config table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            oJournalEntry.TaxDate = dt
            oJournalEntry.ReferenceDate = dt
            oJournalEntry.Memo = "Reversal of Estimated cost for " & sClinicCod

            iCount = 1

            Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "CostCenter")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas based on MBMS for creating journal entry", sFuncName)
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COSTCENTER") Then
                    oDv.RowFilter = "CostCenter = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "'"

                    If oDv.Count > 0 Then

                        sCostCenter = String.Empty
                        sCostCenter = oDv(0)(37).ToString.Trim

                        dCmoney = 0
                        dPayClient = 0
                        dOper = 0

                        Dim oNewDt As DataTable = oDv.ToTable
                        Dim oNewDv As DataView = New DataView(oNewDt)

                        If oNewDv.Count > 0 Then
                            Dim oDtGroup_New As DataTable = oNewDv.Table.DefaultView.ToTable(True, "F1")
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting values for respective invoice", sFuncName)
                            For k As Integer = 0 To oDtGroup_New.Rows.Count - 1
                                If Not (oDtGroup_New.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup_New.Rows(k).Item(0).ToString.ToUpper.Trim() = "INVOICE") Then
                                    sInvoice = oDtGroup_New.Rows(k).Item(0).ToString.Trim()

                                    sSql = "SELECT COUNT(""U_invoice"") AS ""MNO"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE ""U_invoice"" = '" & sInvoice & "' AND IFNULL(""U_RevJournalEntry"",'') = ''"
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                                    Dim iInvCount As Integer = GetCode(sSql, p_oCompDef.sSAPDBName)

                                    If iInvCount > 0 Then
                                        sSql = "SELECT SUM(""U_cmoney"") ""U_cmoney"",SUM(""U_pay_client"") ""U_pay_client"" ,SUM(""U_oper"") ""U_oper""  " & _
                                           " FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" " & _
                                           " WHERE ""U_invoice"" = '" & sInvoice & "' AND ""U_OcrCode"" = '" & sCostCenter & "' " & _
                                           " AND ""U_incurred_month"" = '" & sIncurMonth & "' AND IFNULL(""U_RevJournalEntry"",'') = '' "
                                        oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery(sSql)
                                        If oRecordSet.RecordCount > 0 Then
                                            dCmoney = dCmoney + oRecordSet.Fields.Item("U_cmoney").Value
                                            dPayClient = dPayClient + oRecordSet.Fields.Item("U_pay_client").Value
                                            dOper = dOper + oRecordSet.Fields.Item("U_oper").Value
                                        End If
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                                    Else
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice " & sInvoice & " not found in the cost accrual table", sFuncName)
                                    End If
                                End If
                            Next

                            dCmoney = Math.Round(dCmoney, 2)
                            dPayClient = Math.Round(dPayClient, 2)
                            dOper = Math.Round(dOper, 2)

                            If dCmoney <> 0 Then
                                If iCount > 1 Then
                                    oJournalEntry.Lines.Add()
                                End If
                                oJournalEntry.Lines.AccountCode = sCMoneyAct
                                oJournalEntry.Lines.Credit = dCmoney
                                If Not sCostCenter = String.Empty Then
                                    oJournalEntry.Lines.CostingCode = sCostCenter
                                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                                End If
                                iCount = iCount + 1
                                bIsLineAdded = True
                            End If
                            If dOper <> 0 Then
                                If iCount > 1 Then
                                    oJournalEntry.Lines.Add()
                                End If
                                oJournalEntry.Lines.ShortName = sOperAct
                                oJournalEntry.Lines.Debit = dOper
                                If Not sCostCenter = String.Empty Then
                                    oJournalEntry.Lines.CostingCode = sCostCenter
                                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                                End If
                                iCount = iCount + 1
                                bIsLineAdded = True
                            End If
                            If dPayClient <> 0 Then
                                If iCount > 1 Then
                                    oJournalEntry.Lines.Add()
                                End If
                                oJournalEntry.Lines.ShortName = sPayClntAct
                                oJournalEntry.Lines.Debit = dPayClient
                                If Not sCostCenter = String.Empty Then
                                    oJournalEntry.Lines.CostingCode = sCostCenter
                                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                                End If
                                iCount = iCount + 1
                                bIsLineAdded = True
                            End If

                            Dim dTotval As Double
                            dTotval = (dCmoney - dOper - dPayClient)
                            dTotval = Math.Round(dTotval, 2)

                            If dTotval <> 0 Then
                                If iCount > 1 Then
                                    oJournalEntry.Lines.Add()
                                End If
                                oJournalEntry.Lines.ShortName = sActCode
                                oJournalEntry.Lines.Debit = dTotval
                                If Not sCostCenter = String.Empty Then
                                    oJournalEntry.Lines.CostingCode = sCostCenter
                                    'oJournalEntry.Lines.CostingCode2 = sCostCenter
                                End If
                                iCount = iCount + 1
                                bIsLineAdded = True
                            End If

                        End If
                    End If
                End If
            Next

            If bIsLineAdded = True Then
                If oJournalEntry.Add() <> 0 Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)
                    oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iJournalEntryNo, iDocNo As Integer
                    p_oCompany.GetNewObjectCode(iJournalEntryNo)

                    sSql = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & iJournalEntryNo & "'"
                    oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(sSql)
                    If oRecordSet.RecordCount > 0 Then
                        iDocNo = oRecordSet.Fields.Item("Number").Value
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                    Console.WriteLine("Document Created Successfully :: " & iDocNo)

                    Dim sXcelInvNo As String
                    Dim oRs As SAPbobsCOM.Recordset
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    oDv.RowFilter = Nothing

                    oDtGroup = oDv.Table.DefaultView.ToTable(True, "F1")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            sXcelInvNo = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_status"" = 'C',""U_RevJournalEntry"" = '" & iJournalEntryNo & "',""U_RevJrnlNo"" = '" & iDocNo & "' " & _
                                     " WHERE ""U_source"" = 'MS002' AND ""U_invoice"" = '" & sXcelInvNo & "' AND IFNULL(""U_RevJournalEntry"",'') = '' "
                            oRs.DoQuery(sQuery)

                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateReverseJournal_CostAccrual_NonCap = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateReverseJournal_CostAccrual_NonCap = RTN_ERROR
        End Try
    End Function

    Private Function CreateReverseJournal_CostAccrual_Capitation(ByVal oCompany As SAPbobsCOM.Company, ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateReverseJournal_CostAccrual_Capitation"
        Dim sSql, sQuery As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCostCenter As String = String.Empty
        Dim sIncuredMnth As String = String.Empty
        Dim sClinicCod As String = String.Empty
        Dim dCmoney, dPayClient, dOper, dCMoneyClient As Double
        Dim sOperAct, sPayClntAct_Debit, sPayClnt_Credit, sActCode, sType, sCMoneyClient As String
        Dim iErrCode, iCount As Integer
        Dim bIsLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim sSubCode As String = oDv(0)(2).ToString.Trim
            sClinicCod = oDv(0)(1).ToString.Trim
            sIncuredMnth = oDv(0)(39).ToString.Trim

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim dt As Date = CDate(sIncuredMnth.Substring(0, iIndex))

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                    " WHERE UPPER(A.""U_Field"") = 'CMONEY-PAY_CLIENT' AND UPPER(A.""U_Type"") = 'CAPITATION' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            sCMoneyClient = GetActCode(sQuery, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                     " WHERE UPPER(A.""U_Field"") = 'OPER' AND UPPER(IFNULL(A.""U_Type"",'')) = 'CAPITATION'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            sOperAct = GetActCode(sQuery, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                     " WHERE UPPER(A.""U_Field"") = 'PAY_CLIENT' AND UPPER(IFNULL(A.""U_Type"",'')) = 'CAPITATION' AND A.""U_ActType"" = 'D'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            sPayClntAct_Debit = GetActCode(sQuery, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                     " WHERE UPPER(A.""U_Field"") = 'PAY_CLIENT' AND UPPER(IFNULL(A.""U_Type"",'')) = 'CAPITATION' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            sPayClnt_Credit = GetActCode(sQuery, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                     " WHERE IFNULL(A.""U_Field"",'') = '' AND IFNULL(A.""U_Type"",'') = ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sQuery, sFuncName)
            sActCode = GetActCode(sQuery, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

            If sCMoneyClient = "" Then
                sErrDesc = "Account code for Cmoney-PayClient column cannot be null/Check the account code in config table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sOperAct = "" Then
                sErrDesc = "Account code for Oper column cannot be null/Check the account code in config table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sPayClntAct_Debit = "" Then
                sErrDesc = "Account code for Pay_Client column for debit cannot be null/Check the account code in config table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sPayClnt_Credit = "" Then
                sErrDesc = "Account code for Pay_Client column for Credit cannot be null/Check the account code in config table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If
            If sActCode = "" Then
                sErrDesc = "Account code cannot be null/Check the account code in config table"
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            oJournalEntry.TaxDate = dt
            oJournalEntry.ReferenceDate = dt
            oJournalEntry.Memo = "Reversal of Estimated cost for " & sClinicCod

            iCount = 1

            Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "CostCenter")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas based on MBMS for creating journal entry", sFuncName)
            For i As Integer = 0 To oDtGroup.Rows.Count - 1
                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                    oDv.RowFilter = "CostCenter = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "'"

                    If oDv.Count > 0 Then

                        sCostCenter = String.Empty
                        sCostCenter = oDv(0)(37).ToString.Trim

                        dCmoney = 0
                        dPayClient = 0
                        dOper = 0

                        Dim oNewDt As DataTable = oDv.ToTable
                        Dim oNewDv As DataView = New DataView(oNewDt)

                        If oNewDv.Count > 0 Then
                            Dim oDtGroup_New As DataTable = oNewDv.Table.DefaultView.ToTable(True, "F1")
                            For k As Integer = 0 To oDtGroup_New.Rows.Count - 1
                                If Not (oDtGroup_New.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup_New.Rows(k).ToString.ToUpper().Trim() = "INVOICE") Then
                                    Dim sInvoice As String = oDtGroup_New.Rows(k).Item(0).ToString.Trim()

                                    sSql = "SELECT COUNT(""U_invoice"") AS ""MNO"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE UPPER(""U_invoice"") = '" & sInvoice.ToUpper() & "'"
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                                    Dim iInvCount As Integer = GetCode(sSql, p_oCompDef.sSAPDBName)

                                    If iInvCount > 0 Then
                                        sSql = "SELECT SUM(""U_cmoney"") ""U_cmoney"",SUM(""U_pay_client"") ""U_pay_client"" ,SUM(""U_oper"") ""U_oper""  " & _
                                           " FROM ""@AE_COSTACCRUAL"" " & _
                                           " WHERE ""U_invoice"" = '" & sInvoice & "' AND ""U_OcrCode"" = '" & sCostCenter & "' " & _
                                           " AND ""U_incurred_month"" = '" & sIncuredMnth & "' AND IFNULL(""U_RevJournalEntry"",'') = '' "
                                        oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery(sSql)
                                        If oRecordSet.RecordCount > 0 Then
                                            dCmoney = dCmoney + oRecordSet.Fields.Item("U_cmoney").Value
                                            dPayClient = dPayClient + oRecordSet.Fields.Item("U_pay_client").Value
                                            dOper = dOper + oRecordSet.Fields.Item("U_oper").Value
                                        End If
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                                    Else
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice " & sInvoice & " not found in the cost accrual table", sFuncName)
                                    End If
                                End If
                            Next
                        End If

                        dCmoney = Math.Round(dCmoney, 2)
                        dPayClient = Math.Round(dPayClient, 2)
                        dOper = Math.Round(dOper, 2)

                        dCMoneyClient = (dCmoney - dPayClient)
                        dCMoneyClient = Math.Round(dCMoneyClient, 2)

                        If dCMoneyClient <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.AccountCode = sCMoneyClient
                            oJournalEntry.Lines.Credit = dCMoneyClient
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                'oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If
                        If dPayClient <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.AccountCode = sPayClntAct_Debit
                            oJournalEntry.Lines.Credit = dPayClient
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                'oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If
                        If dPayClient <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.ShortName = sPayClnt_Credit
                            oJournalEntry.Lines.Debit = dPayClient
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                'oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If
                        If dOper <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.ShortName = sOperAct
                            oJournalEntry.Lines.Debit = dOper
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                'oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If


                        Dim dTotval As Double
                        dTotval = ((dCMoneyClient + dPayClient) - (dOper + dPayClient))
                        dTotval = Math.Round(dTotval, 2)

                        If dTotval <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.ShortName = sActCode
                            oJournalEntry.Lines.Debit = dTotval
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                ' oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If

                    End If
                End If
            Next

            If bIsLineAdded = True Then
                If oJournalEntry.Add() <> 0 Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)
                    oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iJournalEntryNo, iDocNo As Integer
                    p_oCompany.GetNewObjectCode(iJournalEntryNo)

                    sSql = "SELECT ""Number"" FROM " & p_oCompDef.sSAPDBName & ".""OJDT"" WHERE ""TransId"" = '" & iJournalEntryNo & "'"
                    oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery(sSql)
                    If oRecordSet.RecordCount > 0 Then
                        iDocNo = oRecordSet.Fields.Item("Number").Value
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)

                    Console.WriteLine("Document Created Successfully :: " & iDocNo)

                    Dim oRs As SAPbobsCOM.Recordset
                    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    oDv.RowFilter = Nothing

                    oDtGroup = oDv.Table.DefaultView.ToTable(True, "F1")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_status"" = 'C',""U_RevJrnlNo"" = '" & iDocNo & "', ""U_RevJournalEntry"" = '" & iJournalEntryNo & "' " & _
                                     " WHERE ""U_source"" = 'MS002' AND IFNULL(""U_RevJournalEntry"",'') = '' AND ""U_invoice"" = '" & sInvoice & "' "

                            oRs.DoQuery(sQuery)

                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateReverseJournal_CostAccrual_Capitation = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateReverseJournal_CostAccrual_Capitation = RTN_ERROR
        End Try
    End Function

    Private Function GetActCode(ByVal sSql As String, ByVal sDBName As String, ByVal sUser As String, ByVal sPwd As String) As String
        Dim sFuncName As String = "GetActCode"
        Dim oDs As DataSet
        Dim sActCode As String = String.Empty

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        oDs = ExecuteSQLQuery_Hana(sSql, sDBName)

        If oDs.Tables(0).Rows.Count > 0 Then
            sActCode = oDs.Tables(0).Rows(0).Item(0).ToString
        End If

        Return sActCode
    End Function

End Module
