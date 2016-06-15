Module modARInvoice_Cardcode

    Private dtSoAcrlInvList As DataTable
    Private dtCostAcrlInvList As DataTable
    Private dtInsurerList As DataTable
    Private dtMBMSList As DataTable
    Private dtItemCode As DataTable
    Private dtCardCode As DataTable

    Public Function ProcessARInvoice_CardCode(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessARInvoice_CardCode"
        Dim sSQL As String

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            'sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE ""U_status"" = 'C' "
            sSQL = "SELECT DISTINCT ""U_invoice"" FROM( " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE IFNULL(""U_ARINV_DocEntry"",'') <> '' UNION ALL " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE IFNULL(""U_RevJournalEntry"",'') <> '' )T1 "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
            dtSoAcrlInvList = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

            'sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE ""U_status"" = 'C' "
            sSQL = "SELECT DISTINCT ""U_invoice"" FROM( " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_ARINV_DocEntry"",'') <> '' UNION ALL " & _
                   " SELECT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_Glar_NC_Rev_Entry"",'') <> '' )T1 "

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
                    If InsertIntoARTable(oDvFinalView, file.Name, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
            ProcessARInvoice_CardCode = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_ERROR)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessARInvoice_CardCode = RTN_ERROR
        End Try
    End Function

    Public Function ProcessARInvoice_OLD(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        'Dim sFuncName As String = "ProcessARInvoice_OLD"
        'Dim sSQL As String

        'Try
        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        '    sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE ""U_status"" = 'C' "
        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
        '    dtSoAcrlInvList = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

        '    sSQL = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE ""U_status"" = 'C' "
        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
        '    dtCostAcrlInvList = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

        '    Dim odtDatatable As DataTable
        '    odtDatatable = oDv.ToTable

        '    odtDatatable.Columns.Add("CostCenter", GetType(String))
        '    odtDatatable.Columns.Add("Insurer", GetType(String))
        '    odtDatatable.Columns.Add("IncuredMonth", GetType(Date))

        '    For intRow As Integer = 0 To odtDatatable.Rows.Count - 1
        '        If Not (odtDatatable.Rows(intRow).Item(1).ToString.Trim() = String.Empty Or odtDatatable.Rows(intRow).Item(1).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
        '            Console.WriteLine("Processing excel line " & intRow & " to get MBMS and Insurer from config table")

        '            Dim sCompCode As String = odtDatatable.Rows(intRow).Item(1).ToString
        '            Dim sCompName As String = odtDatatable.Rows(intRow).Item(0).ToString
        '            sCompName = sCompName.Replace("'", " ")
        '            Dim sSchemeCode As String = odtDatatable.Rows(intRow).Item(3).ToString
        '            Dim sRemarks As String = odtDatatable.Rows(intRow).Item(29).ToString
        '            sRemarks = sRemarks.Replace("'", " ")
        '            Dim sDiagDesc As String = odtDatatable.Rows(intRow).Item(23).ToString
        '            sDiagDesc = sDiagDesc.Replace("'", " ")

        '            If sCompCode = "" Then
        '                sErrDesc = "Company Code should not be empty / Check Line " & intRow
        '                Call WriteToLogFile(sErrDesc, sFuncName)
        '                Console.WriteLine(sErrDesc)
        '                Throw New ArgumentException(sErrDesc)
        '            End If

        '            Dim sInvoice As String = odtDatatable.Rows(intRow).Item(17).ToString.Trim
        '            dtSoAcrlInvList.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
        '            If dtSoAcrlInvList.DefaultView.Count > 0 Then
        '                sErrDesc = "Invoice has been created previously for invoice no :: " & sInvoice
        '                Console.WriteLine(sErrDesc)
        '                Call WriteToLogFile(sErrDesc, sFuncName)
        '                Throw New ArgumentException(sErrDesc)
        '            End If

        '            dtCostAcrlInvList.DefaultView.RowFilter = "U_invoice = '" & sInvoice & "'"
        '            If dtCostAcrlInvList.DefaultView.Count > 0 Then
        '                sErrDesc = "Invoice has been created previously for invoice no :: " & sInvoice
        '                Console.WriteLine(sErrDesc)
        '                Call WriteToLogFile(sErrDesc, sFuncName)
        '                Throw New ArgumentException(sErrDesc)
        '            End If

        '            Dim iIndex As Integer = odtDatatable.Rows(intRow).Item(16).ToString.IndexOf(" ")
        '            Dim sDate As String = odtDatatable.Rows(intRow).Item(16).ToString.Substring(0, iIndex)
        '            Dim dt As Date
        '            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
        '            Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)
        '            Dim dIncurMnth As Date = CDate(dt.Date.AddDays(-(dt.Day - 1)).AddMonths(1).AddDays(-1).ToString())

        '            Dim sCostCenter As String = GetCostCenter(sCompCode, dt, sSchemeCode, p_oCompDef.sSAPDBName)
        '            Dim sInsurer As String = GetInsurer(sCompCode, dt, sSchemeCode, p_oCompDef.sSAPDBName)

        '            If sCostCenter = "" Then
        '                sErrDesc = "MBMS column cannot be null / Check Cost Center for respective company code in config table/Check line " & intRow
        '                Call WriteToLogFile(sErrDesc, sFuncName)
        '                Console.WriteLine(sErrDesc)
        '                Throw New ArgumentException(sErrDesc)
        '            End If
        '            If sInsurer = "" Then
        '                sErrDesc = "Insurer column cannot be null / Check Insurer for the respective company code in config table /Check line " & intRow
        '                Call WriteToLogFile(sErrDesc, sFuncName)
        '                Console.WriteLine(sErrDesc)
        '                Throw New ArgumentException(sErrDesc)
        '            End If

        '            odtDatatable.Rows(intRow)("F1") = sCompName
        '            odtDatatable.Rows(intRow)("F24") = sDiagDesc
        '            odtDatatable.Rows(intRow)("F30") = sRemarks
        '            odtDatatable.Rows(intRow)("CostCenter") = sCostCenter
        '            odtDatatable.Rows(intRow)("Insurer") = sInsurer
        '            odtDatatable.Rows(intRow)("IncuredMonth") = dIncurMnth

        '        End If
        '    Next

        '    Dim oDvFinalView As DataView
        '    oDvFinalView = New DataView(odtDatatable)

        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
        '    Console.WriteLine("Connecting Company")
        '    If ConnectToCompany(p_oCompany, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        '    If p_oCompany.Connected Then
        '        Console.WriteLine("Company connection Successful")
        '        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction", sFuncName)

        '        If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        '        If oDvFinalView.Count > 0 Then

        '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoTable()", sFuncName)

        '            Console.WriteLine("Inserting datas in AR Table")
        '            If InsertIntoARTable(oDvFinalView, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '            Console.WriteLine("Data insert into AR Table Successful")

        '            oDvFinalView.RowFilter = "F5 NOT LIKE 'OUT*'"
        '            Dim odt As New DataTable
        '            odt = oDvFinalView.ToTable
        '            Dim oNewDv As DataView = New DataView(odt)

        '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for A/R Invoice Creation", sFuncName)
        '            Dim oDtGroup As DataTable = oNewDv.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")

        '            Console.WriteLine("Processing datas for A/R Invoice Creation")
        '            For i As Integer = 0 To oDtGroup.Rows.Count - 1
        '                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
        '                    oNewDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
        '                                             " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'"

        '                    If oNewDv.Count > 0 Then
        '                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoice()", sFuncName)
        '                        If CreateARInvoice(oNewDv, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '                    End If

        '                End If
        '            Next
        '            Console.WriteLine("A/R Invoice Creation successful")

        '            oNewDv.RowFilter = Nothing

        '            If oNewDv.Count > 0 Then
        '                Console.WriteLine("Updating Sales Accrual Status and invoice number")
        '                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("calling UpdateSOAccrual()", sFuncName)
        '                If UpdateSOAccrual(oNewDv, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '                Console.WriteLine("Updation of Status and Invoice number in Sales Accrual table successful")
        '            End If

        '            oNewDv.RowFilter = Nothing
        '            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for Creating Reverse sales estimate journal", sFuncName)
        '            oDtGroup = oNewDv.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")

        '            Console.WriteLine("Creating Reverse sales estimate journal")
        '            For i As Integer = 0 To oDtGroup.Rows.Count - 1
        '                If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
        '                    oNewDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
        '                                             " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'"

        '                    If oNewDv.Count > 0 Then
        '                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("calling CreateReverseJournal()", sFuncName)
        '                        Dim dtSoAcrlDatas As DataTable
        '                        dtSoAcrlDatas = oNewDv.ToTable
        '                        Dim oDVSoAcrlDatas As DataView = New DataView(dtSoAcrlDatas)
        '                        If CreateReverseJournal(p_oCompany, oDVSoAcrlDatas, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '                    End If

        '                End If
        '            Next
        '            Console.WriteLine("Reverse estimates of Sales journal created successfully")

        '            oDvFinalView.RowFilter = Nothing
        '            oDvFinalView.RowFilter = "F5 LIKE 'OUT*'"
        '            Dim odtOUT As New DataTable
        '            odtOUT = oDvFinalView.ToTable

        '            Dim oOutDv As DataView = New DataView(odtOUT)

        '            If oOutDv.Count > 0 Then
        '                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for creating A/P invoice", sFuncName)
        '                Console.WriteLine("Creating A/P Invoice")

        '                oDtGroup = oOutDv.Table.DefaultView.ToTable(True, "F5", "CostCenter")
        '                For i As Integer = 0 To oDtGroup.Rows.Count - 1
        '                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
        '                        oOutDv.RowFilter = "F5 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "

        '                        If oOutDv.Count > 0 Then
        '                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateAPInvoice", sFuncName)
        '                            If CreateAPInvoice(p_oCompany, oOutDv, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '                        End If

        '                    End If
        '                Next
        '                Console.WriteLine("A/P invoice created successfully")

        '                oOutDv.RowFilter = Nothing

        '                Console.WriteLine("Updating Status For Cost Accrual datas")
        '                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ", sFuncName)
        '                If UpdateCostAccrual(p_oCompany, oOutDv, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '                Console.WriteLine("Cost accrual status update successful")


        '                oOutDv.RowFilter = Nothing
        '                oDtGroup = oOutDv.Table.DefaultView.ToTable(True, "F5", "CostCenter")
        '                Console.WriteLine("Creating Reverse Journal for Cost Accrual")

        '                For i As Integer = 0 To oDtGroup.Rows.Count - 1
        '                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
        '                        oOutDv.RowFilter = "F5 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "

        '                        If oOutDv.Count > 0 Then
        '                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateReverseJournal_CostAccrual", sFuncName)
        '                            Dim dtCostAcrlDatas As DataTable
        '                            dtCostAcrlDatas = oOutDv.ToTable
        '                            Dim oDVCostAcrlDatas As DataView = New DataView(dtCostAcrlDatas)
        '                            If CreateReverseJournal_CostAccrual(p_oCompany, oDVCostAcrlDatas, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '                        End If

        '                    End If
        '                Next
        '                Console.WriteLine("Reverse Journal for cost accrual completed")
        '            End If
        '        End If
        '    End If

        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
        '    If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
        '    FileMoveToArchive(file, file.FullName, RTN_SUCCESS)

        '    'Insert Success Notificaiton into Table..
        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
        '    AddDataToTable(p_oDtSuccess, file.Name, "Success")
        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & file.FullName, sFuncName)

        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        '    ProcessARInvoice_OLD = RTN_SUCCESS
        'Catch ex As Exception
        '    sErrDesc = ex.Message
        '    Call WriteToLogFile(sErrDesc, sFuncName)

        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
        '    If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        '    'Insert Error Description into Table
        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
        '    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)
        '    'error condition

        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
        '    FileMoveToArchive(file, file.FullName, RTN_ERROR)

        '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        '    ProcessARInvoice_OLD = RTN_ERROR
        'End Try
    End Function

    Private Function InsertIntoARTable(ByVal oDv As DataView, ByVal sFileName As String, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
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
                    sClinicCode = "V" & oDv(i)(4).ToString

                    sSql = " INSERT INTO " & p_oCompDef.sSAPDBName & ".""@AE_MS007_AR""(""Code"",""Name"",""U_company"",""U_company_code"",""U_C"",""U_scheme_code"",""U_cln_code"",""U_m_id_type"",""U_m_id"",""U_m_lastname"",""U_m_given_name"",""U_m_christian""," & _
                            " ""U_relation"",""U_id_type"",""U_id"",""U_lastname"",""U_given_name"",""U_christian"",""U_txn_date"",""U_invoice"",""U_treatment"",""U_charge"",""U_pay_comp"",""U_pay_client"",""U_diag"",""U_diag_desc"", " & _
                            " ""U_refer_from_name"",""U_policy_num"",""U_cert_num"",""U_treat_code"",""U_remark_fg"",""U_remark1"",""U_paiddate"",""U_status"",""U_status_code"",""U_cust_no"",""U_scheme_remark"",""U_dept1"",""U_dept2""," & _
                            " ""U_dept3"",""U_ds1"",""U_ds2"",""U_ds3"",""U_in_time"",""U_insco"",""U_sl_fr"",""U_sl_to"",""U_CompTotRecCnt"",""U_CompTotBillAmt"",""U_scheme_desc"",""U_OcrCode"",""U_Insurer"",""U_Incurred_month"",""U_ar_code"",""U_ap_code"",""U_Type"",""U_FileName"")" & _
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
                            " '" & oDv(i)(48).ToString & "','" & oDv(i)(49).ToString & "','" & oDv(i)(50).ToString & "','" & sCompCode & "','" & sClinicCode & "','" & oDv(i)(51).ToString & "','" & sFileName & "' )"

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

    Private Function ProcessDatas_NonTPAListings(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessDatas_NonTPAListings"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oDv.RowFilter = "Type NOT LIKE 'CAPITATION*'"
            Dim odt As New DataTable
            odt = oDv.ToTable
            Dim oNewDv As DataView = New DataView(odt)

            If oNewDv.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping databased based on incurred month and MBMS", sFuncName)
                'F2 - Company Code F18 Invoice
                Dim oDtGroup As DataTable = oNewDv.Table.DefaultView.ToTable(True, "CostCenter", "IncuredMonth")

                Console.WriteLine("Processing Datas for A/R invoice Creation and Reversal Journal")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COSTCENTER") Then
                        oNewDv.RowFilter = "CostCenter = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and IncuredMonth ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' "

                        If oNewDv.Count > 0 Then
                            Console.WriteLine("Processing grouped data line : " & i)
                            Dim odtAcrualDts As DataTable
                            odtAcrualDts = oNewDv.ToTable
                            Dim oDvAcrlDatas As DataView = New DataView(odtAcrualDts)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoice_NonTpaListing()", sFuncName)
                            If CreateARInvoice_NonTpaListing(oDvAcrlDatas, p_oCompany, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            Console.WriteLine("Invoice created successfully for grouped data line " & i)
                        End If
                    End If
                Next

            End If

            oDv.RowFilter = "AcrlType NOT LIKE 'CAPITATION*'"
            Dim oNonCap_RevJrnl_NonTpaList As New DataTable
            oNonCap_RevJrnl_NonTpaList = oDv.ToTable
            Dim oDvNonCap_RevJrnl_NonTpaList As DataView = New DataView(oNonCap_RevJrnl_NonTpaList)

            If oDvNonCap_RevJrnl_NonTpaList.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping databased based on company code,incurred month and MBMS", sFuncName)
                'F2 - Company Code F18 Invoice
                Dim oDtGroup As DataTable = oDvNonCap_RevJrnl_NonTpaList.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")

                Console.WriteLine("Processing Datas for creating Reversal Journal entry")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
                        oDvNonCap_RevJrnl_NonTpaList.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                           " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "

                        If oDvNonCap_RevJrnl_NonTpaList.Count > 0 Then
                            Dim odtAcrualDts As DataTable
                            odtAcrualDts = oDvNonCap_RevJrnl_NonTpaList.ToTable
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

    Private Function CreateARInvoice_NonTpaListing(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateARInvoice"
        Dim sSql As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sSqlQuery As String = String.Empty
        Dim sDocDate As String = String.Empty
        Dim dTotPayComp, dPayComp As Double
        Dim sCostCenter As String = String.Empty
        Dim iRetCode, iErrCode As Integer
        Dim bLineAdded As Boolean = False
        Dim iCount As Integer = 0
        Dim sIncuredMnth As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim sVatGroup As String = String.Empty
            Dim sItemCode As String = String.Empty

            sSql = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_ITEMCODE"" WHERE UPPER(""U_FileCode"") = 'MS007' AND UPPER(""U_Field"") = 'PAY_COMP' AND ""U_Outnetwork"" = 'N'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
            sItemCode = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            If sItemCode = "" Then
                sErrDesc = "Check ItemCode in configuration table"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dtItemCode.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
            If dtItemCode.DefaultView.Count = 0 Then
                sErrDesc = "Item Code not found in SAP. ItemCode :: " & sItemCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            sDocDate = file.Name.Substring(9, 8)

            Dim dt As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)
            'Dim dDocDate As Date = CDate(dt.Date.AddDays(-(dt.Day - 1)).AddMonths(1).AddDays(-1).ToString())

            dPayComp = 0
            dTotPayComp = 0

            For i As Integer = 0 To oDv.Count - 1
                dPayComp = CDbl(oDv(i)(20).ToString.Trim)
                dTotPayComp = dTotPayComp + dPayComp
            Next

            Dim sCCode As String = file.Name
            sCCode = sCCode.Substring(18, sCCode.Length - 18)
            Dim iPos As Integer = sCCode.IndexOf("_")
            sCCode = sCCode.Substring(0, iPos)

            If dTotPayComp > 0 Then
                'sCardCode = "C" & oDv(0)(1).ToString.Trim
                sCardCode = sCCode
                sCostCenter = oDv(0)(48).ToString.Trim
                sIncuredMnth = oDv(0)(50).ToString.Trim

                Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
                Dim dIncuredMnth As Date = CDate(sIncuredMnth.Substring(0, iIndex))

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

                Console.WriteLine("Creating Invoice for " & sCardCode)

                Dim oSalInvoice As SAPbobsCOM.Documents
                oSalInvoice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                oSalInvoice.CardCode = sCardCode
                oSalInvoice.DocDate = dt
                oSalInvoice.DocDueDate = dt
                oSalInvoice.UserFields.Fields.Item("U_Footer").Value = p_oCompDef.sARInvFooter
                oSalInvoice.Comments = "Consultation fee for " & dIncuredMnth.Month & " " & dIncuredMnth.Year

                If iCount > 1 Then
                    oSalInvoice.Lines.Add()
                End If

                oSalInvoice.Lines.ItemCode = sItemCode
                oSalInvoice.Lines.ItemDescription = "Consultation fee for " & dIncuredMnth.Month.ToString() & "-" & dIncuredMnth.Year
                oSalInvoice.Lines.Quantity = 1
                If Not (sVatGroup = String.Empty) Then
                    oSalInvoice.Lines.VatGroup = sVatGroup
                End If
                oSalInvoice.Lines.UnitPrice = dTotPayComp
                If Not (sCostCenter = String.Empty) Then
                    oSalInvoice.Lines.CostingCode = sCostCenter
                    oSalInvoice.Lines.COGSCostingCode = sCostCenter
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                iRetCode = oSalInvoice.Add()

                If iRetCode <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    bLineAdded = False
                    Dim iDocNo, iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalInvoice)

                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim sQuery As String

                    sSql = "SELECT ""DocNum"" FROM " & p_oCompDef.sSAPDBName & ".""OINV"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSql)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = String.Empty
                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" SET ""U_ARINV_DocNo"" = '" & iDocNo & "',""U_ARINV_DocEntry"" = '" & iDocEntry & "' " & _
                                     " WHERE ""U_OcrCode"" = '" & sCostCenter & "' AND ""U_invoice"" = '" & sInvoice & "' " & _
                                     " AND ""U_Incurred_month"" = '" & sIncuredMnth & "' AND IFNULL(""U_ARINV_DocNo"",'') = '' AND IFNULL(""U_ARINV_DocEntry"",'') = '' "
                            objRS.DoQuery(sQuery)

                            sQuery = String.Empty
                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_MS007_AR"" SET ""U_ARInvoiceNo"" = '" & iDocNo & "' " & _
                                     " WHERE ""U_invoice"" = '" & sInvoice & "' AND IFNULL(""U_ARInvoiceNo"",'') = '' "
                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateARInvoice_NonTpaListing = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateARInvoice_NonTpaListing = RTN_ERROR
            Throw New Exception(ex.Message)
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
        Dim sSource As String = file.Name.Substring(0, 5)
        Dim iErrCode As Integer
        Dim sXcelInvNo As String
        Dim sCompCode1 As String = String.Empty
        Dim sAcrlJrnlId As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sDocDate = file.Name.Substring(9, 8)

            Dim dt As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS007_GL_REV"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE A.""U_FileCode"" = 'MS007' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            'sCreditAct = GetActCode_RevEstiSaleGL(sSource, "C")
            sCreditAct = GetStringValue(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS007_GL_REV"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSql = sSql & " WHERE A.""U_FileCode"" = 'MS007' AND A.""U_ActType"" = 'D'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)
            'sDebitAct = GetActCode_RevEstiSaleGL(sSource, "D")
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

                    '""U_status"" = 'O' AND
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
                If Not (sCostCenter = String.Empty) Then
                    oJournalEntry.Memo = "Reversal of Estimated sales for " & sCompCode & " " & sCostCenter
                Else
                    oJournalEntry.Memo = "Reversal of Estimated sales for " & sCompCode
                End If

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

    Private Function ProcessDatas_TPAListings(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessDatas_TPAListings"
        Dim sSQL As String = String.Empty
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oDv.RowFilter = "Type NOT LIKE 'CAPITATION*'"
            Dim odt As New DataTable
            odt = oDv.ToTable
            Dim oNewDv As DataView = New DataView(odt)

            If oNewDv.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping databased based on incurred month and MBMS", sFuncName)
                Dim oDtGroup As DataTable = oNewDv.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")

                Console.WriteLine("Processing Datas for A/R invoice Creation and Reversal Journal")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
                        oNewDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "'  and CostCenter ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                           "  and IncuredMonth ='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'"

                        If oNewDv.Count > 0 Then
                            Console.WriteLine("Processing grouped data line : " & i)
                            Dim odtARInvDts As DataTable
                            odtARInvDts = oNewDv.ToTable
                            Dim oDvARInvDts As DataView = New DataView(odtARInvDts)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateARInvoice_NonTpaListing()", sFuncName)
                            If CreateARInvoice_Capitation_TpaListing(oDvARInvDts, p_oCompany, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            Console.WriteLine("Invoice created successfully for grouped data line " & i)

                        End If
                    End If
                Next
            End If

            oDv.RowFilter = Nothing

            oDv.RowFilter = "AcrlType NOT LIKE 'CAPITATION*'"
            Dim oCapDt_RevJrnl_TpaList As New DataTable
            oCapDt_RevJrnl_TpaList = oDv.ToTable
            Dim oCapDv_RevJrnl_TpaList As DataView = New DataView(oCapDt_RevJrnl_TpaList)

            If oCapDv_RevJrnl_TpaList.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping databased based on company code,incurred month and MBMS", sFuncName)
                'F2 - Company Code
                Dim oDtGroup As DataTable = oCapDv_RevJrnl_TpaList.Table.DefaultView.ToTable(True, "F2", "CostCenter", "IncuredMonth")

                Console.WriteLine("Processing Datas for A/R invoice Creation and Reversal Journal")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "COMPANY_CODE") Then
                        oCapDv_RevJrnl_TpaList.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and CostCenter ='" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                           " and IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' "

                        If oCapDv_RevJrnl_TpaList.Count > 0 Then
                            Dim odtARInvDts As DataTable
                            odtARInvDts = oCapDv_RevJrnl_TpaList.ToTable
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

    Private Function CreateARInvoice_Capitation_TpaListing(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateARInvoice_Capitation_TpaListing"
        Dim sSQL As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sSqlQuery As String = String.Empty
        Dim sDocDate As String = String.Empty
        Dim dTotPayComp, dPayComp As Double
        Dim sCostCenter As String = String.Empty
        Dim iRetCode, iErrCode As Integer
        Dim bLineAdded As Boolean = False
        Dim iCount As Integer = 0
        Dim sIncuredMnth As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim sVatGroup As String = String.Empty
            Dim sItemCode As String = String.Empty

            sSQL = "SELECT ""U_SAPItemCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_ITEMCODE"" WHERE UPPER(""U_FileCode"") = 'MS007' AND UPPER(""U_Field"") = 'PAY_COMP' " & _
                   " AND ""U_Outnetwork"" = 'Y' AND IFNULL(""U_Type"",'') = ''"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            sItemCode = GetStringValue(sSQL, p_oCompDef.sSAPDBName)

            If sItemCode = "" Then
                sErrDesc = "Check ItemCode in configuration table"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            dtItemCode.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
            If dtItemCode.DefaultView.Count = 0 Then
                sErrDesc = "Item Code not found in SAP. ItemCode :: " & sItemCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            sDocDate = file.Name.Substring(9, 8)

            Dim dt As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)
            'Dim dDocDate As Date = CDate(dt.Date.AddDays(-(dt.Day - 1)).AddMonths(1).AddDays(-1).ToString())

            dPayComp = 0
            dTotPayComp = 0

            For i As Integer = 0 To oDv.Count - 1
                dPayComp = CDbl(oDv(i)(20).ToString.Trim)
                dTotPayComp = dTotPayComp + dPayComp
            Next

            'Dim iPos As Integer = file.Name.LastIndexOf("_")
            'Dim sCCode As String = file.Name.Substring(iPos + 1, (file.Name.Length - iPos - 1))
            'sCCode = sCCode.Replace(".xls", "")

            Dim sCCode As String = file.Name
            sCCode = sCCode.Substring(18, sCCode.Length - 18)
            Dim iPos As Integer = sCCode.IndexOf("_")
            If iPos > 0 Then
                sCCode = sCCode.Substring(0, iPos)
            End If
            sCCode = sCCode.Replace(".xls", "")

            If dTotPayComp > 0 Then
                'sCardCode = "C" & oDv(0)(1).ToString.Trim
                sCardCode = sCCode
                sCostCenter = oDv(0)(48).ToString.Trim
                sIncuredMnth = oDv(0)(50).ToString.Trim

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

                Console.WriteLine("Creating Invoice for " & sCardCode)

                Dim oSalInvoice As SAPbobsCOM.Documents
                oSalInvoice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                oSalInvoice.CardCode = sCardCode
                oSalInvoice.DocDate = dt
                oSalInvoice.DocDueDate = dt
                oSalInvoice.UserFields.Fields.Item("U_Footer").Value = p_oCompDef.sARInvFooter

                If iCount > 1 Then
                    oSalInvoice.Lines.Add()
                End If

                oSalInvoice.Lines.ItemCode = sItemCode
                oSalInvoice.Lines.Quantity = 1
                If Not (sVatGroup = String.Empty) Then
                    oSalInvoice.Lines.VatGroup = sVatGroup
                End If
                oSalInvoice.Lines.UnitPrice = dTotPayComp
                If Not (sCostCenter = String.Empty) Then
                    oSalInvoice.Lines.CostingCode = sCostCenter
                    oSalInvoice.Lines.COGSCostingCode = sCostCenter
                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding Document", sFuncName)

                iRetCode = oSalInvoice.Add()

                If iRetCode <> 0 Then
                    p_oCompany.GetLastError(iErrCode, sErrDesc)
                    Throw New ArgumentException(sErrDesc)
                Else
                    bLineAdded = False
                    Dim iDocNo, iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalInvoice)

                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim sQuery As String

                    sSQL = "SELECT ""DocNum"" FROM " & p_oCompDef.sSAPDBName & ".""OINV"" WHERE ""DocEntry"" ='" & iDocEntry & "'"
                    objRS.DoQuery(sSQL)
                    If objRS.RecordCount > 0 Then
                        iDocNo = objRS.Fields.Item("DocNum").Value
                    End If
                    Console.WriteLine("Document Created successfully :: " & iDocNo)

                    Dim oDtGroup As DataTable = oDv.Table.DefaultView.ToTable(True, "F18")
                    For k As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(k).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            Dim sInvoice As String = oDtGroup.Rows(k).Item(0).ToString.ToUpper().Trim()

                            sQuery = String.Empty
                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_ARINV_DocNo"" = '" & iDocNo & "',""U_ARINV_DocEntry"" = '" & iDocEntry & "' " & _
                                     " WHERE ""U_OcrCode"" = '" & sCostCenter & "' AND IFNULL(""U_ARINV_DocNo"",'') = '' AND IFNULL(""U_ARINV_DocEntry"",'') = '' " & _
                                     " AND ""U_incurred_month"" = '" & sIncuredMnth & "' AND ""U_invoice"" = '" & sInvoice & "'"

                            objRS.DoQuery(sQuery)

                            sQuery = String.Empty
                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_MS007_AR"" SET ""U_ARInvoiceNo"" = '" & iDocNo & "' " & _
                                     " WHERE ""U_invoice"" = '" & sInvoice & "' AND IFNULL(""U_ARInvoiceNo"",'') = '' "
                            objRS.DoQuery(sQuery)
                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRS)

                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateARInvoice_Capitation_TpaListing = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateARInvoice_Capitation_TpaListing = RTN_ERROR

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
        Dim sSource As String = file.Name.Substring(0, 5)
        Dim iErrCode As Integer
        Dim sXcelInvNo As String
        Dim sCompCode1 As String = String.Empty
        Dim sAcrlJrnlId As String = String.Empty
        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sDocDate = file.Name.Substring(9, 8)

            Dim dt As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd"}
            Date.TryParseExact(sDocDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dt)
            'Dim dDocDate As Date = CDate(dt.Date.AddDays(-(dt.Day - 1)).AddMonths(1).AddDays(-1).ToString())

            sSQL = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAR_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSQL = sSQL & " WHERE A.""U_FileCode"" = 'MS007' AND A.""U_ActType"" = 'D'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            'sCreditAct = GetActCode_RevEstiSaleGL(sSource, "C")
            sCreditAct = GetStringValue(sSQL, p_oCompDef.sSAPDBName)

            sSQL = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_OUT_GLAR_NONCAP"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
            sSQL = sSQL & " WHERE A.""U_FileCode"" = 'MS007' AND A.""U_ActType"" = 'C'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSQL, sFuncName)
            'sDebitAct = GetActCode_RevEstiSaleGL(sSource, "D")
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

                    'AND ""U_status"" = 'O'
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

            'For i As Integer = 0 To odv.Count - 1
            '    dPayCompAmt = CDbl(odv(i)(20).ToString.Trim)
            '    dTotPayCompAmt = dTotPayCompAmt + dPayCompAmt
            'Next

            If dTotPayCompAmt > 0 Then
                sCompCode = odv(0)(1).ToString.Trim
                sCostCenter = odv(0)(48).ToString.Trim

                oJournalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                oJournalEntry.TaxDate = dt
                oJournalEntry.ReferenceDate = dt
                oJournalEntry.Memo = "Reversal of Est TPA Claim for " & sCompCode & " " & sCostCenter

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

End Module
