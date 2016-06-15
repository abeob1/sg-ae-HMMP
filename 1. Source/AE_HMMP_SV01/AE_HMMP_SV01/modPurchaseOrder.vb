Module modPurchaseOrder

    Private dtSoAcrlInvList As DataTable
    Private dtCostAcrlInvList As DataTable
    Private dtInsurerList As DataTable
    Private dtMBMSList As DataTable
    Private dtCardCode As DataTable

    Public Function ProcessPurchaseOrder(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessPurchaseOrder"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSql = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE IFNULL(""U_JrnlEntry"",'') <> '' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtSoAcrlInvList = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE IFNULL(""U_JrnlEntry"",'') <> '' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtCostAcrlInvList = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT DISTINCT ""CardCode"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtCardCode = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

            Dim odtDatatable As DataTable
            odtDatatable = oDv.ToTable

            odtDatatable.Columns.Add("CostCenter", GetType(String))
            odtDatatable.Columns.Add("Insurer", GetType(String))
            odtDatatable.Columns.Add("IncuredMonth", GetType(Date))
            odtDatatable.Columns.Add("Type", GetType(String))

            For intRow As Integer = 0 To odtDatatable.Rows.Count - 1
                If Not (odtDatatable.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or odtDatatable.Rows(intRow).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    Console.WriteLine("Processing excel line " & intRow & " to get MBMS and Insurer from config table")

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

                    Dim sType As String = String.Empty
                    Dim sArCode As String = String.Empty
                    sArCode = "C" & sCompCode

                    'dtCardCode.DefaultView.RowFilter = "CardCode = '" & sArCode & "'"
                    'If dtCardCode.DefaultView.Count = 0 Then
                    '    sErrDesc = "Cardcode not found in SAP / Check Cardcode :: " & sArCode
                    '    Console.WriteLine(sErrDesc)
                    '    Call WriteToLogFile(sErrDesc, sFuncName)
                    '    Throw New ArgumentException(sErrDesc)
                    'End If

                    sSql = "SELECT ""U_Type"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" WHERE ""CardCode"" = '" & sArCode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                    sType = GetActCode(sSql, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

                    If sType = "" Then
                        sType = p_oCompDef.sType
                    End If

                    Dim sInvoice As String = odtDatatable.Rows(intRow).Item(0).ToString.Trim
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

                    odtDatatable.Rows(intRow)("F6") = sCompName
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

                If oDvFinalView.Count > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoTable()", sFuncName)

                    Console.WriteLine("Inserting datas into PO Table")
                    If InsertIntoPOTable(oDvFinalView, file.Name, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Console.WriteLine("Data insert into PO Table Successful")

                    '***************************NEW LOGIC STARTS********************************
                    Dim oDtGroup As DataTable = oDvFinalView.Table.DefaultView.ToTable(True, "F1", "F2", "CostCenter", "IncuredMonth")
                    Console.WriteLine("Processing datas for type not capitation")
                    For i As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            oDvFinalView.RowFilter = "F1 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F2 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                     " and CostCenter='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' and IncuredMonth='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "' "

                            If oDvFinalView.Count > 0 Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoCostAcrual_PO()", sFuncName)
                                If InsertIntoCostAcrual_PO(oDvFinalView, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    Next
                    Console.WriteLine("Processing of Cost accrual datas completed for type not capitation")

                    oDvFinalView.RowFilter = Nothing

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas based on Type column", sFuncName)

                    oDvFinalView.RowFilter = "Type NOT LIKE 'CAPITATION*'"
                    Dim odtNonCap As New DataTable
                    odtNonCap = oDvFinalView.ToTable

                    Dim oNonCapDv As DataView = New DataView(odtNonCap)

                    If oNonCapDv.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for creating journal entry for type not capitation", sFuncName)

                        oDtGroup = oNonCapDv.Table.DefaultView.ToTable(True, "F2", "F3", "IncuredMonth")
                        Console.WriteLine("Creating journal entry for type not capitation")

                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oNonCapDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' AND F3 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' AND IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'"
                                If oNonCapDv.Count > 0 Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling PO_CreateJournalEntry()", sFuncName)
                                    Dim dtNonCapFinal As DataTable
                                    dtNonCapFinal = oNonCapDv.ToTable
                                    Dim oDvNonCapFinal As DataView = New DataView(dtNonCapFinal)
                                    If PO_CreateJournalEntry_NonCapitation(p_oCompany, oDvNonCapFinal, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                            End If
                        Next
                        Console.WriteLine("Journal Entry Creation Successful for type not capitation")
                    End If

                    oDvFinalView.RowFilter = Nothing

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas based on Type column", sFuncName)

                    oDvFinalView.RowFilter = "Type LIKE 'CAPITATION*'"
                    Dim odtCap As New DataTable
                    odtCap = oDvFinalView.ToTable

                    Dim oCapDv As DataView = New DataView(odtCap)

                    If oCapDv.Count > 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for creating journal entry for type capitation", sFuncName)

                        oDtGroup = oCapDv.Table.DefaultView.ToTable(True, "F2", "F3", "IncuredMonth")
                        Console.WriteLine("Creating journal entry for type capitation")

                        For i As Integer = 0 To oDtGroup.Rows.Count - 1
                            If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                                oCapDv.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' AND F3 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' AND IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'"
                                If oCapDv.Count > 0 Then
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling PO_CreateJournalEntry()", sFuncName)
                                    Dim dtCapFinal As DataTable
                                    dtCapFinal = oCapDv.ToTable
                                    Dim oDvCapFinal As DataView = New DataView(dtCapFinal)
                                    If PO_CreateJournalEntry_Capitation(p_oCompany, oDvCapFinal, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                            End If
                        Next
                        Console.WriteLine("Journal Entry Creation Successful for type capitation")
                    End If

                    '***************************NEW LOGIC ENDS********************************

                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
                If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
                FileMoveToArchive(file, file.FullName, RTN_SUCCESS)

                'Insert Success Notificaiton into Table..
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                AddDataToTable(p_oDtSuccess, file.Name, "Success")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & file.FullName, sFuncName)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessPurchaseOrder = RTN_SUCCESS

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
            ProcessPurchaseOrder = RTN_ERROR
        End Try
    End Function

    Public Function ProcessPurchaseOrder_OLD(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessPurchaseOrder_OLD"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSql = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL"" WHERE IFNULL(""U_JrnlEntry"",'') <> '' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtSoAcrlInvList = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT DISTINCT ""U_invoice"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" WHERE ""U_status"" = 'C' "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtCostAcrlInvList = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

            sSql = "SELECT DISTINCT ""CardCode"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" "
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSql, sFuncName)
            dtCardCode = ExecuteQueryReturnDataTable(sSql, p_oCompDef.sSAPDBName)

            Dim odtDatatable As DataTable
            odtDatatable = oDv.ToTable

            odtDatatable.Columns.Add("CostCenter", GetType(String))
            odtDatatable.Columns.Add("Insurer", GetType(String))
            odtDatatable.Columns.Add("IncuredMonth", GetType(Date))
            odtDatatable.Columns.Add("Type", GetType(String))

            For intRow As Integer = 0 To odtDatatable.Rows.Count - 1
                If Not (odtDatatable.Rows(intRow).Item(0).ToString.Trim() = String.Empty Or odtDatatable.Rows(intRow).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                    Console.WriteLine("Processing excel line " & intRow & " to get MBMS and Insurer from config table")

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

                    sSql = "SELECT " & p_oCompDef.sSAPDBName & ".""U_Type"" FROM ""OCRD"" WHERE ""CardCode"" = '" & sArCode & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL " & sSql, sFuncName)
                    sType = GetActCode(sSql, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd)

                    Dim sInvoice As String = odtDatatable.Rows(intRow).Item(0).ToString.Trim
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

                    odtDatatable.Rows(intRow)("F6") = sCompName
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

                If oDvFinalView.Count > 0 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoTable()", sFuncName)

                    Console.WriteLine("Inserting datas into PO Table")
                    If InsertIntoPOTable(oDvFinalView, file.Name, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Console.WriteLine("Data insert into PO Table Successful")

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas based on Clinic Code and incurred month", sFuncName)

                    'F2 - Clinic Code F1-Invoice
                    Dim oDtGroup As DataTable = oDvFinalView.Table.DefaultView.ToTable(True, "F1", "F2", "CostCenter", "IncuredMonth", "Type")

                    Console.WriteLine("Processing Cost Accrual datas")
                    For i As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "INVOICE") Then
                            oDvFinalView.RowFilter = "F1 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' and F2 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' " & _
                                                     " and CostCenter='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "' and IncuredMonth='" & oDtGroup.Rows(i).Item(3).ToString.Trim() & "' "

                            If oDvFinalView.Count > 0 Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling InsertIntoCostAcrual_PO()", sFuncName)
                                If InsertIntoCostAcrual_PO(oDvFinalView, p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    Next
                    Console.WriteLine("Processing of Cost accrual datas completed")

                    oDvFinalView.RowFilter = Nothing

                    oDtGroup = oDvFinalView.Table.DefaultView.ToTable(True, "F2", "F3", "IncuredMonth")
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping datas for creating journal entry", sFuncName)
                    For i As Integer = 0 To oDtGroup.Rows.Count - 1
                        If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "CLN_CODE") Then
                            oDvFinalView.RowFilter = "F2 = '" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' AND F3 = '" & oDtGroup.Rows(i).Item(1).ToString.Trim() & "' AND IncuredMonth='" & oDtGroup.Rows(i).Item(2).ToString.Trim() & "'"
                            If oDvFinalView.Count > 0 Then
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling PO_CreateJournalEntry()", sFuncName)
                                Dim dtFinalTable As DataTable
                                dtFinalTable = oDvFinalView.ToTable
                                Dim oDvFinal As DataView = New DataView(dtFinalTable)
                                ' If PO_CreateJournalEntry(p_oCompany, oDvFinal, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    Next
                    Console.WriteLine("Journal Entry Creation Successful")

                End If

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
                If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
                FileMoveToArchive(file, file.FullName, RTN_SUCCESS)

                'Insert Success Notificaiton into Table..
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                AddDataToTable(p_oDtSuccess, file.Name, "Success")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & file.FullName, sFuncName)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessPurchaseOrder_OLD = RTN_SUCCESS

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
            ProcessPurchaseOrder_OLD = RTN_ERROR
        End Try
    End Function

    Private Function InsertIntoPOTable(ByVal oDV As DataView, ByVal sFileName As String, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "InsertIntoPOTable"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'LINE BY LINE
            For i As Integer = 1 To oDV.Count - 1
                If Not (oDV(i)(0).ToString.Trim = String.Empty) Then
                    Console.WriteLine("Inserting Line Num : " & i)
                    sSql = String.Empty

                    Dim sCompCode As String = oDV(i)(6).ToString.Trim
                    If sCompCode <> "" Then
                        sCompCode = "C" & sCompCode
                    End If
                    Dim sClinicCode As String = oDV(i)(1).ToString.Trim
                    Dim sSubCode As String = oDV(i)(2).ToString.Trim
                    If sClinicCode <> "" Then
                        sClinicCode = "V" & sClinicCode & sSubCode
                    End If


                    sSql = "INSERT INTO " & p_oCompDef.sSAPDBName & ".""@AE_MS002_PO"" (""Code"",""Name"",""U_invoice"",""U_cln_code"",""U_subcode"",""U_cln_name"",""U_txn_date""," & _
                            " ""U_company"",""U_company_code"",""U_scheme_code"",""U_m_id_type"",""U_m_id"",""U_id_type"",""U_id"",""U_treat_code"", " & _
                            " ""U_treatment"",""U_charge"",""U_pay_comp"",""U_pay_client"",""U_oper"",""U_ds"",""U_reimburse"",""U_cmoney"",""U_diag_desc"", " & _
                            " ""U_refer_from_name"",""U_lastname"",""U_given_name"",""U_christian"",""U_remark_fg"",""U_manualfee"",""U_in_time"",""U_status"", " & _
                            " ""U_sl_fr"",""U_sl_to"",""U_txn_remark_type"",""U_txn_remark"",""U_txn_remark_userid"",""U_create_datetime"",""U_create_userid"", " & _
                            " ""U_OcrCode"",""U_Insurer"",""U_incurred_month"",""U_ar_code"",""U_ap_code"",""U_Type"",""U_FileName"" ) " & _
                            " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS002_PO""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM """ & p_oCompDef.sSAPDBName & """.""@AE_MS002_PO"")," & _
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
                            " '" & sCompCode & "','" & sClinicCode & "',(SELECT ""U_Type"" FROM ""OCRD"" WHERE ""CardCode"" = '" & sCompCode & "'),'" & sFileName & "')"


                    oRecSet.DoQuery(sSql)
                End If
            Next
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoPOTable = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            InsertIntoPOTable = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function InsertIntoCostAcrual_PO(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As String
        Dim sFuncName As String = "InsertIntoCostAcrual_PO"
        Dim sSql As String = String.Empty

        Try
            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim dPayComp, dTotPayComp, dCMoney, dTotCmoney, dOper, dTotOper, dPayClient, dTotPayClient As Double
            For i As Integer = 0 To oDv.Count - 1
                dPayComp = 0
                dCMoney = 0
                dOper = 0
                dPayClient = 0

                dPayComp = CDbl(oDv(i)(15).ToString.Trim)
                dCMoney = CDbl(oDv(i)(20).ToString.Trim)
                dOper = CDbl(oDv(i)(17).ToString.Trim)
                dPayClient = CDbl(oDv(i)(16).ToString.Trim)

                dTotPayComp = dTotPayComp + dPayComp
                dTotCmoney = dTotCmoney + dCMoney
                dTotOper = dTotOper + dOper
                dTotPayClient = dTotPayClient + dPayClient
            Next

            Dim sApCode As String
            Dim sClinicCode As String = oDv(0)(1).ToString.Trim
            Dim sSubCode As String = oDv(0)(2).ToString.Trim
            sApCode = "V" & sClinicCode & sSubCode

            sSql = "INSERT INTO " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL""(""Code"",""Name"",""U_cln_code"",""U_ap_code"",""U_incurred_month"","
            sSql = sSql & " ""U_OcrCode"",""U_Insurer"",""U_invoice"",""U_cmoney"",""U_pay_client"",""U_oper"",""U_pay_comp"",""U_source"",""U_status"",""U_Type"")"
            sSql = sSql & " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL""),"
            sSql = sSql & " '" & oDv(0)(1).ToString & "','" & sApCode & "','" & oDv(0)(39).ToString & "',"
            sSql = sSql & " '" & oDv(0)(37).ToString & "','" & oDv(0)(38).ToString & "','" & oDv(0)(0).ToString & "','" & dTotCmoney & "','" & dTotPayClient & "', "
            sSql = sSql & " '" & dTotOper & "','" & dTotPayComp & "','MS002','O','" & oDv(0)(40).ToString & "')"

            oRecSet.DoQuery(sSql)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoCostAcrual_PO = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while executing query", sFuncName)
            InsertIntoCostAcrual_PO = RTN_ERROR
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function PO_CreateJournalEntry_NonCapitation(ByVal oCompany As SAPbobsCOM.Company, ByVal oDv As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "PO_CreateJournalEntry_NonCapitation"
        Dim sSql, sQuery As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCostCenter As String = String.Empty
        Dim sIncuredMnth As String = String.Empty
        Dim sClinicCod As String = String.Empty
        Dim dCmoney, dPayClient, dOper As Double
        Dim sOperAct, sCMoneyAct, sPayClntAct, sActCode As String
        Dim iErrCode, iCount As Integer
        Dim bIsLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim sSubCode As String = oDv(0)(2).ToString.Trim
            sClinicCod = oDv(0)(1).ToString.Trim
            sIncuredMnth = oDv(0)(39).ToString.Trim

            Dim sApCode As String
            sApCode = "V" & sClinicCod & sSubCode

            dtCardCode.DefaultView.RowFilter = "CardCode = '" & sApCode & "'"
            If dtCardCode.DefaultView.Count = 0 Then
                sErrDesc = "Cardcode not found in SAP / Check Cardcode :: " & sApCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

            Dim iIndex As Integer = sIncuredMnth.IndexOf(" ")
            Dim dt As Date = CDate(sIncuredMnth.Substring(0, iIndex))

            sQuery = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS002_GL"" A INNER JOIN ""OACT"" B ON B.""FormatCode"" = A.""U_GLCode"" " & _
                     " WHERE UPPER(A.""U_Field"") = 'CMONEY'"
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
            oJournalEntry.Memo = "Estimated cost for clinic " & sApCode

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

                        For j As Integer = 0 To oDv.Count - 1
                            dCmoney = dCmoney + CDbl(oDv(j)(20).ToString.Trim)
                            dPayClient = dPayClient + CDbl(oDv(j)(16).ToString.Trim)
                            dOper = dOper + CDbl(oDv(j)(17).ToString.Trim)
                        Next

                        dCmoney = Math.Round(dCmoney, 2)
                        dPayClient = Math.Round(dPayClient, 2)
                        dOper = Math.Round(dOper, 2)

                        If dCmoney <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.AccountCode = sCMoneyAct
                            oJournalEntry.Lines.Debit = dCmoney
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If
                        If dOper <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.ShortName = sOperAct
                            oJournalEntry.Lines.Credit = dOper
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If
                        If dPayClient <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.ShortName = sPayClntAct
                            oJournalEntry.Lines.Credit = dPayClient
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If

                        Dim dTotval As Double
                        dTotval = Math.Round((dCmoney - dOper - dPayClient), 2)

                        If dTotval <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.ShortName = sActCode
                            oJournalEntry.Lines.Credit = dTotval
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                oJournalEntry.Lines.CostingCode2 = sCostCenter
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

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_Journal"" = '" & iDocNo & "', ""U_JrnlEntry"" = '" & iJournalEntryNo & "' " & _
                                     " WHERE ""U_cln_code"" = '" & sClinicCod & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' AND ""U_source"" = 'MS002' " & _
                                     " AND IFNULL(""U_JrnlEntry"",'') = '' AND ""U_invoice"" = '" & sInvoice & "' "
                            oRs.DoQuery(sQuery)

                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            PO_CreateJournalEntry_NonCapitation = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            PO_CreateJournalEntry_NonCapitation = RTN_ERROR
        End Try
    End Function

    Private Function PO_CreateJournalEntry_Capitation(ByVal oCompany As SAPbobsCOM.Company, ByVal oDv As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "PO_CreateJournalEntry_Capitation"
        Dim sSql, sQuery As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oJournalEntry As SAPbobsCOM.JournalEntries
        Dim sCostCenter As String = String.Empty
        Dim sIncuredMnth As String = String.Empty
        Dim sClinicCod As String = String.Empty
        Dim dCmoney, dPayClient, dOper, dCMoneyClient As Double
        Dim sOperAct, sPayClntAct_Debit, sPayClnt_Credit, sActCode, sCMoneyClient As String
        Dim iErrCode, iCount As Integer
        Dim bIsLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Dim sSubCode As String = oDv(0)(2).ToString.Trim
            sClinicCod = oDv(0)(1).ToString.Trim
            sIncuredMnth = oDv(0)(39).ToString.Trim

            Dim sApCode As String
            sApCode = "V" & sClinicCod & sSubCode

            dtCardCode.DefaultView.RowFilter = "CardCode = '" & sApCode & "'"
            If dtCardCode.DefaultView.Count = 0 Then
                sErrDesc = "Cardcode not found in SAP / Check Cardcode :: " & sApCode
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            End If

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
            oJournalEntry.Memo = "Estimated cost for clinic " & sApCode

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

                        For j As Integer = 0 To oDv.Count - 1
                            dCmoney = dCmoney + CDbl(oDv(j)(20).ToString.Trim)
                            dPayClient = dPayClient + CDbl(oDv(j)(16).ToString.Trim)
                            dOper = dOper + CDbl(oDv(j)(17).ToString.Trim)
                        Next

                        dCmoney = Math.Round(dCmoney, 2)
                        dPayClient = Math.Round(dPayClient, 2)
                        dOper = Math.Round(dOper, 2)

                        dCMoneyClient = Math.Round((dCmoney - dPayClient), 2)

                        If dCMoneyClient <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.AccountCode = sCMoneyClient
                            oJournalEntry.Lines.Debit = dCMoneyClient
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If
                        If dPayClient <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.AccountCode = sPayClntAct_Debit
                            oJournalEntry.Lines.Debit = dPayClient
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If
                        If dPayClient <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.ShortName = sPayClnt_Credit
                            oJournalEntry.Lines.Credit = dPayClient
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                oJournalEntry.Lines.CostingCode2 = sCostCenter
                            End If
                            iCount = iCount + 1
                            bIsLineAdded = True
                        End If
                        If dOper <> 0 Then
                            If iCount > 1 Then
                                oJournalEntry.Lines.Add()
                            End If
                            oJournalEntry.Lines.ShortName = sOperAct
                            oJournalEntry.Lines.Credit = dOper
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                oJournalEntry.Lines.CostingCode2 = sCostCenter
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
                            oJournalEntry.Lines.Credit = dTotval
                            If Not sCostCenter = String.Empty Then
                                oJournalEntry.Lines.CostingCode = sCostCenter
                                oJournalEntry.Lines.CostingCode2 = sCostCenter
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

                            sQuery = "UPDATE " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL"" SET ""U_Journal"" = '" & iDocNo & "', ""U_JrnlEntry"" = '" & iJournalEntryNo & "' " & _
                                     " WHERE ""U_cln_code"" = '" & sClinicCod & "' AND ""U_incurred_month"" = '" & sIncuredMnth & "' AND ""U_source"" = 'MS002' " & _
                                     " AND IFNULL(""U_JrnlEntry"",'') = '' AND ""U_invoice"" = '" & sInvoice & "' "
                            oRs.DoQuery(sQuery)

                        End If
                    Next

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournalEntry)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            PO_CreateJournalEntry_Capitation = RTN_SUCCESS

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            PO_CreateJournalEntry_Capitation = RTN_ERROR
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
