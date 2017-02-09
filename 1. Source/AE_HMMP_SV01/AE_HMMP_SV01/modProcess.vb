Module modProcess

#Region "Start"
    Public Sub Start()
        Dim sFuncName As String = "Start()"
        Dim sErrDesc As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Console.WriteLine("Reading Excel Values")

            UploadExcelFiles()

            'Send Error Email if Datable has rows.
            If p_oDtError.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Error()", sFuncName)
                EmailTemplate_Error()
            End If
            p_oDtError.Rows.Clear()

            'Send Success Email if Datable has rows..
            If p_oDtSuccess.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Success()", sFuncName)
                EmailTemplate_Success()
            End If
            p_oDtSuccess.Rows.Clear()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End
        End Try
    End Sub
#End Region
#Region "Upload Excel Files"
    Private Function UploadExcelFiles()
        Dim sFuncName As String = "UploadExcelFiles()"
        Dim sErrDesc As String = String.Empty
        Dim bIsFileExists As Boolean = False
        Dim oDVData As DataView = New DataView

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Excel Upload function Starts", sFuncName)

            p_oDtSuccess = CreateDataTable("FileName", "Status")
            p_oDtError = CreateDataTable("FileName", "Status", "ErrDesc")
            p_oDtReport = CreateDataTable("Type", "DocEntry", "BPCode", "Owner")

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim Files() As System.IO.FileInfo

            'First the program reads all the SO files and process it
            '*****SO****
            Files = DirInfo.GetFiles("*_SO_*.xls")

            For Each file As System.IO.FileInfo In Files
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File : " & file.Name.ToUpper)

                Dim sFileType As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)

                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & file.Name
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                Dim k As Integer = sFileName.IndexOf("_")
                sFileType = sFileName.Substring(k, Len(sFileName) - k)

                If sFileType.Contains("_SO_") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessSalesOrders()", sFuncName)
                        Console.WriteLine("Processing Sales order excel file " & sFileName)
                        If ProcessSalesOrders(file, oDVData, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                End If
            Next

            '*****AR *****
            Files = DirInfo.GetFiles("*_AR_*.xls")

            For Each file As System.IO.FileInfo In Files
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File : " & file.Name.ToUpper)

                Dim sFileType As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)

                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & file.Name
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                Dim k As Integer = sFileName.IndexOf("_")
                sFileType = sFileName.Substring(k, Len(sFileName) - k)
               
                If sFileType.ToUpper.Contains("_WRITEOFF") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessARInvoice()", sFuncName)
                        Console.WriteLine("Processing AR Invoice writeoff excel file " & sFileName)
                        If ProcessARInvoice_CardCode_Writeoff(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                ElseIf sFileType.Contains("_AR_") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessARInvoice()", sFuncName)
                        Console.WriteLine("Processing Sales Invoice excel file " & sFileName)
                        If ProcessARInvoice_CardCode(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                End If
            Next

            '**********AR(P)
            Files = DirInfo.GetFiles("*_AR(P)_*.xls")
            For Each file As System.IO.FileInfo In Files
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File : " & file.Name.ToUpper)

                Dim sFileType As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)
                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & file.Name
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                Dim k As Integer = sFileName.IndexOf("_")
                sFileType = sFileName.Substring(k, Len(sFileName) - k)

                If sFileType.ToUpper.Contains("_WRITEOFF") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessARInvoice_CardCode()", sFuncName)
                        Console.WriteLine("Processing Sales Invoice(P) excel file " & sFileName)
                        If ProcessARInvoice_Writeoff(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                ElseIf sFileType.Contains("_AR(P)_") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessARInvoice_CardCode()", sFuncName)
                        Console.WriteLine("Processing Sales Invoice(P) excel file " & sFileName)
                        If ProcessARInvoice(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                End If
            Next

            '*****PO*****
            Files = DirInfo.GetFiles("*_PO_*.xls")

            For Each file As System.IO.FileInfo In Files
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File : " & file.Name.ToUpper)

                Dim sFileType As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)

                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & file.Name
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                Dim k As Integer = sFileName.IndexOf("_")
                sFileType = sFileName.Substring(k, Len(sFileName) - k)

                If sFileType.Contains("_PO_") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessPurchaseOrder()", sFuncName)
                        Console.WriteLine("Processing Purchase order excel file " & sFileName)
                        If ProcessPurchaseOrder(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                End If
            Next

            '*****AP*****
            Files = DirInfo.GetFiles("*_AP_*.xls")

            For Each file As System.IO.FileInfo In Files
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File : " & file.Name.ToUpper)

                Dim sFileType As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)

                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & file.Name
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                Dim k As Integer = sFileName.IndexOf("_")
                sFileType = sFileName.Substring(k, Len(sFileName) - k)

                If sFileType.ToUpper.Contains("_WRITEOFF") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessAPInvoice()", sFuncName)
                        Console.WriteLine("Processing Purchase Invoice excel file " & sFileName)
                        If ProcessAPInvoice_WriteOff(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                ElseIf sFileType.Contains("_AP_") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessAPInvoice()", sFuncName)
                        Console.WriteLine("Processing Purchase Invoice excel file " & sFileName)
                        If ProcessAPInvoice(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                End If
            Next

            '****************YOT FILES
            Dim DirInfo_Yot As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir & "\YOT")

            Files = DirInfo_Yot.GetFiles("*_YOT_*.Xls")
            For Each file As System.IO.FileInfo In Files
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File : " & file.Name.ToUpper)

                Dim sFileType As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)

                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & file.Name
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                Dim k As Integer = sFileName.IndexOf("_")
                sFileType = sFileName.Substring(k, Len(sFileName) - k)

                If sFileType.Contains("_YOT_") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessYOTDatas()", sFuncName)
                        Console.WriteLine("Processing CMMS(YOT) excel file " & sFileName)
                        If ProcessYOTDatas(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                End If
            Next


            '***********************HMDC
            Dim DirInfo_HMDC As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir & "\HMDC")

            Files = DirInfo_HMDC.GetFiles("*_HMDC_*.xls")
            For Each file As System.IO.FileInfo In Files
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File : " & file.Name.ToUpper)

                Dim sFileType As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)

                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & file.Name
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                Dim k As Integer = sFileName.IndexOf("_")
                sFileType = sFileName.Substring(k, Len(sFileName) - k)

                If sFileType.Contains("_HMDC_") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling processHMDCDatas()", sFuncName)
                        Console.WriteLine("Processing CMMS(HMDC) excel file " & sFileName)
                        If processHMDCDatas(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                End If
            Next

            '**************HMMPD INTERFACE*****************
            Dim DirInfo_HMDMPD As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir & "\HMMPD")

            Files = DirInfo_HMDMPD.GetFiles("RD001_*.xls")
            For Each file As System.IO.FileInfo In Files
                bIsFileExists = True

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File : " & file.Name.ToUpper)

                Dim sFileType As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IsXLBookOpen()", sFuncName)

                If IsXLBookOpen(file.Name) = True Then
                    sErrDesc = "File is in use. Please close the document. File Name : " & file.Name
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug(sErrDesc, sFuncName)

                    'Insert Error Description into Table
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
                    AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)

                    Continue For
                End If

                Dim k As Integer = sFileName.IndexOf("_")
                Dim sFileN As String = String.Empty
                sFileN = file.Name
                'sFileType = sFileName.Substring(k, Len(sFileName) - k)
                'sFileType = sFileN.Substring(1, k)

                If sFileN.Contains("RD001_") Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read Excel file into Dataview", sFuncName)
                    oDVData = GetDataViewFromExcel(file.FullName, file.Extension)

                    If Not oDVData Is Nothing Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessHMMPDDatas()", sFuncName)
                        Console.WriteLine("Processing CMMS(HMMPD) excel file " & sFileName)
                        If ProcessHMMPDDatas(oDVData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in excel. File Name :" & file.Name, sFuncName)
                        Continue For
                    End If
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed successfully", sFuncName)
            UploadExcelFiles = RTN_SUCCESS
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while upload excel files", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
            UploadExcelFiles = RTN_ERROR
        End Try

    End Function
#End Region

End Module
