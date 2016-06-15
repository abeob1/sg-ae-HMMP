Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Data.Common
Imports System.Data.OleDb
Imports System.Xml
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Globalization

Module modCommon

#Region "Get Company  Initializaiton Info"
    Public Function GetInitializationInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "GetInitializationInfo()"
        Dim sConnection As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Company Initialization ", sFuncName)

            oCompDef.sServer = String.Empty
            oCompDef.sLicenseServer = String.Empty
            oCompDef.sSAPUser = String.Empty
            oCompDef.sSAPPwd = String.Empty
            oCompDef.sSAPDBName = String.Empty
            oCompDef.sDBUser = String.Empty
            oCompDef.sDBPwd = String.Empty

            oCompDef.sInboxDir = String.Empty
            oCompDef.sSuccessDir = String.Empty
            oCompDef.sFailDir = String.Empty
            oCompDef.sLogPath = String.Empty

            oCompDef.sEmailFrom = String.Empty
            oCompDef.sEmailTo = String.Empty
            oCompDef.sEmailSubject = String.Empty
            oCompDef.sSMTPServer = String.Empty
            oCompDef.sSMTPPort = String.Empty
            oCompDef.sSMTPUser = String.Empty
            oCompDef.sSMTPPassword = String.Empty
            oCompDef.sCOAcrlCardCode = String.Empty
            oCompDef.sARInvFooter = String.Empty
            oCompDef.sType = String.Empty

            oCompDef.sYOTSAPDbName = String.Empty
            oCompDef.sYOTSAPUserName = String.Empty
            oCompDef.sYOTSAPPassword = String.Empty
            oCompDef.sYOTARInvPayCardcode = String.Empty

            oCompDef.sHMDCSAPDbName = String.Empty
            oCompDef.sHMDCSAPUserName = String.Empty
            oCompDef.sHMDCSAPPassword = String.Empty
            oCompDef.sHMDCARInvPayCardCode = String.Empty

            oCompDef.sHMMPDSAPDbName = String.Empty
            oCompDef.sHMMPDSAPUserName = String.Empty
            oCompDef.sHMMPDSAPPassword = String.Empty

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LicenceServer")) Then
                oCompDef.sLicenseServer = ConfigurationManager.AppSettings("LicenceServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPDBName")) Then
                oCompDef.sSAPDBName = ConfigurationManager.AppSettings("SAPDBName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPUserName")) Then
                oCompDef.sSAPUser = ConfigurationManager.AppSettings("SAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SAPPassword")) Then
                oCompDef.sSAPPwd = ConfigurationManager.AppSettings("SAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("InboxDir")) Then
                oCompDef.sInboxDir = ConfigurationManager.AppSettings("InboxDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SuccessDir")) Then
                oCompDef.sSuccessDir = ConfigurationManager.AppSettings("SuccessDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("FailDir")) Then
                oCompDef.sFailDir = ConfigurationManager.AppSettings("FailDir")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("LogPath")) Then
                oCompDef.sLogPath = ConfigurationManager.AppSettings("LogPath")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailFrom")) Then
                oCompDef.sEmailFrom = ConfigurationManager.AppSettings("EmailFrom")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailTo")) Then
                oCompDef.sEmailTo = ConfigurationManager.AppSettings("EmailTo")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailSubject")) Then
                oCompDef.sEmailSubject = ConfigurationManager.AppSettings("EmailSubject")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPServer")) Then
                oCompDef.sSMTPServer = ConfigurationManager.AppSettings("SMTPServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPort")) Then
                oCompDef.sSMTPPort = ConfigurationManager.AppSettings("SMTPPort")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPUser")) Then
                oCompDef.sSMTPUser = ConfigurationManager.AppSettings("SMTPUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPassword")) Then
                oCompDef.sSMTPPassword = ConfigurationManager.AppSettings("SMTPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("COAcrlCardCode")) Then
                oCompDef.sCOAcrlCardCode = ConfigurationManager.AppSettings("COAcrlCardCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("ARInvoiceFooter")) Then
                oCompDef.sARInvFooter = ConfigurationManager.AppSettings("ARInvoiceFooter")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Type")) Then
                oCompDef.sType = ConfigurationManager.AppSettings("Type")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("YOTSAPDbName")) Then
                oCompDef.sYOTSAPDbName = ConfigurationManager.AppSettings("YOTSAPDbName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("YOTSAPUserName")) Then
                oCompDef.sYOTSAPUserName = ConfigurationManager.AppSettings("YOTSAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("YOTSAPPassword")) Then
                oCompDef.sYOTSAPPassword = ConfigurationManager.AppSettings("YOTSAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("YOTARInvPayCardCode")) Then
                oCompDef.sYOTARInvPayCardcode = ConfigurationManager.AppSettings("YOTARInvPayCardCode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("HMDCSAPDbName")) Then
                oCompDef.sHMDCSAPDbName = ConfigurationManager.AppSettings("HMDCSAPDbName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("HMDCSAPUserName")) Then
                oCompDef.sHMDCSAPUserName = ConfigurationManager.AppSettings("HMDCSAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("HMDCSAPPassword")) Then
                oCompDef.sHMDCSAPPassword = ConfigurationManager.AppSettings("HMDCSAPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("HMDCARInvPayCardcode")) Then
                oCompDef.sHMDCARInvPayCardCode = ConfigurationManager.AppSettings("HMDCARInvPayCardcode")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("HMMPDSAPDbName")) Then
                oCompDef.sHMMPDSAPDbName = ConfigurationManager.AppSettings("HMMPDSAPDbName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("HMMPDSAPUserName")) Then
                oCompDef.sHMMPDSAPUserName = ConfigurationManager.AppSettings("HMMPDSAPUserName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("HMMPDSAPPassword")) Then
                oCompDef.sHMMPDSAPPassword = ConfigurationManager.AppSettings("HMMPDSAPPassword")
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFuncName)
            GetInitializationInfo = RTN_SUCCESS

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
            GetInitializationInfo = RTN_ERROR
        End Try

    End Function
#End Region
#Region "Connect to Company"
    Public Function ConnectToCompany(ByRef oCompany As SAPbobsCOM.Company, ByVal sDBName As String, ByVal sDBUser As String, ByVal sPassword As String, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   ConnectToCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   October 2013
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Try
            sFuncName = "ConnectToCompany()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
            oCompany = New SAPbobsCOM.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)

            oCompany.Server = p_oCompDef.sServer
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            oCompany.CompanyDB = sDBName
            oCompany.UserName = sDBUser
            oCompany.Password = sPassword

            oCompany.LicenseServer = p_oCompDef.sLicenseServer

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            oCompany.UseTrusted = False

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
            iRetValue = oCompany.Connect()

            If iRetValue <> 0 Then
                oCompany.GetLastError(iErrCode, sErrDesc)

                sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                    oCompany.CompanyDB, System.Environment.NewLine, _
                                vbTab, sErrDesc)

                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConnectToCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectToCompany = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Connect to Other Company"
    Public Function ConnectAnotherCompany(ByRef oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String, Optional ByVal sDBName As String = "") As Long
        ' **********************************************************************************
        '   Function    :   ConnectAnotherCompany()
        '   Purpose     :   This function will be providing to proceed the connectivity of 
        '                   using SAP DIAPI function
        '               
        '   Parameters  :   ByRef oCompany As SAPbobsCOM.Company
        '                       oCompany =  set the SAP DI Company Object
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author      :   SRI
        '   Date        :   October 2013
        ' **********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim iRetValue As Integer = -1
        Dim iErrCode As Integer = -1
        Try
            sFuncName = "ConnectAnotherCompany()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing the Company Object", sFuncName)
            oCompany = New SAPbobsCOM.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning the representing database name", sFuncName)

            oCompany.Server = p_oCompDef.sServer
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            oCompany.CompanyDB = p_oCompDef.sHMMPDSAPDbName
            oCompany.UserName = p_oCompDef.sHMMPDSAPUserName
            oCompany.Password = p_oCompDef.sHMMPDSAPPassword

            oCompany.LicenseServer = p_oCompDef.sLicenseServer

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English

            oCompany.UseTrusted = False

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Connecting to the Company Database.", sFuncName)
            iRetValue = oCompany.Connect()

            If iRetValue <> 0 Then
                oCompany.GetLastError(iErrCode, sErrDesc)

                sErrDesc = String.Format("Connection to Database ({0}) {1} {2} {3}", _
                    oCompany.CompanyDB, System.Environment.NewLine, _
                                vbTab, sErrDesc)

                Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ConnectAnotherCompany = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ConnectAnotherCompany = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Create Datatable"
    Public Function CreateDataTable(ByVal ParamArray oColumnName() As String) As DataTable
        Dim oDataTable As DataTable = New DataTable()

        Dim oDataColumn As DataColumn

        For i As Integer = LBound(oColumnName) To UBound(oColumnName)
            oDataColumn = New DataColumn()
            oDataColumn.DataType = Type.GetType("System.String")
            oDataColumn.ColumnName = oColumnName(i).ToString
            oDataTable.Columns.Add(oDataColumn)
        Next

        Return oDataTable

    End Function
#End Region
#Region "Add data to table"
    Public Sub AddDataToTable(ByVal oDt As DataTable, ByVal ParamArray sColumnValue() As String)
        Dim oRow As DataRow = Nothing
        oRow = oDt.NewRow()
        For i As Integer = LBound(sColumnValue) To UBound(sColumnValue)
            oRow(i) = sColumnValue(i).ToString
        Next
        oDt.Rows.Add(oRow)
    End Sub
#End Region
#Region "Check Excel file open or not"
    Public Function IsXLBookOpen(strName As String) As Boolean

        'Function designed to test if a specific Excel
        'workbook is open or not.
        Dim i As Long
        Dim XLAppFx As Excel.Application
        Dim NotOpen As Boolean
        Dim sFuncName As String = "IsXLBookOpen"


        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

        'Find/create an Excel instance
        On Error Resume Next
        XLAppFx = GetObject(, "Excel.Application")
        If Err.Number = 429 Then
            NotOpen = True
            XLAppFx = CreateObject("Excel.Application")
            Err.Clear()
        End If

        'Loop through all open workbooks in such instance

        For i = XLAppFx.Workbooks.Count To 1 Step -1

            If XLAppFx.Workbooks(i).Name = strName Then
                'Perform check to see if name was found
                IsXLBookOpen = True
                Exit For
            Else
                'Set all to False
                IsXLBookOpen = False
            End If
        Next i

        'Close if was closed
        If NotOpen Then XLAppFx.Quit()

        'Release the instance
        XLAppFx = Nothing
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

    End Function
#End Region
#Region "Get View from Excel into Dataview"
    Public Function GetDataViewFromExcel(ByVal CurrFileToUpload As String, ByVal sExtension As String) As DataView

        Dim conStr As String = ""
        Dim sFuncName As String = String.Empty
        Dim dv As DataView

        Try
            sFuncName = "GetDataViewFromExcel"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Select Case sExtension
                Case ".xls"
                    'Excel 97-03
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'"
                    Exit Select
                Case ".xlsx"
                    'Excel 07
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrFileToUpload & ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1'"
                    Exit Select
            End Select

            Dim connExcel As New OleDbConnection(conStr)
            Dim cmdExcel As New OleDbCommand()
            Dim oda As New OleDbDataAdapter()
            Dim dt As New DataTable()

            cmdExcel.Connection = connExcel

            'Get the name of First Sheet
            connExcel.Open()
            Dim dtExcelSchema As DataTable
            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim SheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
            connExcel.Close()

            'Read Data from First Sheet
            connExcel.Open()
            cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
            oda.SelectCommand = cmdExcel
            dt = New DataTable("Data")
            oda.Fill(dt)
            connExcel.Close()

            dv = New DataView(dt)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFuncName)

            Return dv


        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while reading content of " & ex.Message, sFuncName)
            Call WriteToLogFile_Debug(ex.Message, sFuncName)
            Return Nothing
        End Try


    End Function
#End Region
#Region "Move File to Archive"
    Public Sub FileMoveToArchive(ByVal oFile As System.IO.FileInfo, ByVal CurrFileToUpload As String, ByVal iStatus As Integer)

        'Event      :   FileMoveToArchive
        'Purpose    :   For Renaming the file with current time stamp & moving to archive folder
        'Author     :   SRI 
        'Date       :   24 NOV 2013

        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "FileMoveToArchive"

            'Dim RenameCurrFileToUpload = Replace(CurrFileToUpload.ToUpper, ".CSV", "") & "_" & Format(Now, "yyyyMMddHHmmss") & ".csv"
            Dim RenameCurrFileToUpload As String = Mid(oFile.Name, 1, oFile.Name.Length - 4) & "_" & Now.ToString("yyyyMMddhhmmss") & ".xls"

            If iStatus = RTN_SUCCESS Then
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving Excel file to success folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sSuccessDir & "\" & RenameCurrFileToUpload)
            Else
                If p_iDebugMode = DEBUG_ON Then WriteToLogFile_Debug("Moving Excel file to Fail folder", sFuncName)
                oFile.MoveTo(p_oCompDef.sFailDir & "\" & RenameCurrFileToUpload)
            End If
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in renaming/copying/moving", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Sub
#End Region
#Region "Start Transaction"
    Public Function StartTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    StartTransaction()
        '   Purpose    :    Start DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :   Jeeva
        '   Date       :   03 Aug 2015
        '   Change     :
        ' ***********************************************************************************

        Dim sFuncName As String = "StartTransaction"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Transaction", sFuncName)

            If p_oCompany.InTransaction Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback hanging transactions", sFuncName)
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            p_oCompany.StartTransaction()

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Trancation Started Successfully", sFuncName)
            StartTransaction = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile_Debug(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while starting Trancation", sFuncName)
            StartTransaction = RTN_ERROR
        End Try

    End Function
#End Region
#Region "Commit Transaction"
    Public Function CommitTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    CommitTransaction()
        '   Purpose    :    Commit DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc=Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    Jeeva
        '   Date       :    03 Aug 2015
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = "CommitTransaction"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If p_oCompany.InTransaction Then
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Transaction is Active", sFuncName)
            End If

            CommitTransaction = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Commit Transaction Complete", sFuncName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while committing Transaciton", sFuncName)
            CommitTransaction = RTN_ERROR
        End Try
    End Function
#End Region
#Region "Rollback Transaction"
    Public Function RollbackTransaction(ByRef sErrDesc As String) As Long
        ' ***********************************************************************************
        '   Function   :    RollbackTransaction()
        '   Purpose    :    Rollback DI Company Transaction
        '
        '   Parameters :    ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :   0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :   Jeeva
        '   Date       :   31 July 2015
        '   Change     :
        ' ***********************************************************************************
        Dim sFuncName As String = String.Empty

        Try
            sFuncName = "RollbackTransaction()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_oCompany.InTransaction Then
                p_oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No transaction is active", sFuncName)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Success", sFuncName)
            RollbackTransaction = RTN_SUCCESS
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
            RollbackTransaction = RTN_ERROR
        End Try

    End Function
#End Region
#Region "Execute SQL Query"

    Public Function ExecuteQueryReturnDataTable(ByVal sQueryString As String, ByVal sCompanyDB As String) As DataTable

        Dim sFuncName As String = "ExecuteQueryReturnDataTable"
        'Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & sCompanyDB & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd & ""
        Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & sCompanyDB

        Dim oCmd As New Odbc.OdbcCommand
        Dim oDS As DataSet = New DataSet
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()
        Dim dtDetail As DataTable = New DataTable

        ''SQL CODES
        'Dim oCon As SqlConnection
        'Dim oSQLAdapter As SqlDataAdapter

        Try
            Con.ConnectionString = sConstr
            Con.Open()

            oCmd.CommandText = CommandType.Text
            oCmd.CommandText = sQueryString
            oCmd.Connection = Con
            oCmd.CommandTimeout = 0

            Dim da As New Odbc.OdbcDataAdapter(oCmd)
            da.Fill(dtDetail)
            dtDetail.TableName = "Data"

            'oCmd.CommandType = CommandType.Text
            'oCmd.CommandText = sQueryString
            'oCmd.Connection = oCon
            'If oCon.State = ConnectionState.Closed Then
            '    oCon.Open()
            'End If

            'oSQLAdapter.SelectCommand = oCmd

            'oSQLAdapter.Fill(dtDetail)
            'dtDetail.TableName = "Data"

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try

        ExecuteQueryReturnDataTable = dtDetail

    End Function

    Public Function ExecuteQueryReturnDataTable_AnotherDB(ByVal sQueryString As String, ByVal sCompanyDB As String) As DataTable

        Dim sFuncName As String = "ExecuteQueryReturnDataTable"
        'Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & sCompanyDB & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd & ""
        Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & sCompanyDB

        Dim oCmd As New Odbc.OdbcCommand
        Dim oDS As DataSet = New DataSet
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()
        Dim dtDetail As DataTable = New DataTable

        Try
            Con.ConnectionString = sConstr
            Con.Open()

            oCmd.CommandText = CommandType.Text
            oCmd.CommandText = sQueryString
            oCmd.Connection = Con
            oCmd.CommandTimeout = 0

            Dim da As New Odbc.OdbcDataAdapter(oCmd)
            da.Fill(dtDetail)
            dtDetail.TableName = "Data"

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try

        ExecuteQueryReturnDataTable_AnotherDB = dtDetail

    End Function

    Public Function ExecuteSQLQuery(ByVal sQuery As String) As DataSet

        '**************************************************************
        ' Function      : ExecuteSQLQuery
        ' Purpose       : Execute SQL
        ' Parameters    : ByVal sSQL - string command Text
        ' Author        : Sri
        ' Date          : 
        ' Change        :
        '**************************************************************

        Dim sFuncName As String = String.Empty

        ' Dim sConstr As String = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sSAPDBName
        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd
        ''Dim oCon As New Odbc.OdbcConnection(sConstr)
        ''Dim oCmd As New Odbc.OdbcCommand
        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet

        Try
            sFuncName = "ExecuteQuery()"
            oCon.ConnectionString = sConstr
            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDs)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)

        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while executing query", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
        Return oDs
    End Function

    Public Function ExecuteSQLQuery_Hana(ByVal sSql As String, ByVal sDbName As String) As DataSet
        Dim sFuncName As String = "ExecuteSQLQuery_Hana"
        Dim sErrDesc As String = String.Empty

        Dim cmd As New Odbc.OdbcCommand
        Dim ods As New DataSet
        'Dim oSQLCommand As SqlCommand = Nothing
        'Dim oSQLAdapter As New SqlDataAdapter
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()
        'Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd & ""

        Try

            Con.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & sDbName
            Con.Open()

            cmd.CommandType = CommandType.Text
            cmd.CommandText = sSql
            cmd.Connection = Con
            cmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(cmd)
            da.Fill(ods)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try
        Return ods

    End Function

    Public Function ExecuteSQLQuery_Hana_AnotherDb(ByVal sSql As String) As DataSet
        Dim sFuncName As String = "ExecuteSQLQuery_Hana_AnotherDb"
        Dim sErrDesc As String = String.Empty

        Dim cmd As New Odbc.OdbcCommand
        Dim ods As New DataSet
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()

        Try

            Con.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sHMMPDSAPUserName & ";PWD=" & p_oCompDef.sHMMPDSAPPassword & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & p_oCompDef.sHMMPDSAPDbName
            Con.Open()

            cmd.CommandType = CommandType.Text
            cmd.CommandText = sSql
            cmd.Connection = Con
            cmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(cmd)
            da.Fill(ods)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try
        Return ods

    End Function

    Public Function ExecuteTargetCompSQLQuery(ByVal sSql As String, ByVal sCompanyDB As String) As DataSet
        Dim sFuncName As String = "ExecuteSQLQuery"
        Dim sErrDesc As String = String.Empty

        Dim cmd As New Odbc.OdbcCommand
        Dim ods As New DataSet
        'Dim oSQLCommand As SqlCommand = Nothing
        'Dim oSQLAdapter As New SqlDataAdapter
        Dim oDbProviderFactoryObj As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.Odbc")
        Dim Con As DbConnection = oDbProviderFactoryObj.CreateConnection()

        Try

            Con.ConnectionString = "DRIVER={HDBODBC32};UID=" & p_oCompDef.sDBUser & ";PWD=" & p_oCompDef.sDBPwd & ";SERVERNODE=" & p_oCompDef.sServer & ";CS=" & sCompanyDB
            Con.Open()

            cmd.CommandType = CommandType.Text
            cmd.CommandText = sSql
            cmd.Connection = Con
            cmd.CommandTimeout = 0
            Dim da As New Odbc.OdbcDataAdapter(cmd)
            da.Fill(ods)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ExecuteSQL Query Error", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            Con.Dispose()
        End Try
        Return ods
    End Function

#End Region
#Region "Get Cost Center from MDMS table"
    Public Function GetCostCenter(ByRef sDBCode As String, ByVal dTxnDate As Date, ByVal sSchemeCode As String, ByVal sDBName As String) As String
        Dim sFuncName As String = "GetCostCenter"
        Dim sSql As String
        Dim oDs As DataSet
        Dim sCostCenter As String = String.Empty

        'sSql = "SELECT TOP 1 ""U_MBMS"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MBMS"" WHERE ""U_company_code"" = '" & sDBCode & "' AND ""U_Effective_Date"" <= '" & dTxnDate.ToString("yyyy-MM-dd") & "' " & _
        '       " AND ""U_Scheme"" = '" & sSchemeCode & "' ORDER BY ""U_Effective_Date"" DESC "
        sSql = "SELECT B.""PrcCode"" AS ""U_MBMS"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MBMS"" A " & _
               " INNER JOIN " & p_oCompDef.sSAPDBName & ".""OPRC"" B ON UPPER(B.""PrcCode"") = UPPER(A.""U_MBMS"") " & _
               " WHERE A.""U_company_code"" = '" & sDBCode & "' AND A.""U_Effective_Date"" <= '" & dTxnDate.ToString("yyyy-MM-dd") & "' " & _
               " AND A.""U_Scheme"" = '" & sSchemeCode & "' ORDER BY A.""U_Effective_Date"" DESC "

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        oDs = ExecuteSQLQuery_Hana(sSql, sDBName)

        If oDs.Tables(0).Rows.Count > 0 Then
            sCostCenter = oDs.Tables(0).Rows(0).Item("U_MBMS").ToString
        End If

        Return sCostCenter
    End Function
#End Region
#Region "Get Insurer from MDMS table"
    Public Function GetInsurer(ByRef sDBCode As String, ByVal dTxnDate As Date, ByVal sSchemeCode As String, ByVal sDBName As String) As String
        Dim sFuncName As String = "GetInsurer"
        Dim sSql As String
        Dim oDs As DataSet
        Dim sInsurer As String = String.Empty

        sSql = "SELECT TOP 1 ""U_Insurer"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MBMS"" WHERE ""U_company_code"" = '" & sDBCode & "' AND ""U_Effective_Date"" <= '" & dTxnDate.ToString("yyyy-MM-dd") & "' " & _
               " AND ""U_Scheme"" = '" & sSchemeCode & "' ORDER BY ""U_Effective_Date"" DESC "
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        oDs = ExecuteSQLQuery_Hana(sSql, sDBName)

        If oDs.Tables(0).Rows.Count > 0 Then
            sInsurer = oDs.Tables(0).Rows(0).Item("U_Insurer").ToString
        End If

        Return sInsurer
    End Function
#End Region
#Region "get Interger value"
    Public Function GetCode(ByVal sSql As String, ByVal sDBName As String) As Integer
        Dim sFuncName As String = "GetCode"
        Dim oDs As DataSet
        Dim sOutput As Integer = 0

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        oDs = ExecuteSQLQuery_Hana(sSql, sDBName)

        If oDs.Tables(0).Rows.Count > 0 Then
            sOutput = oDs.Tables(0).Rows(0).Item(0).ToString
        End If

        Return sOutput
    End Function
#End Region
#Region "Get Single string value"
    Public Function GetStringValue(ByVal sSql As String, ByVal sDbName As String) As String
        Dim sFuncName As String = "GetItemCode"
        Dim oDs As DataSet
        Dim sValue As String = String.Empty

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        oDs = ExecuteSQLQuery_Hana(sSql, sDbName)

        If oDs.Tables(0).Rows.Count > 0 Then
            sValue = oDs.Tables(0).Rows(0).Item(0).ToString
        End If

        Return sValue
    End Function
#End Region
#Region "Get Amount value"
    Public Function getAmt(ByVal sSql As String, ByVal sDBName As String, ByVal sUser As String, ByVal sPwd As String) As Double
        Dim sFuncName As String = "getAmt"
        Dim oDs As DataSet
        Dim dAmt As Double = 0

        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        oDs = ExecuteSQLQuery_Hana(sSql, sDBName)

        If oDs.Tables(0).Rows.Count > 0 Then
            dAmt = oDs.Tables(0).Rows(0).Item(0).ToString
        End If

        Return dAmt
    End Function
#End Region
#Region "Insert datas into Sales Accural table"
    Public Function InsertIntoSOAccural(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "InsertIntoSOAccural"
        Dim sSql As String = String.Empty
        Dim sType As String = String.Empty

        Try
            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim dPayComp, dTotPayComp As Double
            For i As Integer = 0 To oDv.Count - 1
                dPayComp = CDbl(oDv(i)(20).ToString.Trim)
                dTotPayComp = dTotPayComp + dPayComp
            Next

            Dim sCompCode As String = oDv(0)(1).ToString.Trim
            Dim sArCode As String = "C" & oDv(0)(1).ToString.Trim

            sSql = "SELECT ""U_Type"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" WHERE ""CardCode"" = '" & sArCode & "'"
            oRecSet.DoQuery(sSql)
            If oRecSet.RecordCount > 0 Then
                sType = oRecSet.Fields.Item("U_Type").Value
            End If

            If sType.ToUpper <> "CAPITATION" Then
                sSql = "INSERT INTO " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL""(""Code"",""Name"",""U_company_code"",""U_ar_code"",""U_Incurred_month"",""U_OcrCode"",""U_Insurer"",""U_invoice"",""U_total_sales"",""U_status"",""U_Type"")"
                sSql = sSql & " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sSAPDBName & ".""@AE_SOACCRUAL""), "
                sSql = sSql & " '" & sCompCode & "','" & sArCode & "','" & oDv(0)(50).ToString & "','" & oDv(0)(48).ToString & "','" & oDv(0)(49).ToString & "',"
                sSql = sSql & " '" & oDv(0)(17).ToString & "','" & dTotPayComp & "','O','" & sType & "')"

                oRecSet.DoQuery(sSql)
            End If
            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoSOAccural = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while executing query", sFuncName)
            InsertIntoSOAccural = RTN_SUCCESS
            Throw New Exception(ex.Message)
        End Try
    End Function
#End Region
#Region "SQL Transaction queries"
    Public Sub SQLTransaction(ByVal sSQL As String)
        Dim sFuncName As String = "SQLTransaction"
        Dim sConstr As String = "Provider=SQLOLEDB;Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sSAPDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd
        Dim oCon As OleDb.OleDbConnection

        Try
            oCon = New OleDb.OleDbConnection(sConstr)
            oCon.Open()

            Dim dbc As OleDbCommand = oCon.CreateCommand()
            dbc.CommandText = sSQL
            dbc.ExecuteNonQuery()
            dbc.Dispose()
        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while executing query", sFuncName)
            Throw New Exception(ex.Message)
        Finally
            oCon.Dispose()
        End Try
    End Sub
#End Region
#Region "Insert into Cost Accrual table"
    Public Function InsertIntoCostAccrual(ByVal oDv As DataView, ByVal oCompany As SAPbobsCOM.Company, ByVal sSource As String, ByVal sErrDesc As String) As Long
        Dim sFuncName As String = "InsertIntoCostAccrual"
        Dim sSql As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sSource = sSource.Substring(0, 5)

            'Dim sClnCode As String = p_oCompDef.sCOAcrlCardCode
            Dim sClnCode As String = oDv(0)(4).ToString.Trim()

            Dim dPayComp, dTotPayComp As Double
            For i As Integer = 0 To oDv.Count - 1
                dPayComp = CDbl(oDv(i)(20).ToString.Trim)
                dTotPayComp = dTotPayComp + dPayComp
            Next

            Dim oRecSet As SAPbobsCOM.Recordset
            oRecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSql = "INSERT INTO " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL""(""Code"",""Name"",""U_cln_code"",""U_ap_code"",""U_incurred_month"",""U_OcrCode"",""U_Insurer"",""U_invoice"",""U_pay_comp"",""U_source"",""U_status"",""U_Type"")"
            sSql = sSql & " VALUES((SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL""),(SELECT IFNULL(MAX(TO_INT(""Code"")),0) + 1 FROM " & p_oCompDef.sSAPDBName & ".""@AE_COSTACCRUAL""),"
            sSql = sSql & " '" & sClnCode & "','" & sClnCode & "','" & oDv(0)(50).ToString & "','" & oDv(0)(48).ToString & "','" & oDv(0)(49).ToString & "',"
            sSql = sSql & " '" & oDv(0)(17).ToString & "','" & dTotPayComp & "','MS007','O','" & oDv(0)(51).ToString & "')"

            oRecSet.DoQuery(sSql)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            InsertIntoCostAccrual = RTN_SUCCESS

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error while executing query", sFuncName)
            InsertIntoCostAccrual = RTN_SUCCESS
            Throw New Exception(ex.Message)
        End Try

    End Function
#End Region
#Region "Get Account Code for Reversal of estimates Sales GL"
    Public Function GetActCode_RevEstiSaleGL(ByVal sFileCode As String, ByVal sAcctType As String) As String
        Dim sFuncName As String = "GetActCode_CostAccrual"
        Dim sSql As String
        Dim oDs As DataSet
        Dim sAcctcode As String = String.Empty

        sSql = "SELECT B.""AcctCode"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_MS007_GL_REV"" A INNER JOIN OACT B ON B.""FormatCode"" = A.""U_GLCode"" "
        sSql = sSql & " WHERE A.""U_FileCode"" = '" & sFileCode & "' AND A.""U_ActType"" = '" & sAcctType & "'"
        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing SQL" & sSql, sFuncName)

        oDs = ExecuteSQLQuery_Hana(sSql, p_oCompDef.sSAPDBName)

        If oDs.Tables(0).Rows.Count > 0 Then
            sAcctcode = oDs.Tables(0).Rows(0).Item("AcctCode").ToString
        End If

        Return sAcctcode
    End Function
#End Region

End Module
