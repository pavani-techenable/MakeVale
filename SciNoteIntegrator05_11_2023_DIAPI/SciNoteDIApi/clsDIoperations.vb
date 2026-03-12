Imports SAPbobsCOM
Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient
Public Class clsDIoperations
    Dim RetValue As Long
    Dim ErrCode As Long
    Dim ErrMsg As String = ""
    Dim lastkey As String = ""
    Dim InvoiceEntry As String = ""
    Dim DelEntry As String = ""
    Dim BatchonSec As String = ""
    Dim InvLineNum As String = ""
    Dim Transferentry As String = ""

    Public Sub createSalesInvoice(ByVal objSalesInvoice As SalesInvoiceHeaderData, ByRef errors As List(Of String), ByRef BubbleEvent As Boolean, ByRef DocEntry As Integer)
        Try
            'objCompany.StartTransaction()
            Dim SaleInvoice As SAPbobsCOM.Documents
            SaleInvoice = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            SaleInvoice.DocDate = objSalesInvoice.DocDate
            SaleInvoice.TaxDate = objSalesInvoice.TaxDate
            SaleInvoice.CardCode = objSalesInvoice.CardCode

            SaleInvoice.UserFields.Fields.Item("U_acode").Value = objSalesInvoice.Agent
            SaleInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
            For Each ObjDetailData As SalesInvoiceDetailData In objSalesInvoice.DetailData
                SaleInvoice.Lines.ItemCode = ObjDetailData.ItemCode
                SaleInvoice.Lines.Quantity = ObjDetailData.Quantity
                SaleInvoice.Lines.Price = ObjDetailData.Price
                SaleInvoice.Lines.TaxCode = ObjDetailData.TaxCode
                SaleInvoice.Lines.WarehouseCode = ObjDetailData.WhsCode
                For Each objBatchDetails As SalesInvoiceBatchData In ObjDetailData.BatchData
                    SaleInvoice.Lines.BatchNumbers.BatchNumber = objBatchDetails.BatchNumber
                    SaleInvoice.Lines.BatchNumbers.Quantity = objBatchDetails.Quantity
                    SaleInvoice.Lines.BatchNumbers.Add()
                Next
                SaleInvoice.Lines.Add()
            Next
            For Each objExpenseData As SalesInvoiceExpenseData In objSalesInvoice.ExpenseData
                If objExpenseData.ExpenseCode <> "" Then
                    SaleInvoice.Expenses.ExpenseCode = objExpenseData.ExpenseCode
                    SaleInvoice.Expenses.LineTotal = objExpenseData.LineTotal
                    SaleInvoice.Expenses.TaxCode = objExpenseData.TaxCode
                    SaleInvoice.Expenses.Add()
                End If

            Next
            RetValue = SaleInvoice.Add()
            If RetValue <> 0 Then
                objCompany.GetLastError(ErrCode, ErrMsg)
                errors.Add(ErrCode.ToString + ": " + ErrMsg)
                BubbleEvent = False
            Else
                DocEntry = objCompany.GetNewObjectKey()
            End If
        Catch ex As Exception
            errors.Add(ErrCode + ": " + ErrMsg)
            BubbleEvent = False
        End Try

    End Sub


    Public Sub createPurchaseOrder(ByVal objSalesInvoice As SalesInvoiceHeaderData, ByRef errors As List(Of String), ByRef BubbleEvent As Boolean, ByRef DocEntry As Integer)
        Try
            Dim PurchaseOrder As SAPbobsCOM.Documents
            PurchaseOrder = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            PurchaseOrder.DocDate = objSalesInvoice.DocDate
            '  PurchaseOrder.TaxDate = objSalesInvoice.TaxDate
            PurchaseOrder.CardCode = objSalesInvoice.CardCode

            PurchaseOrder.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
            For Each ObjDetailData As SalesInvoiceDetailData In objSalesInvoice.DetailData
                PurchaseOrder.Lines.ItemCode = ObjDetailData.ItemCode
                PurchaseOrder.Lines.Quantity = ObjDetailData.Quantity
                PurchaseOrder.Lines.Price = ObjDetailData.Price
                ' PurchaseOrder.Lines.TaxCode = ObjDetailData.TaxCode
                If ObjDetailData.WhsCode <> "" Then
                    PurchaseOrder.Lines.WarehouseCode = ObjDetailData.WhsCode
                End If
                PurchaseOrder.Lines.UserFields.Fields.Item("U_SO_DocEntry").Value = ObjDetailData.SO_DocEntry
                PurchaseOrder.Lines.UserFields.Fields.Item("U_SO_Linenum").Value = ObjDetailData.SO_LineNum
                PurchaseOrder.Lines.Add()
            Next

            RetValue = PurchaseOrder.Add()
            If RetValue <> 0 Then
                objCompany.GetLastError(ErrCode, ErrMsg)
                errors.Add(ErrCode.ToString + ": " + ErrMsg)
                BubbleEvent = False
            Else
                DocEntry = objCompany.GetNewObjectKey()
            End If
        Catch ex As Exception
            errors.Add(ErrCode + ": " + ErrMsg)
            BubbleEvent = False
        End Try

    End Sub
    Public Sub readFromXML()
        Try
            Dim xmldoc As New XmlDocument
            Dim fs As New FileStream("config.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)
            Dim nodes As XmlNodeList = xmldoc.DocumentElement.SelectNodes("/Data")
            For Each node As XmlNode In nodes
                server = node.SelectSingleNode("Server").InnerText
                Database = node.SelectSingleNode("Database").InnerText
                dbName = node.SelectSingleNode("DBName").InnerText
                dbUser = node.SelectSingleNode("DbUser").InnerText
                dbPassword = node.SelectSingleNode("DbPassword").InnerText
                sapUser = node.SelectSingleNode("sapUser").InnerText
                sapPassword = node.SelectSingleNode("sapPassword").InnerText
                attPath = node.SelectSingleNode("attachPath").InnerText
                attachPathLog = node.SelectSingleNode("attachPathLog").InnerText
            Next
        Catch ex As Exception
            InsertMailErrorLog("Connection", ex.Message(), "Server", "")
        End Try


    End Sub


    Public Sub createConnectionToCompany_()
        Try
            'Dim Server
            Dim lErrCode, sErrMsg
            objCompany = New SAPbobsCOM.Company
            objCompany.DbServerType = BoDataServerTypes.dst_MSSQL2012 Or BoDataServerTypes.dst_MSSQL2016 Or BoDataServerTypes.dst_MSSQL2019
            objCompany.Server = server
            objCompany.language = BoSuppLangs.ln_English
            objCompany.UseTrusted = False
            objCompany.DbUserName = dbUser
            objCompany.DbPassword = dbPassword
            objCompany.CompanyDB = dbName
            objCompany.UserName = sapUser
            objCompany.Password = sapPassword
            writeLog(String.Format("Server Details Server : {0} User : {1} Password : {2} CompanyDb : {3} sapUser : {4} Pass : {5}", server, dbUser, dbPassword, dbName, sapUser, sapPassword), True)
            '// Connecting to a company DB
            Dim lRetCode = objCompany.Connect

            If lRetCode <> 0 Then
                objCompany.GetLastError(lErrCode, sErrMsg)
                writeLog("Cmp error", True)
                writeLog(sErrMsg, True)
            Else
                writeLog("Connected to " & objCompany.CompanyName, True)

            End If
        Catch ex As Exception
            writeLog(ex.Message.ToString, False)
        End Try
    End Sub

    Public Sub createConnectionToCompany()
        Try
            'Dim Server
            Dim lErrCode, sErrMsg
            objCompany = New SAPbobsCOM.Company
            objCompany.DbServerType = BoDataServerTypes.dst_MSSQL2016
            objCompany.Server = server
            objCompany.LicenseServer = server
            objCompany.language = BoSuppLangs.ln_English
            objCompany.UseTrusted = False
            objCompany.DbUserName = dbUser
            objCompany.DbPassword = dbPassword
            objCompany.CompanyDB = dbName
            objCompany.UserName = sapUser
            objCompany.Password = sapPassword

            '// Connecting to a company DB
            Dim lRetCode = objCompany.Connect

            If lRetCode <> 0 Then
                objCompany.GetLastError(lErrCode, sErrMsg)
                writeLog(sErrMsg, False)
            Else
                writeLog("Connected to " & objCompany.CompanyName, True)

            End If
        Catch ex As Exception
            writeLog(ex.Message.ToString, False)
        End Try
    End Sub
    Public Sub createSciNoteUDF()
        Dim lRetCode As Integer
        Dim sErrMsg As String = ""
        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = objCompany.GetBusinessObject(BoObjectTypes.oUserFields)
        Try

            Try
                oUserFieldsMD.TableName = "OCTR"
                oUserFieldsMD.Name = "SCIPID"
                oUserFieldsMD.Description = "SCI Project ID"
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUserFieldsMD.EditSize = 20
                '// Adding the Field to the Table
                lRetCode = oUserFieldsMD.Add
                '// Check for errors
                If lRetCode <> 0 Then
                    objCompany.GetLastError(lRetCode, sErrMsg)
                    If InStr(sErrMsg, "Error Creating Field") = 0 Then
                        'MsgBox(sErrMsg)
                    End If
                End If
            Catch ex As Exception

            End Try
            Try
                oUserFieldsMD.TableName = "OSCL"
                oUserFieldsMD.Name = "SCIPID"
                oUserFieldsMD.Description = "SCI Project ID"
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUserFieldsMD.EditSize = 20
                '// Adding the Field to the Table
                lRetCode = oUserFieldsMD.Add
                '// Check for errors
                If lRetCode <> 0 Then
                    objCompany.GetLastError(lRetCode, sErrMsg)
                    If InStr(sErrMsg, "Error Creating Field") = 0 Then
                        'MsgBox(sErrMsg)
                    End If
                End If
            Catch ex As Exception

            End Try
            Try
                oUserFieldsMD.TableName = "OSCL"
                oUserFieldsMD.Name = "SCIEID"
                oUserFieldsMD.Description = "SCI Experiment ID"
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUserFieldsMD.EditSize = 20
                '// Adding the Field to the Table
                lRetCode = oUserFieldsMD.Add
                '// Check for errors
                If lRetCode <> 0 Then
                    objCompany.GetLastError(lRetCode, sErrMsg)
                    If InStr(sErrMsg, "Error Creating Field") = 0 Then
                        'MsgBox(sErrMsg)
                    End If
                End If
            Catch ex As Exception

            End Try
            Try
                oUserFieldsMD.TableName = "OCLG"
                oUserFieldsMD.Name = "SCIPID"
                oUserFieldsMD.Description = "SCI Project ID"
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUserFieldsMD.EditSize = 20
                '// Adding the Field to the Table
                lRetCode = oUserFieldsMD.Add
                '// Check for errors
                If lRetCode <> 0 Then
                    objCompany.GetLastError(lRetCode, sErrMsg)
                    If InStr(sErrMsg, "Error Creating Field") = 0 Then
                        'MsgBox(sErrMsg)
                    End If
                End If
            Catch ex As Exception

            End Try
            Try
                oUserFieldsMD.TableName = "OCLG"
                oUserFieldsMD.Name = "SCIEID"
                oUserFieldsMD.Description = "SCI Experiment ID"
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUserFieldsMD.EditSize = 20
                '// Adding the Field to the Table
                lRetCode = oUserFieldsMD.Add
                '// Check for errors
                If lRetCode <> 0 Then
                    objCompany.GetLastError(lRetCode, sErrMsg)
                    If InStr(sErrMsg, "Error Creating Field") = 0 Then
                        'MsgBox(sErrMsg)
                    End If
                End If
            Catch ex As Exception

            End Try
            Try
                oUserFieldsMD.TableName = "OCLG"
                oUserFieldsMD.Name = "SCITID"
                oUserFieldsMD.Description = "SCI Task ID"
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUserFieldsMD.EditSize = 20
                '// Adding the Field to the Table
                lRetCode = oUserFieldsMD.Add
                '// Check for errors
                If lRetCode <> 0 Then
                    objCompany.GetLastError(lRetCode, sErrMsg)
                    If InStr(sErrMsg, "Error Creating Field") = 0 Then
                        'MsgBox(sErrMsg)
                    End If
                End If
            Catch ex As Exception

            End Try
            Try
                oUserFieldsMD.TableName = "OCLG"
                oUserFieldsMD.Name = "SCITR"
                oUserFieldsMD.Description = "SCI Result"
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
                oUserFieldsMD.SubType = BoFldSubTypes.st_Address
                '   oUserFieldsMD.EditSize = 20
                '// Adding the Field to the Table
                lRetCode = oUserFieldsMD.Add
                '// Check for errors
                If lRetCode <> 0 Then
                    objCompany.GetLastError(lRetCode, sErrMsg)
                    If InStr(sErrMsg, "Error Creating Field") = 0 Then
                        'MsgBox(sErrMsg)
                    End If
                End If
            Catch ex As Exception

            End Try
        Catch ex As Exception

        End Try
    End Sub
    Public Sub createServiceContract(ByVal CardCode As String, ByVal projectId As String, ByVal projectDescription As String, ByVal startDate As DateTime, ByRef errors As List(Of String), ByRef BubbleEvent As Boolean, ByRef DocEntry As Integer)
        Try
            Dim SalesOrder As SAPbobsCOM.IServiceContracts
            SalesOrder = Nothing
            SalesOrder = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceContracts)

            SalesOrder.ContractType = BoContractTypes.ct_Customer
            SalesOrder.Description = projectDescription
            SalesOrder.StartDate = startDate
            SalesOrder.EndDate = DateAdd(DateInterval.Day, 100, startDate)
            SalesOrder.CustomerCode = CardCode
            SalesOrder.Status = BoSvcContractStatus.scs_Approved
            SalesOrder.UserFields.Fields.Item("U_SCIPID").Value = projectId
            SalesOrder.UserFields.Fields.Item("U_Sync").Value = "Yes"
            RetValue = SalesOrder.Add()
            If RetValue <> 0 Then
                objCompany.GetLastError(ErrCode, ErrMsg)
                writeLog(String.Format("Project ID {0} error msg  '{1}'", projectId, ErrMsg), True)
                errors.Add(ErrCode.ToString + ": " + ErrMsg)
                BubbleEvent = False
            Else
                DocEntry = objCompany.GetNewObjectKey()
            End If
        Catch ex As Exception
            writeLog(String.Format("exception  createSalesContract for project Id '{0}' & errors '{1}'", projectId, "T4" + ex.Message()), True)
            errors.Add(ErrCode.ToString + ": " + ErrMsg)
            BubbleEvent = False
        End Try

    End Sub

    Public Sub createServiceCalls(ByVal pId As String, ByVal expId As String, CardCode As String, name As String, description As String, ContractId As Integer, created_at As DateTime, ByRef errors As List(Of String), ByRef BubbleEvent As Boolean, ByRef DocEntry As Integer)
        Try
            Dim SericeCalls As SAPbobsCOM.IServiceCalls
            SericeCalls = Nothing
            If objCompany Is Nothing Then
                createConnectionToCompany()
            End If
            SericeCalls = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
            SericeCalls.ServiceBPType = ServiceTypeEnum.srvcSales
            SericeCalls.CustomerCode = CardCode
            SericeCalls.ContractID = ContractId
            SericeCalls.Subject = name
            SericeCalls.Description = description
            SericeCalls.CreationDate = created_at
            SericeCalls.UserFields.Fields.Item("U_SCIPID").Value = pId
            SericeCalls.UserFields.Fields.Item("U_SCIEID").Value = expId
            SericeCalls.UserFields.Fields.Item("U_Sync").Value = "Yes"
            RetValue = SericeCalls.Add()
            If RetValue <> 0 Then
                objCompany.GetLastError(ErrCode, ErrMsg)
                writeLog(String.Format("Experiment ID {0} error msg  '{1}'", expId, ErrMsg), True)
                errors.Add(ErrCode.ToString + ": " + ErrMsg)
                BubbleEvent = False
            Else
                DocEntry = objCompany.GetNewObjectKey()
            End If
        Catch ex As Exception
            writeLog(String.Format("exception  createSalesContract for Experiment Id '{0}' & errors '{1}'", expId, "  " + ex.Message()), True)
            errors.Add(ErrCode.ToString + ": " + ErrMsg)
            BubbleEvent = False
        End Try

    End Sub

    'Public Sub updateActivityAttachment(ByVal Code As String, ByVal filename As String, ByRef errors_ As List(Of String))
    '    Dim errors As New List(Of String)
    '    'Dim DocEntry
    '    Try
    '        Dim vContact As SAPbobsCOM.Contacts
    '        vContact = objCompany.GetBusinessObject(BoObjectTypes.oContacts)
    '        vContact.GetByKey(Code)
    '        'vContact.SaveXML("C:\Temp\" + "001" + ".xml")
    '        'vContact.Browser.ReadXml("C:\Temp\" + "002" + ".xml", 0)
    '        'vContact.CardCode = CardCode
    '        'vContact.Notes = Description
    '        ''vContact.Parentobjecttype. = "191"
    '        '''vNewActivity.DocType = "191"
    '        ''vContact.DocEntry = 308
    '        'vContact.Details = name
    '        ' vContact .Attachments .
    '        'vContact.UserFields.Fields.Item("U_taskId").Value = taskId
    '        'vNewActivity.UserFields.Fields.Item("U_Sync").Value = "Yes"
    '        vContact.AttachmentEntry
    '        vContact.Attachments.Add()
    '        RetValue = vContact.Update()

    '        If RetValue <> 0 Then
    '            objCompany.GetLastError(ErrCode, ErrMsg)
    '            'writeLog(String.Format("Experiment ID {0} error msg  '{1}'", expId, ErrMsg), True)
    '            errors.Add(ErrCode.ToString + ": " + ErrMsg)
    '            'BubbleEvent = False
    '        Else
    '            DocEntry = objCompany.GetNewObjectKey()
    '        End If
    '    Catch ex As Exception
    '        writeLog(String.Format("exception  createSalesContract for task Id '{0}' & errors '{1}'", "", taskId + ex.Message()), True)
    '        errors.Add(ErrCode.ToString + ": " + ErrMsg)
    '        'BubbleEvent = False
    '    End Try


    'End Sub

    Public Sub AddAttachmentInSAP(ByVal attachPath As String)
        Dim sqlQ1 As String = "select Distinct T1.DocEntry  from [resultattachment] T0 Inner Join tasks T1 on T1.s_id = T0.ta_id where isnull(T0.Sync,'') = '' and isnull(T1.DocEntry,0)<>0 "
        'Dim sqlQ1 As String = "select Distinct T1.DocEntry  from [resultattachment] T0 Inner Join tasks T1 on T1.s_id = T0.ta_id "
        Dim sqlQ As String = "select T0.*,T1.DocEntry  from [resultattachment] T0 Inner Join tasks T1 on T1.s_id = T0.ta_id where T1.DocEntry = {0}"
        Dim dtAtt As New DataTable
        dtAtt = getDataTable(sqlQ1)

        For Each row As DataRow In dtAtt.Rows
            Try
                Dim aContractId As String = row("DocEntry")

                Dim ttBool As Boolean = False
                Dim dtAttachments As New DataTable
                Dim sqlQ11 = String.Format(sqlQ, aContractId)
                dtAttachments = getDataTable(sqlQ11)
                Dim oAtt As SAPbobsCOM.Attachments2
                oAtt = objCompany.GetBusinessObject(BoObjectTypes.oAttachments2)
                Dim Line = 0
                Dim errors As New List(Of String)
                For Each rw As DataRow In dtAttachments.Rows
                    Try
                        Dim file_name As String = rw("file_id") + "_" + rw("file_name")
                        writeLog(String.Format("1.file name {0}", file_name), True)
                        Dim fname As String = file_name.Split(".")(0)
                        Dim fext As String = file_name.Split(".")(1)
                        Dim id As String = rw("id")

                        Dim sqlCheckAttachement = String.Format("select * from atc1 where AbsEntry in (Select AtcEntry from OCLG where ClgCode <> {0} and [FileName] = '{1}' )", aContractId, fname)
                        Dim dataAttach As DataTable = getDataTableSAP(sqlCheckAttachement)
                        If dataAttach.Rows.Count > 0 Then
                            writeLog(String.Format("Attachament already exists in other contract ,Kindly revise file name {0}", fname), True)
                            'writeLogAttach(String.Format("Attachament already exists in other contract ,Kindly revise file name {0} SAP Activity {1}", fname, aContractId), True)
                            'ttBool = True
                            'Exit For
                        End If
                        'Dim FileName = String.Format("C:\SciNoteFiles\{0}", file_name)
                        Dim FileName = String.Format(attachPath, file_name)
                        writeLog(String.Format("2.file URL {0}", file_name), True)
                        oAtt.Lines.Add()
                        'Dim Line
                        'Dim GetLineNoSQL As String = String.Format("Select isnull(Max(T0.Line),0) Line from ATC1 T0 inner Join OCLG T1 on T0.AbsEntry = T1.AtcEntry where ClgCode = {0}", aContractId)
                        'Dim dtLine As DataTable = getDataTableSAP(GetLineNoSQL)
                        'Line = dtLine.Rows(0)(0)
                        'Line = Line + 1
                        oAtt.Lines.SetCurrentLine(Line)
                        Line = Line + 1
                        oAtt.Lines.FileName = fname '"Screenshot 2022-11-21 at 11.42.34 PM"
                        oAtt.Lines.FileExtension = fext '"png"
                        oAtt.Lines.SourcePath = System.IO.Path.GetDirectoryName(FileName)
                        oAtt.Lines.Override = BoYesNoEnum.tYES
                        oAtt.Lines.Add()
                    Catch ex As Exception

                    End Try
                Next
                If ttBool = False Then
                    Dim ret As Integer = oAtt.Add
                    If ret <> 0 Then
                        objCompany.GetLastError(ErrCode, ErrMsg)
                        'writeLog(String.Format("Experiment ID {0} error msg  '{1}'", expId, ErrMsg), True)
                        errors.Add(ErrCode.ToString + ": " + ErrMsg)
                        'BubbleEvent = False
                    Else
                        Dim aDocEntry = objCompany.GetNewObjectKey()
                        Dim vContact As SAPbobsCOM.Contacts
                        vContact = objCompany.GetBusinessObject(BoObjectTypes.oContacts)
                        vContact.GetByKey(aContractId)
                        vContact.AttachmentEntry = aDocEntry

                        vContact.Attachments.Add()
                        ret = vContact.Update()
                        If ret <> 0 Then
                            objCompany.GetLastError(ErrCode, ErrMsg)
                            Dim ii = 0
                        Else
                            Dim updateresultattachment = String.Format("Update T0 set Sync = 'Yes' from [resultattachment] T0 Inner Join tasks T1 on T1.s_id = T0.ta_id where T1.DocEntry = '{0}'", aContractId)
                            ExcuteNonQuerySciNote(updateresultattachment)
                        End If
                    End If
                End If

            Catch ex As Exception

            End Try
        Next

    End Sub

    Public Sub createActivity(ByVal eId As String, ByVal pId As String, ByVal CardCode As String, ByVal name As String, ByVal taskId As String, ByRef errors_ As List(Of String), ByRef DocEntry As String, ByVal Description As String, ByVal created_at As Date)
        Dim errors As New List(Of String)
        'Dim DocEntry
        Try
            writeLog("Experiment Task:" + eId + " " + DateTime.Today.ToString, True)
            writeLog("Pid Task:" + pId + " " + DateTime.Today.ToString, True)
            writeLog("Company Name S :" + objCompany.CompanyName + " " + DateTime.Today.ToString, True)

            Dim vContact As SAPbobsCOM.Contacts
            vContact = objCompany.GetBusinessObject(BoObjectTypes.oContacts)
            vContact = objCompany.GetBusinessObject(BoObjectTypes.oContacts)
            'vContact.GetByKey(1366)
            'vContact.SaveXML("C:\Temp\" + "001" + ".xml")
            'vContact.Browser.ReadXml("C:\Temp\" + "002" + ".xml", 0)
            vContact.Activity = BoActivities.cn_Task
            vContact.ActivityType = 6
            Dim dt As DataTable = getDataTableSAP(String.Format("Select CstmrCode from  OCTR where U_SCIPID={0}", pId))
            vContact.CardCode = dt(0)(0) 'CardCode
            If Description <> "" Then
                vContact.Notes = Description
            End If
            vContact.StartDate = created_at
            'vContact.Parentobjecttype. = "191"
            ''vNewActivity.DocType = "191"
            'vContact.DocEntry = 308
            vContact.Details = name
            ' vContact .Attachments .
            vContact.UserFields.Fields.Item("U_SCIPID").Value = pId
            vContact.UserFields.Fields.Item("U_SCIEID").Value = eId
            vContact.UserFields.Fields.Item("U_SCITID").Value = taskId
            Dim sqlResult As String = String.Format("SELECT dbo.[udf_StripHTML]( result ) result FROM [resulttext] Where  p_id = {0} and e_id = {1} and ta_id ={2}", pId, eId, taskId)
            Dim dtr As DataTable = getDataTable(sqlResult)
            If dtr.Rows.Count > 0 Then
                Dim rest = (dtr(0)(0)).ToString().Trim

                vContact.UserFields.Fields.Item("U_SCITR").Value = rest
            End If

            'vContact.UserFields.Fields.Item("U_Sync").Value = "Yes"
            RetValue = vContact.Add()

            If RetValue <> 0 Then
                objCompany.GetLastError(ErrCode, ErrMsg)
                'writeLog(String.Format("Experiment ID {0} error msg  '{1}'", expId, ErrMsg), True)
                errors.Add(ErrCode.ToString + ": " + ErrMsg)
                writeLog("Task eror" + ErrMsg, True)
                'BubbleEvent = False
            Else
                DocEntry = objCompany.GetNewObjectKey()
            End If
        Catch ex As Exception
            writeLog("Company Name :" + objCompany.CompanyName + " " + DateTime.Today.ToString, True)
            writeLog(String.Format("exception  createSalesContract for task Id '{0}' & errors '{1}'", "", taskId + ex.Message()), True)
            errors.Add(ErrCode.ToString + ": " + ErrMsg)
            'BubbleEvent = False
        End Try


    End Sub
    Public Sub updateActivity(ByRef errors_ As List(Of String), ByRef DocEntry As String, ByVal Description As String)
        Dim errors As New List(Of String)
        'Dim DocEntry
        Try
            Dim vContact As SAPbobsCOM.Contacts
            vContact = objCompany.GetBusinessObject(BoObjectTypes.oContacts)
            'vContact.GetByKey(1366)
            'vContact.SaveXML("C:\Temp\" + "001" + ".xml")
            'vContact.Browser.ReadXml("C:\Temp\" + "002" + ".xml", 0)
            'vContact.CardCode = CardCode
            Dim GetVal = vContact.GetByKey(DocEntry)
            vContact.Notes = Description
            'vContact.Parentobjecttype. = "191"
            ''vNewActivity.DocType = "191"
            'vContact.DocEntry = 308
            'vContact.Details = name
            ' vContact .Attachments .
            'vContact.UserFields.Fields.Item("U_taskId").Value = taskId
            'vNewActivity.UserFields.Fields.Item("U_Sync").Value = "Yes"
            RetValue = vContact.Update()

            If RetValue <> 0 Then
                objCompany.GetLastError(ErrCode, ErrMsg)
                'writeLog(String.Format("Experiment ID {0} error msg  '{1}'", expId, ErrMsg), True)
                errors.Add(ErrCode.ToString + ": " + ErrMsg)
                'BubbleEvent = False
            Else
                DocEntry = objCompany.GetNewObjectKey()
            End If
        Catch ex As Exception
            writeLog(String.Format("exception  update SAP Activity for task Id '{0}' & errors '{1}'", "", DocEntry + ex.Message()), True)
            errors.Add(ErrCode.ToString + ": " + ErrMsg)
            'BubbleEvent = False
        End Try


    End Sub
    Public Sub updateServiceByActivity(ByVal ActivityId As Integer, ByVal ServiceId As Integer, ByVal rw As Integer, ByVal taskId As String, ByRef errors_ As List(Of String), ByRef DocEntry As String)
        Try
            ''' Update activity to service call
            'If DocEntry <> "" Then
            Dim SericeCalls As SAPbobsCOM.IServiceCalls
            SericeCalls = Nothing
            SericeCalls = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
            SericeCalls.ServiceBPType = ServiceTypeEnum.srvcSales
            SericeCalls.GetByKey(ServiceId)
            SericeCalls.Activities.Add()
            SericeCalls.Activities.SetCurrentLine(rw)
            'Dim act As Integer = Convert.ToInt32(DocEntry)
            SericeCalls.Activities.ActivityCode = ActivityId
            RetValue = SericeCalls.Update

            If RetValue <> 0 Then
                objCompany.GetLastError(ErrCode, ErrMsg)
                'writeLog(String.Format("Experiment ID {0} error msg  '{1}'", expId, ErrMsg), True)
                'errors.Add(ErrCode.ToString + ": " + ErrMsg)
                'BubbleEvent = False
            Else
                DocEntry = objCompany.GetNewObjectKey()
            End If
            'End If


        Catch ex As Exception

        End Try
    End Sub
    Public Sub createSalesDiliveryBased(ByVal objSalesInvoice As SalesInvoiceHeaderData, ByRef errors As List(Of String), ByRef DocEntry As Integer)
        Try
            'objCompany.StartTransaction()
            Dim SaleInvoice As SAPbobsCOM.Documents
            SaleInvoice = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
            SaleInvoice.DocDate = objSalesInvoice.DocDate
            SaleInvoice.TaxDate = objSalesInvoice.TaxDate
            SaleInvoice.CardCode = objSalesInvoice.CardCode
            Dim BaseEntry = objSalesInvoice.BaseEntry

            Dim SOLinesDt = getDataTableSAP(String.Format("Select * from RDR1 Where DocEntry={0}", BaseEntry))
            For Each row As DataRow In SOLinesDt.Rows
                'SaleInvoice.Lines.ItemCode = row("ItemCode")
                'SaleInvoice.Lines.Quantity = row("Quantity")
                SaleInvoice.Lines.BaseEntry = row("DocEntry")
                SaleInvoice.Lines.BaseLine = row("LineNum")
                SaleInvoice.Lines.BaseType = 17
                Dim ItemCode = row("ItemCode")
                Dim Warehouse = row("WhsCode")
                Dim Quantity = row("Quantity")
                writeLog(String.Format("Delivery for ItemCode '{0}' and Warehouse = '{1}' ,ORDR Entry '{2}' , Quantity = '{3}'", ItemCode, Warehouse, row("DocEntry"), Quantity), True)
                Dim BatchDT = getDataTableSAP(String.Format("Select * from OIBT Where ItemCode='{0}' and WhsCode = '{1}' and Quantity >0", ItemCode, Warehouse))
                writeLog(String.Format(String.Format("Select * from OIBT Where ItemCode='{0}' and WhsCode = '{1}' and Quantity >0", ItemCode, Warehouse) + "    ,ORDR Entry '{2}'", ItemCode, Warehouse, row("DocEntry")), True)
                Dim BatchQty = 0
                For Each Batchrow As DataRow In BatchDT.Rows
                    SaleInvoice.Lines.BatchNumbers.BatchNumber = Batchrow("BatchNum")
                    BatchQty = Batchrow("Quantity")
                    If BatchQty >= Quantity Then
                        SaleInvoice.Lines.BatchNumbers.Quantity = Quantity
                        SaleInvoice.Lines.BatchNumbers.Add()

                        writeLog(String.Format("+Batch Qty '{0}' & Quantity '{1}'", BatchQty, Quantity), True)
                        Exit For
                    ElseIf BatchQty < Quantity Then
                        SaleInvoice.Lines.BatchNumbers.Quantity = BatchQty
                        SaleInvoice.Lines.BatchNumbers.Add()
                        Quantity = Quantity - BatchQty
                        writeLog(String.Format("-Batch Qty '{0}' & Quantity '{1}'", BatchQty, Quantity), True)
                    End If


                Next
                SaleInvoice.Lines.Add()
            Next
            Dim sqlExpense As String = String.Format("Select * from rdr3 where Docentry = {0} ", BaseEntry)
            Dim expenseDT As DataTable = getDataTableSAP(sqlExpense)
            'For Each objExpenseData As SalesInvoiceExpenseData In objSalesInvoice.ExpenseData
            'If objExpenseData.ExpenseCode <> "" Then
            For Each row As DataRow In expenseDT.Rows
                SaleInvoice.Expenses.BaseDocEntry = BaseEntry
                SaleInvoice.Expenses.BaseDocType = "17"
                SaleInvoice.Expenses.BaseDocLine = row("LineNum")
                'SaleInvoice.Expenses.ExpenseCode = row("ExpnsCode")
                'SaleInvoice.Expenses.LineTotal = row("BaseSum")
                writeLog(String.Format("-base entry '{0}'--''{1}", BaseEntry, row("LineNum")), True)

                SaleInvoice.Expenses.Add()
            Next
            'SaleInvoice.Expenses.BaseDocEntry = BaseEntry
            'SaleInvoice.Expenses.ExpenseCode = objExpenseData.ExpenseCode
            'SaleInvoice.Expenses.LineTotal = objExpenseData.LineTotal
            ''SalesOrder.Expenses.TaxCode = objExpenseData.TaxCode
            'SaleInvoice.Expenses.Add()
            'End If

            'Next

            RetValue = SaleInvoice.Add()
            If RetValue <> 0 Then
                objCompany.GetLastError(ErrCode, ErrMsg)
                errors.Add(ErrCode.ToString + ": " + ErrMsg)

            Else
                DocEntry = objCompany.GetNewObjectKey()
                'SaleInvoice.DocumentStatus = BoStatus.bost_Close
            End If
        Catch ex As Exception
            errors.Add(ErrCode.ToString + ": " + ErrMsg)

        End Try

    End Sub

    Public Sub createSalesIvoiceBased(ByVal objSalesInvoice As SalesInvoiceHeaderData, ByRef errors As List(Of String), ByRef DocEntry As Integer)
        Try
            'objCompany.StartTransaction()
            Dim SaleInvoice As SAPbobsCOM.Documents
            SaleInvoice = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            SaleInvoice.DocDate = objSalesInvoice.DocDate
            SaleInvoice.TaxDate = objSalesInvoice.TaxDate
            SaleInvoice.CardCode = objSalesInvoice.CardCode
            Dim BaseEntry = objSalesInvoice.BaseEntry

            Dim SOLinesDt = getDataTableSAP(String.Format("Select * from DLN1 Where DocEntry={0}", BaseEntry))
            For Each row As DataRow In SOLinesDt.Rows
                'SaleInvoice.Lines.ItemCode = row("ItemCode")
                'SaleInvoice.Lines.Quantity = row("Quantity")
                SaleInvoice.Lines.BaseEntry = row("DocEntry")
                SaleInvoice.Lines.BaseLine = row("LineNum")
                SaleInvoice.Lines.BaseType = 15
                SaleInvoice.Lines.Add()
            Next
            Dim sqlExpense As String = String.Format("Select * from dln3 where Docentry = {0} ", BaseEntry)
            Dim expenseDT As DataTable = getDataTableSAP(sqlExpense)
            'For Each objExpenseData As SalesInvoiceExpenseData In objSalesInvoice.ExpenseData
            'If objExpenseData.ExpenseCode <> "" Then
            For Each row As DataRow In expenseDT.Rows
                SaleInvoice.Expenses.BaseDocEntry = BaseEntry
                SaleInvoice.Expenses.BaseDocType = "15"
                SaleInvoice.Expenses.BaseDocLine = row("LineNum")
                'SaleInvoice.Expenses.ExpenseCode = row("ExpnsCode")
                'SaleInvoice.Expenses.LineTotal = row("BaseSum")
                writeLog(String.Format("-base entry '{0}'--''{1}", BaseEntry, row("LineNum")), True)

                SaleInvoice.Expenses.Add()
            Next
            'For Each objExpenseData As SalesInvoiceExpenseData In objSalesInvoice.ExpenseData
            '    If objExpenseData.ExpenseCode <> "" Then
            '        SaleInvoice.Expenses.ExpenseCode = objExpenseData.ExpenseCode
            '        SaleInvoice.Expenses.LineTotal = objExpenseData.LineTotal
            '        SaleInvoice.Expenses.TaxCode = sapEXPTAXCODE
            '        SaleInvoice.Expenses.Add()
            '    End If

            'Next

            RetValue = SaleInvoice.Add()
            If RetValue <> 0 Then
                objCompany.GetLastError(ErrCode, ErrMsg)
                errors.Add(ErrCode.ToString + ": " + ErrMsg)

            Else
                DocEntry = objCompany.GetNewObjectKey()
            End If
        Catch ex As Exception
            errors.Add(ErrCode + ": " + ErrMsg)

        End Try

    End Sub

    'Public Sub AddUDO(ByVal objUDOHeader As clsOUDO)


    '    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

    '    oUserObjectMD = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

    '    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
    '    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
    '    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
    '    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
    '    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
    '    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
    '    oUserObjectMD.Code = objUDOHeader.Code
    '    oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
    '    oUserObjectMD.Name = objUDOHeader.Name
    '    oUserObjectMD.ObjectType = objUDOHeader.TYPE  'SAPbobsCOM.BoUDOObjType.boud_Document
    '    oUserObjectMD.TableName = objUDOHeader.TableName
    '    For Each objChild As clsUDO1 In objUDOHeader.List_of_Childs
    '        oUserObjectMD.ChildTables.TableName = objChild.TableName
    '        oUserObjectMD.ChildTables.Add()
    '    Next
    '    For Each findColumns As clsUDO2 In objUDOHeader.List_of_FindColumns
    '        oUserObjectMD.FindColumns.Add()
    '        oUserObjectMD.FindColumns.ColumnAlias = findColumns.ColAlias
    '    Next
    '    ' Handle UDO Form

    '    If objUDOHeader.CanDefForm = "Y" Then
    '        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
    '        For Each headerColumns As clsUDO3 In objUDOHeader.List_of_foundColumns
    '            If headerColumns.SonNum = 0 Then
    '                oUserObjectMD.FormColumns.FormColumnAlias = headerColumns.ColAlias
    '                oUserObjectMD.FormColumns.FormColumnDescription = headerColumns.ColDesc
    '                oUserObjectMD.FormColumns.Add()
    '            End If
    '        Next
    '        If objUDOHeader.List_of_ChildTableColumns.Count > 0 Then
    '            oUserObjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES
    '            oUserObjectMD.RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tYES
    '            For Each colMatrix As clsUDO4 In objUDOHeader.List_of_ChildTableColumns
    '                oUserObjectMD.EnhancedFormColumns.ColumnAlias = colMatrix.ColAlias
    '                oUserObjectMD.EnhancedFormColumns.ColumnDescription = colMatrix.ColDesc
    '                oUserObjectMD.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES
    '                oUserObjectMD.EnhancedFormColumns.ColumnNumber = colMatrix.ColumnNum
    '                oUserObjectMD.EnhancedFormColumns.ChildNumber = colMatrix.SonNum
    '                oUserObjectMD.EnhancedFormColumns.Add()
    '            Next
    '        End If
    '    End If

    '    RetValue = oUserObjectMD.Add()
    '    Dim lRetCode As Integer = 0
    '    Dim sErrMsg As String = ""
    '    '// check for errors in the process
    '    If RetValue <> 0 Then
    '        '  If RetValue = -1 Then
    '        ' chkUDOAfter.SetItemChecked(10, True)
    '        ' Else
    '        objCompany.GetLastError(lRetCode, sErrMsg)
    '        MsgBox(sErrMsg)
    '        ' End If
    '    Else
    '        MsgBox("UDO: " & oUserObjectMD.Name & " was added successfully")
    '        '  chkUDOAfter.SetItemChecked(9, True)
    '    End If

    '    oUserObjectMD = Nothing

    '    GC.Collect() 'Release the handle to the table
    'End Sub

End Class
Public Class SalesInvoiceHeaderData
    Public Property DocDate As Date
    Public Property TaxDate As Date
    Public Property CardCode As String
    Public Property CardName As String
    Public Property Agent As String
    Public Property BaseEntry As String
    Public Property NumAtCard As String
    Public Property Comment As String
    Public Property BillingAddress As String
    Public Property cardNameB As String
    Public Property ShippingAddress As String
    ''' <summary>
    ''' Address2B Billing FirstName + LastName
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property StreetB As String
    Public Property StreetNoB
    Public Property BlockB
    Public Property CityB
    Public Property CountyB
    Public Property CountryB
    Public Property StateB
    Public Property BuildingB

    Public Property StreetS As String
    Public Property StreetNoS
    Public Property BlockS
    Public Property CityS
    Public Property CountyS
    Public Property CountryS
    Public Property StateS
    Public Property BuildingS
    Public Property DetailData As List(Of SalesInvoiceDetailData)
    Public Property ExpenseData As List(Of SalesInvoiceExpenseData)
End Class
Public Class SalesInvoiceDetailData
    Public Property ItemCode As String
    Public Property Quantity As Decimal
    Public Property Price As Decimal
    Public Property TaxCode As String
    Public Property WhsCode As String
    Public Property SO_DocEntry As String
    Public Property SO_LineNum As String
    Public Property Quantity_To_Purchase_Inst As Decimal
    Public Property BatchData As List(Of SalesInvoiceBatchData)
End Class
Public Class SalesInvoiceBatchData
    Public Property BatchNumber As String
    Public Property Quantity As Decimal
End Class
Public Class SalesInvoiceExpenseData
    Public Property ExpenseCode As String
    Public Property LineTotal As Decimal
    Public Property TaxCode As String

    Public Property BaseLine
    Public Property BaseType

End Class
