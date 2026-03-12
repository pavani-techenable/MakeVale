Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading

Module globalVar
    Public objCompany As SAPbobsCOM.Company
    Public woo_order_url
    Public woo_uid, sapExpnsCode
    Public server, Database, sapUser, sapPassword, attPath, attachPathLog As String
    Public dbName As String
    Public dbPassword As String
    Public dbUser As String
    Public sapWHSCODE As String
    Public sapEXPTAXCODE As String
    Public SQLQueryForCustMapping As String
    'Public TWSAPDB = "TWSAPDB"

    Public Sub writeLog(ByVal msgString As String, ByVal boolSuccess As Boolean)
        Dim today As String = Date.Now.ToString("ddMMMyyyy")
        ' Dim strFile As String = "ASTL_Doff_Log" + today + ".txt"
        Dim SPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        'Dim strFile As String = SPath + "\MakeValeSAPDIAPI_Log" + today + ".txt"
        ' Dim strFile As String = "P:\SAPWOOINT\Sci\Log\SciNoteSAPDIAPI_Log" + today + ".txt"
        Dim strFile As String = SPath + "\SciNoteSAPDIAPI_Log" + today + ".txt"
        Dim fileExists As Boolean = File.Exists(strFile)
        If boolSuccess Then
            msgString = "Success: " & msgString
        Else
            msgString = "Error: " & msgString
        End If
        Using sw As New StreamWriter(File.Open(strFile, FileMode.Append))
            sw.WriteLine(
                IIf(fileExists,
                     DateTime.Now & "-" & msgString,
                  DateTime.Now & "-" & msgString))
        End Using
    End Sub

    Public Sub InsertMailErrorLog(errorMsg As String,
                              errorDescription As String, TableName As String, id As String)

        Dim today As String = Date.Today.ToString("yyyy-MM-dd")
        Dim nowDateTime As DateTime = DateTime.Now
        Dim nowTime As String = DateTime.Now.ToString("HH:mm:ss")



        Dim dt As DataTable = getDataTable(
    "SELECT ISNULL(MAX(DocumentId),0) + 1 AS NextLineId FROM Mailerrorlog")

        Dim NextErrorId As Integer = Convert.ToInt32(dt.Rows(0)("NextLineId"))



        Dim Sdatabasname As String = Database

        If id = String.Empty Then
            id = "0"
        End If
        Dim sql As String = "INSERT INTO Mailerrorlog " &
    "(DocumentId, CreateEDate, ErrorMsg, ErrorDescription, EmailSent,CompanyName,ObjectName,objectid) " &
    "VALUES (" &
    NextErrorId & ",
         '" &
    today & "', '" &
    errorMsg.Replace("'", "''") & "', '" &
    errorDescription.Replace("'", "''") & "', 'N','" + Database + "','" + TableName + "','" + id + "')"

        ExcuteNonQuery_SciNote(sql, True)

    End Sub

    Public Sub ExcuteNonQuery_SciNote(ByVal cmdString As String, Optional writelog_ As Boolean = False)
        Try
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand

            Try
                con.ConnectionString = "server=" + server + ";user=sa;password=" + dbPassword + ";database=" + Database
                '  con.ConnectionString = "server=" + "SERVER\SAPSERVER" + ";user=sa;password=abc@1234;database=" + dbName
                con.Open()
                cmd.Connection = con
                cmd.CommandText = cmdString
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                If writelog_ = True Then
                    writeLog(ex.Message.ToString, False)
                End If

                ' MessageBox.Show("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        Catch ex As Exception

        End Try
    End Sub


    Public Sub writeLogAttach(ByVal msgString As String, ByVal boolSuccess As Boolean)
        Dim today As String = Date.Now.ToString("ddMMMyyyy")

        'Dim SPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

        Dim strFile As String = String.Format(attachPathLog, today + ".txt")

        Dim fileExists As Boolean = File.Exists(strFile)
        If boolSuccess Then
            msgString = "Success: " & msgString
        Else
            msgString = "Error: " & msgString
        End If
        Using sw As New StreamWriter(File.Open(strFile, FileMode.Append))
            sw.WriteLine(
                IIf(fileExists,
                     DateTime.Now & "-" & msgString,
                  DateTime.Now & "-" & msgString))
        End Using
    End Sub
    Public Sub ExcuteNonQuery(ByVal cmdString As String)
        Try
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            'Dim server As String = server
            'Dim dbName As String = dbName
            'Dim dbPassword As String = dbPassword
            'Dim dbUser As String = objMain.objCompany.DbUserName
            Try
                con.ConnectionString = "server=" + server + ";user=sa;password=" + dbPassword + ";database=" + dbName
                '  con.ConnectionString = "server=" + "SERVER\SAPSERVER" + ";user=sa;password=abc@1234;database=" + dbName
                con.Open()
                cmd.Connection = con
                cmd.CommandText = cmdString
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                writeLog(ex.Message.ToString, False)
                ' MessageBox.Show("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        Catch ex As Exception

        End Try
    End Sub
    Public Sub ExcuteNonQuerySciNote(ByVal cmdString As String, Optional writelog_ As Boolean = False)
        Try
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            'Dim server As String = server
            'Dim dbName As String = dbName
            'Dim dbPassword As String = dbPassword
            'Dim dbUser As String = objMain.objCompany.DbUserName
            Try
                con.ConnectionString = "server=" + server + ";user=sa;password=" + dbPassword + ";database=" + Database
                '  con.ConnectionString = "server=" + "SERVER\SAPSERVER" + ";user=sa;password=abc@1234;database=" + dbName
                con.Open()
                cmd.Connection = con
                cmd.CommandText = cmdString
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                If writelog_ = True Then
                    writeLog(ex.Message.ToString, False)
                End If

                ' MessageBox.Show("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        Catch ex As Exception

        End Try
    End Sub
    Public Function getDataTable(ByVal sqlQuery As String)
        Dim oDT As New DataTable
        Dim Connection = "server=" + server + ";user=sa;password=" + dbPassword + ";database=" + Database
        Dim myConnection As New SqlConnection(Connection)
        Try
            myConnection.Open()
            Dim objda As New SqlDataAdapter(sqlQuery, myConnection)
            Dim dt As New DataTable()

            objda.Fill(dt)
            myConnection.Close()
            oDT = dt
        Catch ex As Exception
            myConnection.Close()
            '   psReturn = ex.Message
        End Try
        Return oDT
    End Function

    Public Function getDataTableSAP(ByVal sqlQuery As String)
        Dim oDT As New DataTable
        Dim Connection = "server=" + server + ";user=sa;password=" + dbPassword + ";database=" + dbName
        Dim myConnection As New SqlConnection(Connection)
        Try
            myConnection.Open()
            Dim objda As New SqlDataAdapter(sqlQuery, myConnection)
            Dim dt As New DataTable()

            objda.Fill(dt)
            myConnection.Close()
            oDT = dt
        Catch ex As Exception
            myConnection.Close()
            '   psReturn = ex.Message
        End Try
        Return oDT
    End Function

    Public Sub Main()
        Try
            writeLog(String.Format("main started"), True)

            'writeLog("Main started", False)
            Dim objDiOperations = New clsDIoperations
            objDiOperations.readFromXML()
            objDiOperations.createConnectionToCompany()
            objDiOperations.createSciNoteUDF()
            'InsertMailErrorLog("Connection", "teswting", "Server", "")

            writeLog(String.Format("Connected to company"), True)
            '' Dim GB_CardCode As String
            ' Application.Exit()
            'objDiOperations.createActivity()
            'Application.Exit()
            'Dim dt_projectData As New DataTable
            'dt_projectData = getDataTable("Select * from ServiceContractData")
            'For Each row As DataRow In dt_projectData.Rows
            '    Try
            '        Dim DocEntry As Integer = 0
            '        Dim errors_ As New List(Of String)
            '        Dim CardCode As String = row("CardCode")
            '        Dim projectId As String = row("projectId")
            '        Dim projectDescription As String = row("projectDescription")
            '        Dim startDate As DateTime = row("start_date")
            '        objDiOperations.createServiceContract(CardCode, projectId, projectDescription, startDate, errors_, False, DocEntry)
            '        If errors_.Count = 0 And DocEntry <> 0 Then
            '            Dim updateProject = String.Format("Update projects set Sync = 'Yes',DocEntry = '{0}' where s_id = '{1}'", DocEntry.ToString, projectId)
            '            ExcuteNonQuerySciNote(updateProject)
            '        End If
            '    Catch ex As Exception

            '    End Try
            'Next

            Dim dt_expData As New DataTable
            dt_expData = getDataTable("Select * from experimentsData")
            Dim expId As String = String.Empty
            For Each row As DataRow In dt_expData.Rows
                Try
                    Dim DocEntry As Integer = 0
                    Dim errors_ As New List(Of String)
                    Dim CardCode As String = row("CardCode")
                    expId = row("s_id")
                    Dim pId As String = row("p_id")
                    Dim name As String = row("name")
                    Dim description As String = row("description")
                    Dim ContractId As String = row("ContractId")
                    Dim created_at As DateTime = row("created_at")
                    objDiOperations.createServiceCalls(pId, expId, CardCode, name, description, ContractId, created_at, errors_, False, DocEntry)
                    If errors_.Count = 0 And DocEntry <> 0 Then
                        Dim updateProject = String.Format("Update experiments set Sync = 'Yes',DocEntry = '{0}' where s_id = '{1}'", DocEntry.ToString, expId)
                        ExcuteNonQuerySciNote(updateProject)
                    End If
                Catch ex As Exception
                    InsertMailErrorLog("Posting", ex.Message(), "experiments", expId)
                End Try
            Next
            writeLog(String.Format("Activity Started"), True)
            'Create Activity 
            Dim strtaskdata As String = "Select t0.*,t1.DocEntry,t2.CardCode   from tasks t0 inner join experiments t1 on t1.t_id = t0.t_id and t0.p_id =t1.p_id and t0.e_id =t1.s_id inner join teams t2 on t2.s_id =t0.t_id   where 
t1.DocEntry is not null  and t0.DocEntry is Null"
            Dim dt_taskData As New DataTable
            dt_taskData = getDataTable(strtaskdata)
            Dim taskId As String = ""

            For Each row As DataRow In dt_taskData.Rows
                Try
                    writeLog(String.Format("Started loop in activity creation "), True)


                    Dim DocEntry As Integer = 0
                    Dim errors_ As New List(Of String)
                    Dim name As String = ""
                    If Not IsDBNull(row("name")) Then
                        name = row("name")
                    Else
                        name = ""
                    End If

                    If Not IsDBNull(row("s_id")) Then
                        taskId = row("s_id")


                    End If
                    Dim pId As String = ""
                    If Not IsDBNull(row("p_id")) Then
                        pId = row("p_id")


                    End If
                    Dim eId As String = ""
                    If Not IsDBNull(row("e_id")) Then
                        eId = row("e_id")


                    End If
                    Dim Decription As String = ""
                    If Not IsDBNull(row("description")) Then
                        Decription = row("description")


                    End If
                    Dim cardCode As String = ""
                    If Not IsDBNull(row("CardCode")) Then
                        cardCode = row("CardCode")


                    End If
                    Dim created_at As Date
                    writeLog(String.Format("loop in activity creation "), True)
                    writeLog("M Task:" + taskId + " " + DateTime.Today.ToString, True)
                    writeLog("M Experiment Task:" + eId + " " + DateTime.Today.ToString, True)
                    writeLog("M Pid Task:" + pId + " " + DateTime.Today.ToString, True)
                    writeLog("M Description: " + Decription + " " + DateTime.Today.ToString, True)
                    Try
                        created_at = row("created_at")
                    Catch ex As Exception
                        InsertMailErrorLog("tasks", ex.Message(), "Logic", "")
                    End Try
                    objDiOperations.createActivity(eId, pId, cardCode, name, taskId, errors_, DocEntry, Decription, created_at)
                    If errors_.Count = 0 And DocEntry <> 0 Then
                        Dim updateProject = String.Format("Update tasks set DocEntry = '{0}' where s_id = '{1}'", DocEntry.ToString, taskId)
                        ExcuteNonQuerySciNote(updateProject)
                    End If
                Catch ex As Exception
                    InsertMailErrorLog("Posting", ex.Message(), "tasks", taskId)

                    writeLog(String.Format("Exception loop in activity creation " + ex.Message), True)
                End Try
            Next
            Try
                'Update Activity 
                Dim taskdataUpdate As String = "Select T0.id,T0.t_id,T0.p_id,T0.e_id,T0.s_id,T0.name,T0.started_on,T0.due_date,dbo.[udf_StripHTML](T0.description)description,T0.state,T0.archived,T0.status_id,T0.status_name,T0.prev_status_id,T0.prev_status_name,T0.next_status_id,T0.next_status_name,T0.x,T0.y,T0.created_at,T0.updated_at,T0.Sync,T0.DocEntry,T0.updatedes  from tasks t0 inner join experiments t1 on t1.t_id = t0.t_id and t0.p_id =t1.p_id and t0.e_id =t1.s_id inner join teams t2 on t2.s_id =t0.t_id   where 
t1.DocEntry is not null  and t0.DocEntry is not Null and T0.updatedes ='Yes'"
                Dim dt_taskDataup As New DataTable
                dt_taskDataup = getDataTable(taskdataUpdate)
                Dim taskIdexp As String = String.Empty
                For Each row As DataRow In dt_taskDataup.Rows
                    Try
                        Dim DocEntry As Integer = row("DocEntry")
                        Dim errors_ As New List(Of String)
                        Dim name As String = row("name")
                        taskIdexp = row("s_id")
                        Dim Decription As String = row("description")

                        objDiOperations.updateActivity(errors_, DocEntry, Decription)
                        If errors_.Count = 0 And DocEntry <> 0 Then
                            Dim updateProject = String.Format("Update tasks set updatedes = 'No' where DocEntry = '{0}'", DocEntry.ToString, taskIdexp)
                            ExcuteNonQuerySciNote(updateProject)
                        End If
                    Catch ex As Exception
                        InsertMailErrorLog("Posting", ex.Message(), "tasks", taskIdexp)
                    End Try
                Next
            Catch ex As Exception
                InsertMailErrorLog("experiments", ex.Message(), "Logic", "")
            End Try
            'Link Activity to Service Call
            Dim strlinkSer As String = "Select t0.s_id ,t0.DocEntry ActivityId,t1 .DocEntry ServiceId  from tasks t0 inner join experiments t1 on t1.t_id = t0.t_id and t0.p_id =t1.p_id and t0.e_id =t1.s_id inner join teams t2 on t2.s_id =t0.t_id   where t1.DocEntry is not null  and t0.DocEntry  is not Null and t0.Sync is null"
            Dim dt_linkSer As New DataTable
            dt_linkSer = getDataTable(strlinkSer)
            Dim taskIdService As String = String.Empty
            For Each row As DataRow In dt_linkSer.Rows
                Try
                    Dim DocEntry As Integer = 0
                    Dim errors_ As New List(Of String)
                    Dim ActivityId As String = row("ActivityId")
                    Dim ServiceId As String = row("ServiceId")
                    taskIdService = row("s_id")
                    'Dim cardCode As String = row("CardCode")
                    Dim rwCnt As String = String.Format("Select Count(*) from SCL5 where SrvcCallId = {0}", ServiceId)
                    Dim dtRw As DataTable = getDataTableSAP(rwCnt)
                    Dim rw As Integer = dtRw.Rows(0)(0)
                    objDiOperations.updateServiceByActivity(ActivityId, ServiceId, rw, taskIdService, errors_, DocEntry)
                    If errors_.Count = 0 And DocEntry <> 0 Then
                        Dim updateProject = String.Format("Update tasks set Sync = 'Yes' where s_id = '{0}'", taskIdService)
                        ExcuteNonQuerySciNote(updateProject)
                    End If
                Catch ex As Exception
                    InsertMailErrorLog("Posting", ex.Message(), "tasks", taskIdService)
                End Try
            Next

            writeLog(String.Format("Attachment :{0}", attPath), True)
            objDiOperations.AddAttachmentInSAP(attPath)

            Application.Exit()
                Catch ex As Exception
            InsertMailErrorLog("Main", ex.Message(), "Logic", "")
        End Try
    End Sub

End Module
