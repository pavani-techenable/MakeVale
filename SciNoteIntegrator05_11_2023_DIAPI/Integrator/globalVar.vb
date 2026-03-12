
'Imports SAPbobsCOM
Imports System.IO
Imports System.Data.SqlClient
Imports System.Net

Imports System.Text

Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports RestSharp

Imports System.Xml
Module globalVar

    Public Server, Database, DBName, DbUser, DbPassword, Url, refreshToken, SAPDBName, SQL_Project, attpath, attachUrl As String

    Public active_type, live_token_url, demo_token_url, live_tok_url, dem_tok_url As String

    Public Sub writeLog(ByVal msgString As String, ByVal boolSuccess As Boolean)
        Dim today As String = Date.Now.ToString("ddMMMyyyy")
        ' Dim strFile As String = "ASTL_Doff_Log" + today + ".txt"
        Dim SPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim strFile As String = SPath + "\SciNote_Log" + today + ".txt"
        '  Dim fileExists As Boolean = File.Exists(strFile)
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

            Try
                con.ConnectionString = "server=" + server + ";user=sa;password=" + dbPassword + ";database=" + dbName
                '  con.ConnectionString = "server=" + "SERVER\SAPSERVER" + ";user=sa;password=abc@1234;database=" + dbName
                con.Open()
                cmd.Connection = con
                cmd.CommandText = cmdString
                cmd.ExecuteNonQuery()

            Catch ex As Exception
                writeLog(cmdString, False)
                writeLog(ex.Message.ToString, False)
                ' MessageBox.Show("Error while inserting record on table..." & ex.Message, "Insert Records")
            Finally
                con.Close()
            End Try
        Catch ex As Exception

        End Try
    End Sub
    Public Sub ExcuteNonQuery_SciNote(ByVal cmdString As String, Optional writelog_ As Boolean = False)
        Try
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand

            Try
                con.ConnectionString = "server=" + Server + ";user=sa;password=" + DbPassword + ";database=" + Database
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
        Dim Connection = "server=" + Server + ";user=sa;password=" + DbPassword + ";database=" + Database
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

    Public Sub readFromXML()
        Dim xmldoc As New XmlDocument
        Dim fs As New FileStream("config.xml", FileMode.Open, FileAccess.Read)
        xmldoc.Load(fs)
        Dim nodes As XmlNodeList = xmldoc.DocumentElement.SelectNodes("/Data")
        For Each node As XmlNode In nodes
            Server = node.SelectSingleNode("Server").InnerText
            Database = node.SelectSingleNode("Database").InnerText
            DBName = node.SelectSingleNode("DBName").InnerText
            DbUser = node.SelectSingleNode("DbUser").InnerText
            DbPassword = node.SelectSingleNode("DbPassword").InnerText
            Url = node.SelectSingleNode("Url").InnerText
            refreshToken = node.SelectSingleNode("refreshToken").InnerText
            SAPDBName = node.SelectSingleNode("SAPDBName").InnerText
            SQL_Project = node.SelectSingleNode("SQL_Project").InnerText
            attpath = node.SelectSingleNode("attachPath").InnerText
            attachUrl = node.SelectSingleNode("attachUrl").InnerText
            active_type = node.SelectSingleNode("active_type").InnerText
            live_token_url = node.SelectSingleNode("live_token_url").InnerText
            demo_token_url = node.SelectSingleNode("demo_token_url").InnerText
            live_tok_url = node.SelectSingleNode("live_tok_url").InnerText
            dem_tok_url = node.SelectSingleNode("dem_tok_url").InnerText




        Next
    End Sub
    Public Sub attachmenttab(ByVal refToken As String)
        Try
            'readFromXML()
            'initSciNoteDB()
            'Dim url_ = Url
            Dim dtTasks As DataTable = getDataTable("Select  * from Tasks order by id desc")
            For Each row As DataRow In dtTasks.Rows
                Try
                    '  t_id	p_id	e_id	s_id
                    Dim DocEntry As Integer = 0
                    Dim errors_ As New List(Of String)
                    Dim TeamsId As String = row("t_id")
                    Dim ProjectId As String = row("p_id")
                    Dim ExperimentId As String = row("e_id")
                    Dim taskId As String = row("s_id")
                    Dim tasksUrl = Url + String.Format("{0}/projects/{1}/experiments/{2}/tasks/{3}/results", TeamsId, ProjectId, ExperimentId, taskId)
                    Dim urip As Uri = New Uri(tasksUrl)
                    Dim responsesp As String = ""
                    responsesp = Woo_GetAllTeams(urip, refToken)
                    Dim datalistp As clsattchment = JsonConvert.DeserializeObject(Of clsattchment)(responsesp)
                    If datalistp.included.Count > 0 Then
                        For Each attri As includedc In datalistp.included
                            Dim rr As String = attri.attributes.ToString
                            Dim rrs As attrib = JsonConvert.DeserializeObject(Of attrib)(rr)
                            rr = ""
                            Dim file_id As String = rrs.file_id
                            Dim file_name As String = rrs.file_name
                            Dim url As String = rrs.url
                            If url <> Nothing Then
                                Dim SQLinsertatt As String = String.Format("INSERT INTO [dbo].[resultattachment]([t_id] ,[p_id] ,[e_id] ,[ta_id] ,[file_id],[file_name] ,[url]) VALUES ('{0}' ,'{1}' ,'{2}' ,'{3}' ,'{4}','{5}' ,'{6}')", TeamsId, ProjectId, ExperimentId, taskId, file_id, file_name, url)
                                ExcuteNonQuery_SciNote(SQLinsertatt, True)
                            End If

                            Dim aa_ = ""
                        Next
                        Dim test = ""
                    End If

                Catch ex As Exception

                End Try
            Next

        Catch ex As Exception

        End Try
    End Sub
    Public Sub Main()

        readFromXML()
        initSciNoteDB()
        Dim url_ = Url


        '  testRest()
        'Dim url__live As String = "https://makevalegroup.scinote.net/oauth/token"
        'Dim url_refresh_live As String = String.Format("grant_type=refresh_token&client_id=d5843a09-f86a-4526-a484-1de8602a02de&client_secret=2a794653-01fa-4acb-a473-1c3b1d9fbf55&refresh_token={0}&redirect_uri=urn:ietf:wg:oauth:2.0:oob", refreshToken)

        'Dim url__ As String = "https://makevale.scinote.net/oauth/token"
        'Dim url_refresh As String = String.Format("grant_type=refresh_token&client_id=YPH0M6UJPrJVd4OBAZ9gtEApDp4OiL-UIp-sJohxFsU&client_secret=0RNBbF4q6iA3ZhXNypISPJtLEXdAmPuVJFlrJIGxMTA&refresh_token=YGLWMhNG1UwdpueOkOqDm2T9MBu2CYEJpQc1kHBoqjg&redirect_uri=urn:ietf:wg:oauth:2.0:oob", refreshToken)
        Dim url__ As String = ""
        Dim url_refresh As String = ""
        If active_type = "Demo" Then
            url__ = demo_token_url
            url_refresh = String.Format(dem_tok_url, refreshToken)
        ElseIf active_type = "Live" Then
            url__ = live_token_url
            url_refresh = String.Format(live_tok_url, refreshToken)
        End If

        Dim keyRes As String = get_refreshtoken(url__, url_refresh)
        Dim datalistt As tokenKey = JsonConvert.DeserializeObject(Of tokenKey)(keyRes)
        Dim access_Token = datalistt.access_token
        access_Token = "Bearer " + access_Token
        '''End
        Dim uri As Uri = New Uri(Url)
        Dim responses As String = ""
        responses = Woo_GetAllTeams(uri, access_Token)
        Dim datalist As clsTeams = JsonConvert.DeserializeObject(Of clsTeams)(responses)
        For Each dd As data In datalist.data
            writeLog(String.Format("Teams Log"), False)
            Dim data_ = dd.attributes
            Dim id As Integer = dd.id
            Dim type As String = dd.type
            Dim ff_ As New attri
            ff_ = JObject.FromObject(data_).ToObject(Of attri)
            'writeLog(String.Format("Teams Id ={0}", id.ToString), False)
            Dim name = ff_.name, description = ff_.description, space_taken = ff_.space_taken, created_at = ff_.created_at, updated_at = ff_.updated_at
            'writeLog(String.Format("Teams Id ={0}...", id.ToString), False)
            Try
                'Dim created_at_d = DateTime.Parse(created_at)
                'created_at = created_at_d.ToUniversalTime
                '' created_at_d = DateTime.ParseExact(created_at_d, "dd/MM/yyyy h:mm:ss tt", System.Globalization.CultureInfo.InvariantCulture)
                'writeLog(String.Format(created_at.ToString), False)
                'created_at = created_at_d.ToString
                'Dim updated_at_d = DateTime.Parse(updated_at).ToUniversalTime
                'writeLog(String.Format(updated_at_d.ToString), False)
                'updated_at = updated_at_d.ToString()

            Catch ex As Exception

            End Try
            Dim created_at_ = DateTime.Parse(created_at).ToString("yyyy-MM-ddTHH:mm:ss")
            Dim updated_at_ = DateTime.Parse(updated_at).ToString("yyyy-MM-ddTHH:mm:ss")
            Dim insert_cmd = String.Format("INSERT INTO teams (s_id ,name ,description ,space_taken ,created_at ,updated_at ) VALUES ({0} ,'{1}' ,'{2}' ,'{3}' ,'{4}','{5}')", id, name, description, space_taken, created_at_, updated_at_)
            writeLog(String.Format(insert_cmd), False)
            ExcuteNonQuery_SciNote(insert_cmd, True)
            Dim aa_ = ""
        Next

        '--Update Projects By SAP No.
        'Dim sapServicedt As DataTable = getDataTableSAP(" Select CstmrCode,ContractId,Descriptio,U_Project from OCTR where U_Sync = 'No' and isnull(Descriptio,'') <>'' and isnull(U_Project,'')=''")

        'For Each rws As DataRow In sapServicedt.Rows
        '    writeLog(String.Format("Update SAP OCTR Log"), False)
        '    Dim Contract = rws("ContractId")
        '    Dim PID = rws("U_Project")
        '    Dim CardCode = rws("CstmrCode")
        '    Dim Descriptio = rws("Descriptio")
        '    Dim teamsDt_ As DataTable = getDataTable(String.Format("Select * from teams where CardCode = '{0}'", CardCode))
        '    If teamsDt_.Rows.Count > 0 Then
        '        Try


        '            Dim TeamId As String = teamsDt_.Rows(0)("s_id").ToString
        '            Dim projectUrl = Url + String.Format("{0}/projects/", TeamId)
        '            Dim urip As Uri = New Uri(projectUrl)
        '            Dim responsesp As String = ""
        '            responsesp = create_Project(urip, Descriptio, access_Token)
        '            Dim pp = ""
        '            Dim datalistp As clsProjectSNG = JsonConvert.DeserializeObject(Of clsProjectSNG)(responsesp)
        '            ' datalistp.data.
        '            '  For Each pData As datap In datalistp.
        '            writeLog(String.Format("Projects Log"), False)
        '            Dim data_ = datalistp.data.attributes
        '            Dim id As Integer = datalistp.data.id
        '            Dim type As String = datalistp.data.type
        '            Dim ff_ As New attrip
        '            ff_ = JObject.FromObject(data_).ToObject(Of attrip)
        '            Dim name = ff_.name, visibility = ff_.visibility, start_date = ff_.start_date, archived = ff_.archived, created_at = ff_.created_at, updated_at = ff_.updated_at
        '            Dim start_date_ = DateTime.Parse(start_date).ToString("yyyy-MM-ddTHH:mm:ss")

        '            Dim created_at_ = DateTime.Parse(created_at).ToString("yyyy-MM-ddTHH:mm:ss")
        '            Dim updated_at_ = DateTime.Parse(updated_at).ToString("yyyy-MM-ddTHH:mm:ss")
        '            Dim insert_cmd = String.Format("INSERT INTO projects (t_id ,s_id ,name ,visibility ,start_date ,archived ,created_at ,updated_at) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}') ", TeamId, id, name, visibility, start_date_, archived, created_at_, updated_at_)
        '            writeLog(String.Format(insert_cmd), False)
        '            ExcuteNonQuery_SciNote(insert_cmd, True)
        '            Dim updtSQL As String = String.Format("Update T0 set Sync = 'Yes' , DocEntry ='{0}' from projects T0 where T0.DocEntry is null and s_id ='{1}'", Contract, id.ToString)
        '            ExcuteNonQuery_SciNote(updtSQL, True)
        '            Dim SQL_UpdateSAPOCTR As String = String.Format("Update OCTR set  U_Sync = 'Yes',U_Project ='{1}' Where ContractID ={0} ", Contract, id.ToString)
        '            ExcuteNonQuery(SQL_UpdateSAPOCTR)
        '            Dim aa_ = ""
        '            '  Next
        '        Catch ex As Exception

        '        End Try
        '    End If

        '    Dim a = "b"

        'Next




        Dim teamsDt As DataTable = getDataTable("Select * from teams where isnull(CardCode,'')<>''")
        '---GetProjects
        For Each rowt As DataRow In teamsDt.Rows
            Dim teamsId = rowt("s_id")
            Dim projectUrl = Url + String.Format("{0}/projects/", teamsId)
            Dim urip As Uri = New Uri(projectUrl)
            Dim responsesp As String = ""
            responsesp = Woo_GetAllTeams(urip, access_Token)
            Dim datalistp As clsProject = JsonConvert.DeserializeObject(Of clsProject)(responsesp)
            For Each pData As datap In datalistp.data
                writeLog(String.Format("Projects Log"), False)
                Dim data_ = pData.attributes
                Dim id As Integer = pData.id
                Dim type As String = pData.type
                Dim ff_ As New attrip
                ff_ = JObject.FromObject(data_).ToObject(Of attrip)
                Dim name = ff_.name, visibility = ff_.visibility, start_date = ff_.start_date, archived = ff_.archived, created_at = ff_.created_at, updated_at = ff_.updated_at
                Dim start_date_ = DateTime.Parse(start_date).ToString("yyyy-MM-ddTHH:mm:ss")

                Dim created_at_ = DateTime.Parse(created_at).ToString("yyyy-MM-ddTHH:mm:ss")
                Dim updated_at_ = DateTime.Parse(updated_at).ToString("yyyy-MM-ddTHH:mm:ss")
                Dim insert_cmd = String.Format("INSERT INTO projects (t_id ,s_id ,name ,visibility ,start_date ,archived ,created_at ,updated_at) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}') ", teamsId, id, name, visibility, start_date_, archived, created_at_, updated_at_)
                writeLog(String.Format(insert_cmd), False)
                ExcuteNonQuery_SciNote(insert_cmd, True)
                Dim aa_ = ""
            Next
        Next


        '---GetExperiments 
        Dim projectsDt As DataTable = getDataTable("Select * from projects order by t_id,s_id")

        For Each rowt As DataRow In projectsDt.Rows
            writeLog(String.Format("Experiments Log"), False)
            Dim teamsId = rowt("t_id")
            Dim projectsId = rowt("s_id")
            Dim experimentsUrl = Url + String.Format("{0}/projects/{1}/experiments", teamsId, projectsId)
            Dim urip As Uri = New Uri(experimentsUrl)
            Dim responsesp As String = ""
            responsesp = Woo_GetAllTeams(urip, access_Token)
            Dim datalistp As clsProject = JsonConvert.DeserializeObject(Of clsProject)(responsesp)
            For Each pData As datap In datalistp.data
                Dim data_ = pData.attributes
                Dim id As Integer = pData.id
                Dim type As String = pData.type
                Dim ff_ As New attrip
                ff_ = JObject.FromObject(data_).ToObject(Of attrip)
                Dim name = ff_.name, description = ff_.description, archived = ff_.archived, created_at = ff_.created_at, updated_at = ff_.updated_at
                Dim created_at_ = DateTime.Parse(created_at).ToString("yyyy-MM-ddTHH:mm:ss")
                Dim updated_at_ = DateTime.Parse(updated_at).ToString("yyyy-MM-ddTHH:mm:ss")
                Dim insert_cmd = String.Format("INSERT INTO experiments (t_id ,p_id,s_id ,name ,description  ,archived ,created_at ,updated_at) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}') ", teamsId, projectsId, id, name, description, archived, created_at_, updated_at_)
                writeLog(String.Format(insert_cmd), False)
                ExcuteNonQuery_SciNote(insert_cmd, True)
                Dim aa_ = ""
            Next
        Next

        '---GetTasks

        Dim experimentsDt As DataTable = getDataTable("Select * from experiments order by t_id,p_id,s_id")

        For Each rowt As DataRow In experimentsDt.Rows
            Try
                Dim teamsId = rowt("t_id")
                Dim projectsId = rowt("p_id")
                Dim experimentsId = rowt("s_id")
                Dim tasksUrl = Url + String.Format("{0}/projects/{1}/experiments/{2}/tasks", teamsId, projectsId, experimentsId)
                Dim urip As Uri = New Uri(tasksUrl)
                Dim responsesp As String = ""
                responsesp = Woo_GetAllTeams(urip, access_Token)
                Dim datalistp As clsProject = JsonConvert.DeserializeObject(Of clsProject)(responsesp)
                For Each pData As datap In datalistp.data
                    Try
                        writeLog(String.Format("Task Log"), False)

                        Dim data_ = pData.attributes
                        Dim id As Integer = pData.id
                        Dim type As String = pData.type
                        Dim ff_ As New attrip
                        ff_ = JObject.FromObject(data_).ToObject(Of attrip)
                        Dim name = ff_.name, description = ff_.description, archived = ff_.archived, created_at = ff_.created_at, updated_at = ff_.updated_at, started_on = ff_.started_on, due_date = ff_.due_date, state = ff_.state, status_id = ff_.status_id, status_name = ff_.status_name, prev_status_id = ff_.prev_status_id, prev_status_name = ff_.prev_status_name, next_status_id = ff_.next_status_id, next_status_name = ff_.next_status_name, x = ff_.x, y = ff_.y
                        Dim crddt = DateTime.Parse(created_at).ToString("yyyy-MM-ddTHH:mm:ss") '("yyyyMMddHHmmss")
                        Dim updt = DateTime.Parse(updated_at).ToString("yyyy-MM-ddTHH:mm:ss") '("yyyy-MM’-‘dd’T’HH’:’mm’:’ss") '("yyyyMMddHHmmss")

                        Dim TskDT As DataTable = getDataTable(String.Format("Select * from Tasks Where p_id={0} and e_id = {1} and s_id ={2}", projectsId, experimentsId, id))
                        If TskDT.Rows.Count = 0 Then
                            Dim insert_cmd = String.Format("INSERT INTO tasks (t_id ,p_id ,e_id ,s_id ,name ,started_on ,due_date ,description ,state ,archived ,status_id ,status_name ,prev_status_id ,prev_status_name ,next_status_id ,next_status_name ,x ,y ,created_at ,updated_at) VALUES ('{0}','{1}' ,'{2}' ,'{3}','{4}' ,'{5}','{6}' ,'{7}' ,'{8}' ,'{9}' ,'{10}' ,'{11}' ,'{12}' ,'{13}' ,'{14}' ,'{15}' ,'{16}' ,'{17}' ,'{18}' ,'{19}') ", teamsId, projectsId, experimentsId, id, name, started_on, due_date, description, state, archived, status_id, status_name, prev_status_id, prev_status_name, next_status_id, next_status_name, x, y, crddt, updt)
                            writeLog(String.Format(insert_cmd.ToString), False)
                            ExcuteNonQuery_SciNote(insert_cmd, True)
                        Else
                            Dim description_ As String = TskDT.Rows(0)("description")
                            If description <> description_ Then
                                Dim insert_cmd = String.Format("Update tasks set description = '{4}' ,updatedes ='Yes' where t_id={0} and p_id = {1} and e_id = {2} and s_id = {3} ", teamsId, projectsId, experimentsId, id, description)
                                writeLog(String.Format(insert_cmd.ToString), False)
                                ExcuteNonQuery_SciNote(insert_cmd, True)
                            End If
                        End If

                        Dim aa_ = ""
                    Catch ex As Exception

                    End Try
                Next
            Catch ex As Exception

            End Try

        Next

        'Get List of attachment 
        'Dim dtTasks As DataTable = getDataTable("Select top 20 * from Tasks order by id desc")
        'For Each row As DataRow In dtTasks.Rows
        '    Try
        '        '  t_id	p_id	e_id	s_id
        '        Dim DocEntry As Integer = 0
        '        Dim errors_ As New List(Of String)
        '        Dim TeamsId As String = row("t_id")
        '        Dim ProjectId As String = row("p_id")
        '        Dim ExperimentId As String = row("e_id")
        '        Dim taskId As String = row("s_id")
        '        Dim tasksUrl = Url + String.Format("{0}/projects/{1}/experiments/{2}/tasks/{3}/results", TeamsId, ProjectId, ExperimentId, taskId)
        '        Dim urip As Uri = New Uri(tasksUrl)
        '        Dim responsesp As String = ""
        '        responsesp = Woo_GetAllTeams(urip)
        '        Dim datalistp As clsProject = JsonConvert.DeserializeObject(Of clsProject)(responsesp)
        '        Dim test = ""
        '    Catch ex As Exception

        '    End Try
        'Next
        attachmenttab(access_Token)
        Dim dtAttach As DataTable = getDataTable("select * from [resultattachment]")
        For Each row As DataRow In dtAttach.Rows
            Try
                Dim filename As String = row("file_name")
                Dim taskId As String = row("file_id")
                'Dim url As String = "https://makevale.scinote.net" + row("url")
                Dim url As String = attachUrl + row("url")
                Dim urip As Uri = New Uri(url)
                Woo_WriteAttachment(urip, filename, access_Token, taskId)
                '        Dim responsesp As String = ""
            Catch ex As Exception

            End Try
        Next
        Application.Exit()
    End Sub
    Public Sub initSciNoteDB()
        Dim DBQuery = "Create Database SciNoteDBAll"
        ExcuteNonQuery(DBQuery)

        Dim WOO_Order_teams = "Create table teams (id int IDENTITY(1,1) NOT NULL,s_id nvarchar(50)  primary key,name nvarchar(150),description nvarchar(150),space_taken nvarchar(250),created_at Datetime,updated_at datetime)"
        ExcuteNonQuery_SciNote(WOO_Order_teams)
        Dim WOO_Order_projects = "  Create table projects ( id int IDENTITY(1,1) NOT NULL,[t_id] nvarchar(50) not null, s_id nvarchar(50),[name] nvarchar (150) not null,visibility nvarchar(50) not null,start_date datetime,archived bit,created_at Datetime,updated_at datetime,primary key(t_id,s_id))"
        ExcuteNonQuery_SciNote(WOO_Order_projects)
        Dim WOO_Order_experiments = "CREATE TABLE experiments(id int IDENTITY(1,1) NOT NULL,[t_id] nvarchar(50) not null, p_id nvarchar(50) not null,s_id nvarchar(50) not null,name nvarchar(150),description nvarchar(max),archived bit,created_at Datetime,updated_at datetime,primary key(t_id,p_id,s_id))"
        ExcuteNonQuery_SciNote(WOO_Order_experiments)

        Dim WOO_Order_tasks = "CREATE TABLE  tasks(id int IDENTITY(1,1) NOT NULL,[t_id] nvarchar(50) not null, p_id nvarchar(50) not null,e_id nvarchar(50) not null,s_id nvarchar(50) not null,name nvarchar(250),started_on datetime,due_date datetime,description nvarchar(150),state nvarchar(150),archived bit,status_id int,status_name nvarchar(150),prev_status_id int,prev_status_name nvarchar(150),next_status_id int ,next_status_name nvarchar(150),x int,y int,created_at Datetime,updated_at datetime,primary key(t_id,p_id,e_id,s_id))"
        ExcuteNonQuery_SciNote(WOO_Order_tasks)
        Dim WOO_Order_activities = "CREATE TABLE  activities(id int IDENTITY(1,1) NOT NULL,[t_id] nvarchar(50) not null, p_id nvarchar(50) not null,e_id nvarchar(50) not null,ta_id nvarchar(50) not null,s_id nvarchar(50) not null,created_at Datetime,updated_at datetime)"

        ExcuteNonQuery_SciNote(WOO_Order_activities)
        Dim Teams_ext As String = "Alter table teams add CardCode nvarchar(250)"
        ExcuteNonQuery_SciNote(Teams_ext)
        Dim sql_ext As String = "Alter table projects add Sync nvarchar(50), DocEntry nvarchar(50)"
        ExcuteNonQuery_SciNote(sql_ext)
        sql_ext = "Alter table experiments add Sync nvarchar(50), DocEntry nvarchar(50)"
        ExcuteNonQuery_SciNote(sql_ext)
        sql_ext = "Alter table tasks add Sync nvarchar(50), DocEntry nvarchar(50)"
        ExcuteNonQuery_SciNote(sql_ext)

        Dim WOO_Order_resultattachment = "CREATE TABLE  resultattachment (id int IDENTITY(1,1) NOT NULL,[t_id] nvarchar(50) not null, p_id nvarchar(50) not null,e_id nvarchar(50) not null,ta_id nvarchar(50) not null,file_id nvarchar(50) not null,file_name nvarchar(250) not null,url nvarchar(max),primary key(t_id,p_id,e_id,ta_id,file_id))"

        ExcuteNonQuery_SciNote(WOO_Order_resultattachment)
        Dim WOO_att = "Alter table resultattachment add Sync nvarchar(50)"
        ExcuteNonQuery_SciNote(WOO_att)

        sql_ext = "Alter table tasks add updatedes nvarchar(50)"
        ExcuteNonQuery_SciNote(sql_ext)

    End Sub

    Public Function getDataTableSAP(ByVal sqlQuery As String)
        Dim oDT As New DataTable
        Dim Connection = "server=" + Server + ";user=sa;password=" + DbPassword + ";database=" + DBName
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

    Private Function Woo_GetAllTeams(uri As Uri, ByVal refToken As String) As String

        Dim responses As String = ""
        Dim request As WebRequest
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3

        request = WebRequest.Create(uri)


        System.Net.ServicePointManager.UseNagleAlgorithm = False
        System.Net.ServicePointManager.Expect100Continue = False

        request.Method = "GET"

        request.Headers.Add("Authorization", refToken)


        Try
            Dim response As WebResponse = request.GetResponse()
            Using responseStream = request.GetResponse.GetResponseStream
                Using reader As New StreamReader(responseStream)
                    responses = reader.ReadToEnd()
                End Using
            End Using

            'Dim obj = JsonConvert.DeserializeObject < clsTeams > (responses)


            'Dim ServerList As List(Of product) = JsonConvert.DeserializeObject(Of List(Of product))(responses)

            'Dim ServerList As List(Of product) = JsonConvert.DeserializeObject(Of List(Of product))(responses)
            'Dim ordersDt As DataTable = getDataTable("Select * from products")
            ''DataGridView1.DataSource = dv
            'For Each rt As product In ServerList
            '    'Dim obj As Class1 = rt.Property1
            '    Dim Id As String = rt.id
            '    Dim name As String = rt.name
            '    Dim sku As String = rt.sku



            '    Dim dv As New DataView(ordersDt)
            '    dv.RowFilter = String.Format("w_id = {0}", Id)

            '    If dv.Count = 0 Then
            '        Dim ins = String.Format("insert into products (w_id],[name],[syncrequired],sku) values ({0},'{1}','{2}','{3}')", Id, name, "0", sku)
            '        ExcuteNonQueryW00(ins, True)
            '    End If

            'Next



        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try


        Return responses
    End Function


    Private Function Woo_WriteAttachment(uri As Uri, ByVal fileName As String, ByVal RefToken As String, ByVal TaskId As String) As String

        Dim responses As String = ""
        Dim request As WebRequest
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3

        request = WebRequest.Create(uri)


        System.Net.ServicePointManager.UseNagleAlgorithm = False
        System.Net.ServicePointManager.Expect100Continue = False

        request.Method = "GET"

        request.Headers.Add("Authorization", RefToken)
        'Dim LocalFilePath As String = String.Format("P:\SciNoteFiles\{0}", fileName)

        Dim LocalFilePath As String = String.Format(attpath, TaskId + "_" + fileName)
        Try
            'Dim response As WebResponse = request.GetResponse()
            'Using responseStream = request.GetResponse.GetResponseStream
            '    Using reader As New StreamReader(responseStream)
            '        responses = reader.ReadToEnd()
            '    End Using
            'End Using

            Using reader As IO.Stream = request.GetResponse.GetResponseStream
                Using writer As IO.Stream = New IO.FileStream(LocalFilePath, IO.FileMode.OpenOrCreate, IO.FileAccess.ReadWrite)
                    Dim b(1024 * 2) As Byte
                    Dim buffer As Integer = b.Length
                    Do While buffer <> 0
                        buffer = reader.Read(b, 0, b.Length)
                        writer.Write(b, 0, buffer)
                        writer.Flush()
                    Loop
                End Using
            End Using




        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try


        Return responses
    End Function



    'Private Function Woo_GetOrdersByID(uri As Uri) As String

    '    Dim responses As String
    '    Dim request As WebRequest
    '    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
    '    ' ServicePointManager.Expect100Continue = True
    '    request = WebRequest.Create(uri)
    '    '  request.Proxy = DBNull.Value

    '    System.Net.ServicePointManager.UseNagleAlgorithm = False
    '    System.Net.ServicePointManager.Expect100Continue = False
    '    'request.ContentLength = jsonDataBytes.Length
    '    'request.ContentType = contentType
    '    request.Method = "GET"
    '    Dim uid = woo_uid
    '    request.Headers.Add("Authorization", uid)


    '    Try
    '        'Dim response As WebResponse = request.GetResponse()
    '        'Dim dataStream As Stream = response.GetResponseStream()

    '        Using responseStream = request.GetResponse.GetResponseStream
    '            Using reader As New StreamReader(responseStream)
    '                responses = reader.ReadToEnd()
    '            End Using
    '        End Using




    '        Dim read = Newtonsoft.Json.Linq.JObject.Parse(responses)
    '        Dim Order_Id = read.Item("id").ToString
    '        Dim Order_Key = read.Item("order_key").ToString
    '        Dim date_created = read.Item("date_created").ToString
    '        Dim customer_id = read.Item("customer_id").ToString
    '        Dim B_customerName = read.Item("billing").Item("first_name").ToString + " " + read.Item("billing").Item("last_name").ToString
    '        B_customerName = B_customerName.ToString.Replace("'", "''")
    '        Dim b_company = read.Item("billing").Item("company").ToString.Replace("'", "''")
    '        Dim b_address_1 = read.Item("billing").Item("address_1").ToString.Replace("'", "''")
    '        Dim b_address_2 = read.Item("billing").Item("address_2").ToString.Replace("'", "''")
    '        Dim b_city = read.Item("billing").Item("city").ToString.Replace("'", "''")
    '        Dim b_state = read.Item("billing").Item("state").ToString.Replace("'", "''")
    '        Dim b_postcode = read.Item("billing").Item("postcode").ToString.Replace("'", "''")
    '        Dim b_country = read.Item("billing").Item("postcode").ToString.Replace("'", "''")
    '        Dim b_email = read.Item("billing").Item("email").ToString.Replace("'", "''")
    '        Dim b_phone = read.Item("billing").Item("phone").ToString.Replace("'", "'s'")
    '        Dim billcmd = String.Format("INSERT INTO [dbo].[billing]([w_id],[w_order_key],[customerName],[company],[address_1],[address_2],[city],[state],[postcode],[country],[email],[phone]) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}') ", Order_Id, Order_Key, B_customerName, b_company, b_address_1, b_address_2, b_city, b_state, b_postcode, b_country, b_email, b_phone)
    '        ExcuteNonQueryW00(billcmd, True)
    '        Dim updateHeaderStatus = String.Format(" Update orders set intStatus = '{0}' Where w_id = '{1}'", "Sync Billing", Order_Id)
    '        ExcuteNonQueryW00(updateHeaderStatus, True)
    '        Dim s_customerName = read.Item("shipping").Item("first_name").ToString + " " + read.Item("billing").Item("last_name").ToString
    '        s_customerName = s_customerName.ToString.Replace("'", "''")
    '        Dim s_company = read.Item("shipping").Item("company").ToString.Replace("'", "''")
    '        Dim s_address_1 = read.Item("shipping").Item("address_1").ToString.Replace("'", "''")
    '        Dim s_address_2 = read.Item("shipping").Item("address_2").ToString.Replace("'", "''")
    '        Dim s_city = read.Item("shipping").Item("city").ToString.Replace("'", "''")
    '        Dim s_state = read.Item("shipping").Item("state").ToString.Replace("'", "''")
    '        Dim s_postcode = read.Item("shipping").Item("postcode").ToString.Replace("'", "''")
    '        Dim s_country = read.Item("shipping").Item("postcode").ToString.Replace("'", "''")
    '        'Dim s_email = read.Item("shipping").Item("email").ToString
    '        Dim s_phone = read.Item("shipping").Item("phone").ToString.Replace("'", "''")

    '        Dim shipcmd = String.Format("INSERT INTO [dbo].[shipping]([w_id],[w_order_key],[customerName],[company],[address_1],[address_2],[city],[state],[postcode],[country],[phone]) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}') ", Order_Id, Order_Key, s_customerName, s_company, s_address_1, s_address_2, s_city, s_state, s_postcode, s_country, s_phone)
    '        ExcuteNonQueryW00(shipcmd, True)
    '        updateHeaderStatus = String.Format(" Update orders set intStatus = '{0}' Where w_id = '{1}'", "Sync shipping", Order_Id)
    '        ExcuteNonQueryW00(updateHeaderStatus, True)
    '        'Dim objSalesHeaderData As New SalesInvoiceHeaderData
    '        'objSalesHeaderData.CardCode = "M0001"
    '        'objSalesHeaderData.DocDate = Now
    '        'objSalesHeaderData.DetailData = New List(Of SalesInvoiceDetailData)

    '        Dim o As JObject = JObject.Parse(responses)
    '        Dim results As List(Of JToken) = o.Children().ToList
    '        For Each item As JProperty In results
    '            item.CreateReader()
    '            Select Case item.Name
    '                Case "CC"
    '                    Dim strCC = item.Value.ToString
    '                Case "line_items"

    '                    Dim product_id, sku, price As String
    '                    Dim quantity As String

    '                    For Each subitem As JObject In item.Values
    '                        'Dim objDetailData As New SalesInvoiceDetailData
    '                        product_id = subitem("product_id")
    '                        quantity = subitem("quantity")
    '                        sku = subitem("sku")
    '                        price = subitem("price")

    '                        Dim detail = String.Format("INSERT INTO [dbo].[orders_details]([w_id],[w_order_key],[product_id],[quantity],[sku],[price]) VALUES({0},'{1}','{2}','{3}','{4}','{5}')", Order_Id, Order_Key, product_id, quantity, sku, price)
    '                        ExcuteNonQueryW00(detail, True)
    '                        updateHeaderStatus = String.Format(" Update orders set intStatus = '{0}' Where w_id = '{1}'", "Sync ItemLines", Order_Id)
    '                        ExcuteNonQueryW00(updateHeaderStatus, True)
    '                        'objDetailData.ItemCode = product_id
    '                        'objDetailData.Quantity = quantity
    '                        'objDetailData.Price = price
    '                        'objSalesHeaderData.DetailData.Add(objDetailData)
    '                    Next
    '                Case "shipping_lines"
    '                    Dim method_title, method_id, instance_id, total As String


    '                    For Each subitem As JObject In item.Values
    '                        method_title = subitem("method_title").ToString.Replace("'", "''")
    '                        method_id = subitem("method_id")
    '                        instance_id = subitem("instance_id")
    '                        total = subitem("total")

    '                        Dim detail = String.Format("INSERT INTO [dbo].[freight]([w_id],[w_order_key],[method_title],[method_id],[instance_id],[total])VALUES({0},'{1}','{2}','{3}','{4}','{5}')", Order_Id, Order_Key, method_title, method_id, instance_id, total)
    '                        ExcuteNonQueryW00(detail, True)
    '                        updateHeaderStatus = String.Format(" Update orders set intStatus = '{0}' Where w_id = '{1}'", "Sync freight", Order_Id)
    '                        ExcuteNonQueryW00(updateHeaderStatus, True)
    '                    Next


    '            End Select
    '        Next

    '        'Dim errors = New List(Of String)
    '        'Dim DocEntry As Integer = 0

    '        'objDiOperations.createSalesOrder(objSalesHeaderData, errors, True, DocEntry)
    '        'If errors.Count > 0 Then
    '        '    For Each Str As String In errors
    '        '        writeLog(String.Format("Woo Commerce Order Id : {0} ", Order_Key) + Str, False)
    '        '    Next
    '        'End If


    '        'Dim tst = ""

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try


    '    Return responses
    'End Function


    Public Sub get_refreshtoken2(url_ As String, content_ As String)
        Dim request As WebRequest = WebRequest.Create(url_) '("http://10.000.10.123:1234")
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
        System.Net.ServicePointManager.UseNagleAlgorithm = False
        System.Net.ServicePointManager.Expect100Continue = False
        request.Method = "POST"
        Dim postData As String = content_ '"GENESYS_TEST."
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
        request.ContentType = "application/x-www-form-urlencoded"
        request.ContentLength = byteArray.Length
        Dim dataStream As Stream = request.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close()
        Dim response As WebResponse = request.GetResponse()
        Console.WriteLine((CType(response, HttpWebResponse)).StatusDescription)

        'Using CSharpImpl.__Assign(dataStream, response.GetResponseStream())
        '    Dim reader As StreamReader = New StreamReader(dataStream)
        '    Dim responseFromServer As String = reader.ReadToEnd()
        '    Console.WriteLine(responseFromServer)
        'End Using

        response.Close()

    End Sub
    'Public Sub testRest()
    '    Dim clnt As New RestClient("https://makevalegroup.scinote.net/oauth/token?grant_type=refresh_token&client_id=d5843a09-f86a-4526-a484-1de8602a02de&client_secret=2a794653-01fa-4acb-a473-1c3b1d9fbf55&refresh_token=YirjHwcE5gKrnFSOfsx6uN2U1PSD_SyqZhYiiCANAkQ&redirect_uri=urn:ietf:wg:oauth:2.0:oob")
    '    ' clnt.ExecutePost 
    '    Dim restReq As New RestRequest()
    '    restReq.Method = Method.Post
    '    Dim res As RestResponse = clnt.Execute(restReq)
    '    Dim aa = res.Content
    '    Dim bb = ""
    '    ' Dim request As RestSharp.RestRequest()
    '    'Dim client = New RestClient("https://makevalegroup.scinote.net/oauth/token?grant_type=refresh_token&client_id=d5843a09-f86a-4526-a484-1de8602a02de&client_secret=2a794653-01fa-4acb-a473-1c3b1d9fbf55&refresh_token=YirjHwcE5gKrnFSOfsx6uN2U1PSD_SyqZhYiiCANAkQ&redirect_uri=urn:ietf:wg:oauth:2.0:oob");
    '    '    client.Timeout = -1;
    '    '    var request = New RestRequest(Method.POST);
    '    '    IRestResponse response = client.Execute(request);
    '    '    Console.WriteLine(response.Content);
    'End Sub
    Private Function get_refreshtoken(url_ As String, content_ As String) As String
        Dim uri As Uri = New Uri(url_)
        Dim responses As String
        Dim request As WebRequest
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
        ' ServicePointManager.Expect100Continue = True
        request = WebRequest.Create(uri)
        '  request.Proxy = DBNull.Value

        System.Net.ServicePointManager.UseNagleAlgorithm = False
        System.Net.ServicePointManager.Expect100Continue = False
        'request.ContentLength = jsonDataBytes.Length
        'request.ContentType = contentType
        request.Method = "POST"
        'Dim uid = "Basic Y2tfOGIzNDcwM2ZjZjVmZjQ4YzgzM2Y5MzQxMmRiZDU3MzAzZTA5NWYwOTpjc181MzdiNzUxYjFjZDM5MzRhNWZhNjJkMzQzNzhjZGExOTc0NjI2ZDYx"
        ' request.Headers.Add("Authorization", woo_uid)
        'request.Headers.Add("Host", "<calculated when request is sent>")
        request.ContentType = "application/x-www-form-urlencoded"
        request.Timeout = 10000
        Dim jsonSring = content_ '"{  ""short_description"": """",""manage_stock"": true,""stock_quantity"": qty}".Replace("qty", stock.ToString)
        Dim jsonDataBytes = Encoding.UTF8.GetBytes(jsonSring)
        request.ContentLength = jsonDataBytes.Length
        Try
            'Dim response As WebResponse = request.GetResponse()
            'Dim dataStream As Stream = response.GetResponseStream()
            Using requestStream = request.GetRequestStream
                requestStream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
                requestStream.Close()
                Using responseStream = request.GetResponse.GetResponseStream
                    Using reader As New StreamReader(responseStream)
                        responses = reader.ReadToEnd()
                    End Using
                End Using
            End Using
            'Dim ServerList As List(Of product) = JsonConvert.DeserializeObject(Of List(Of product))(responses)
            Dim aaa = ""
            'Dim ordersDt As DataTable = getDataTable("Select * from products")
            'For Each rt As product In ServerList
            '    'Dim obj As Class1 = rt.Property1
            '    Dim Id As String = rt.id
            '    Dim name As String = rt.name




            '    Dim dv As New DataView(ordersDt)
            '    dv.RowFilter = String.Format("w_id = {0}", Id)

            '    If dv.Count = 0 Then
            '        Dim ins = String.Format("insert into products ([w_id],[name],[syncrequired]) values ({0},'{1}','{2}')", Id, name, rt.syncrequired.ToString())
            '        ExcuteNonQueryW00(ins, True)
            '    End If

            'Next



        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


        Return responses
    End Function
    Private Function get_refreshtoken_(url_ As String, content_ As String) As String

        Dim Uri As Uri = New Uri(url_)
        Dim responses As String
        Dim request As WebRequest
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
        request = WebRequest.Create(Uri)
        System.Net.ServicePointManager.UseNagleAlgorithm = False
        System.Net.ServicePointManager.Expect100Continue = False
        request.Method = "POST"
        request.ContentType = "application/x-www-form-urlencoded"
        Dim byteArray As Byte() = System.Text.Encoding.UTF8.GetBytes(content_)
        request.ContentLength = byteArray.Length
        '  Dim dataStream As Stream = request.GetRequestStream()
        request.Timeout = 10000



        Try
            'Dim response As WebResponse = request.GetResponse()
            'Dim dataStream As Stream = response.GetResponseStream()
            Using requestStream = request.GetRequestStream
                requestStream.Write(byteArray, 0, byteArray.Length)
                requestStream.Close()
                Using responseStream = request.GetResponse.GetResponseStream
                    Using reader As New StreamReader(responseStream)
                        responses = reader.ReadToEnd()
                    End Using
                End Using
            End Using
            'Dim ServerList As List(Of product) = JsonConvert.DeserializeObject(Of List(Of product))(responses)
            Dim aaa = ""
            'Dim ordersDt As DataTable = getDataTable("Select * from products")
            'For Each rt As product In ServerList
            '    'Dim obj As Class1 = rt.Property1
            '    Dim Id As String = rt.id
            '    Dim name As String = rt.name




            '    Dim dv As New DataView(ordersDt)
            '    dv.RowFilter = String.Format("w_id = {0}", Id)

            '    If dv.Count = 0 Then
            '        Dim ins = String.Format("insert into products ([w_id],[name],[syncrequired]) values ({0},'{1}','{2}')", Id, name, rt.syncrequired.ToString())
            '        ExcuteNonQueryW00(ins, True)
            '    End If

            'Next



        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try


        Return responses
    End Function

    Private Function create_Project(uri As Uri, ByVal Description As String, ByVal RefToken As String) As String

        Dim responses As String = ""
        Dim request As WebRequest
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
        ' ServicePointManager.Expect100Continue = True
        request = WebRequest.Create(uri)
        '  request.Proxy = DBNull.Value

        System.Net.ServicePointManager.UseNagleAlgorithm = False
        System.Net.ServicePointManager.Expect100Continue = False
        'request.ContentLength = jsonDataBytes.Length
        'request.ContentType = contentType
        request.Method = "POST"
        'Dim uid = "Basic Y2tfOGIzNDcwM2ZjZjVmZjQ4YzgzM2Y5MzQxMmRiZDU3MzAzZTA5NWYwOTpjc181MzdiNzUxYjFjZDM5MzRhNWZhNjJkMzQzNzhjZGExOTc0NjI2ZDYx"
        request.Headers.Add("Authorization", RefToken)
        'request.Headers.Add("Content-Type", "application/json")
        request.ContentType = "application/json"
        request.Timeout = 10000

        'Dim jsonSring = "{  ""short_description"": """",""manage_stock"": true,""stock_quantity"": qty}".Replace("qty", stock.ToString)
        Dim jsonStringinput = "{""data"": {""type"": ""projects"",""attributes"": { ""name"": """ + Description + """,""visibility"": ""visible"",""archived"": false } }}"
        Dim jsonDataBytes = Encoding.UTF8.GetBytes(jsonStringinput)
        request.ContentLength = jsonDataBytes.Length
        Try
            'Dim response As WebResponse = request.GetResponse()
            'Dim dataStream As Stream = response.GetResponseStream()
            Using requestStream = request.GetRequestStream
                requestStream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
                requestStream.Close()
                Using responseStream = request.GetResponse.GetResponseStream
                    Using reader As New StreamReader(responseStream)
                        responses = reader.ReadToEnd()
                    End Using
                End Using
            End Using
            'Dim ServerList As List(Of product) = JsonConvert.DeserializeObject(Of List(Of product))(responses)
            Dim aaa = ""
            'Dim ordersDt As DataTable = getDataTable("Select * from products")
            'For Each rt As product In ServerList
            '    'Dim obj As Class1 = rt.Property1
            '    Dim Id As String = rt.id
            '    Dim name As String = rt.name




            '    Dim dv As New DataView(ordersDt)
            '    dv.RowFilter = String.Format("w_id = {0}", Id)

            '    If dv.Count = 0 Then
            '        Dim ins = String.Format("insert into products ([w_id],[name],[syncrequired]) values ({0},'{1}','{2}')", Id, name, rt.syncrequired.ToString())
            '        ExcuteNonQueryW00(ins, True)
            '    End If

            'Next



        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try


        Return responses
    End Function

End Module

