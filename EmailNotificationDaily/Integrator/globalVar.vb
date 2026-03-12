
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
Imports Microsoft.Win32
Imports System.Net.Mail
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports System.Diagnostics.Eventing.Reader
Imports System.Threading
Imports System.Security.AccessControl
Imports System.Data.Odbc
Module globalVar

    Public Server, Database, DBName, DbUser, DbPassword, Url, refreshToken, SAPDBName, SQL_Project, attpath, attachUrl, sPath As String

    Public active_type, live_token_url, demo_token_url, live_tok_url, dem_tok_url, log_path As String

    Public Sub writeLog(ByVal msgString As String, ByVal boolSuccess As Boolean)
        Dim today As String = Date.Now.ToString("ddMMMyyyy")
        ' Dim strFile As String = "ASTL_Doff_Log" + today + ".txt"
        'Dim SPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim SPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        SPath = log_path
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
                con.ConnectionString = "server=" + Server + ";user=sa;password=" + DbPassword + ";database=" + DBName
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
            SAPDBName = node.SelectSingleNode("SAPDBName").InnerText
            'SQL_Project = node.SelectSingleNode("SQL_Project").InnerText
            log_path = node.SelectSingleNode("log_path").InnerText
        Next
    End Sub

    Public Sub Main()


        readFromXML()
        initSciNoteDB()
        SendDailyErrorEmail1()
        Application.Exit()




    End Sub
    Public Sub SendDailyErrorEmail1()

        Try


            'Dim today As String = Date.Today.ToString("yyyy-MM-dd")
            'Dim dt As DataTable = getDataTable(
            '    "SELECT distinct  ErrorMsg,EmailSent,CreateEDate, ErrorDescription " &
            '    "FROM Mailerrorlog " &
            '    "WHERE CreateEDate = '" & today & "' " &
            '    "AND EmailSent = 'N' "
            '  )


            'Dim recordCount1 As Integer = dt.Rows.Count

            'Dim count As String = dt.Rows.Count

            'ExcuteNonQuery_SciNote(
            '            "UPDATE Mailerrorlog SET EmailSent = 'P' " &
            '            "WHERE CreateEDate = '" & today & "' AND EmailSent = 'N'",
            '            True
            '            )

            Dim today1 As String = Date.Today.ToString("yyyy-MM-dd")
            Dim dt1 As DataTable = getDataTable(
            "SELECT distinct EmailSent,CreateEDate,ErrorMsg,ErrorDescription " &
            "FROM Mailerrorlog " &
            "WHERE CreateEDate = '" & today1 & "' " &
            "AND EmailSent = 'N' "
        )
            Dim count1 As String = dt1.Rows.Count
            Dim Mailenable As String = Convert.ToString(dt1.Rows(0)("EmailSent"))
            WriteLog("SendErrorMailDaily Count" + count1.ToString())
            If Mailenable = "N" Then

                Dim sb As New StringBuilder()
                sb.Append("<h3>Daily Integration Error Report</h3>")
                sb.Append("<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;width:100%;font-family:Arial;font-size:12px;'>")


                sb.Append("<tr style='background-color:#f2f2f2;font-weight:bold;'>")
                sb.Append("<th>Date</th>")
                sb.Append("<th>Error</th>")
                'sb.Append("<th>Error ID</th>")
                sb.Append("<th>Error Description</th>")
                sb.Append("</tr>")


                For Each r As DataRow In dt1.Rows
                    sb.Append("<tr>")
                    sb.Append("<td>" & Convert.ToDateTime(r("CreateEDate")).ToString("yyyy-MM-dd") & "</td>")
                    sb.Append("<td>" & r("ErrorMsg").ToString() & "</td>")
                    'sb.Append("<td>" & r("ErrorId").ToString() & "</td>")
                    sb.Append("<td>" & r("ErrorDescription").ToString().Replace(Environment.NewLine, "<br/>") & "</td>")
                    sb.Append("</tr>")
                Next

                sb.Append("</table>")

                'Dim sent As Boolean = 

                'SendEmail("192.168.16.15", 25, False, "", "", "sapint@makevale.com", "Grant@makevale.com", "Andy@makevale.com", "it@makevale.com", "Daily Integration Error Report", sb.ToString())


                ExcuteNonQuery_SciNote(
            "UPDATE Mailerrorlog SET EmailSent = 'Y' " &
            "WHERE CreateEDate = '" & today1 & "' AND EmailSent = 'N'",
            True
            )
            End If


        Catch ex As Exception
            WriteLog("SendDailyErrorEmail" + ex.Message.ToString())
        End Try

    End Sub

    Public Sub WriteLog(ByVal Str As String)

        Dim appPath As String = System.Windows.Forms.Application.StartupPath & "\Log\"

        ' Create Log directory if it does not exist
        If Not Directory.Exists(appPath) Then
            Directory.CreateDirectory(appPath)
        End If

        Dim chatlog As String = appPath & "SyncWriteLog_" & DateTime.Now.ToString("yyyyMMdd") & ".txt"

        Dim sdate As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Using objWriter As New StreamWriter(chatlog, True)
            objWriter.WriteLine(sdate & " : " & Str)
        End Using

    End Sub
    'Public Sub SendDailyErrorEmail()

    '    Try
    '        Dim nowTime As DateTime = DateTime.Now
    '        Dim today As String = Date.Today.ToString("yyyy-MM-dd")
    '        Dim dt As DataTable = getDataTable(
    '            "SELECT distinct  ErrorMsg,EmailSent,CreateEDate, ErrorDescription " &
    '            "FROM Mailerrorlog " &
    '            "WHERE CreateEDate = '" & today & "' " &
    '            "AND EmailSent = 'N' "
    '          )

    '        'ExcuteNonQuery_SciNote(
    '        '            "UPDATE Mailerrorlog SET EmailSent = 'P' " &
    '        '            "WHERE CreateEDate = '" & today & "' AND EmailSent = 'N'",
    '        '            True
    '        '            )
    '        Dim count As String = dt.Rows.Count

    '        Dim today1 As String = Date.Today.ToString("yyyy-MM-dd")
    '                Dim dt1 As DataTable = getDataTable(
    '        "SELECT distinct EmailSent " &
    '        "FROM Mailerrorlog " &
    '        "WHERE CreateEDate = '" & today1 & "' " &
    '        "AND EmailSent = 'P' "
    '    )
    '                Dim Mailenable As String = Convert.ToString(dt1.Rows(0)("EmailSent"))
    '                'Dim CreateTime As DateTime = Convert.ToDateTime(dt.Rows(0)("CreateEDate"))
    '                If Mailenable = "P" Then

    '                Dim sb As New StringBuilder()
    '                sb.Append("<h3>Daily Integration Error Report</h3>")
    '                sb.Append("<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;width:100%;font-family:Arial;font-size:12px;'>")

    '                ' Table Header
    '                sb.Append("<tr style='background-color:#f2f2f2;font-weight:bold;'>")
    '                sb.Append("<th>Date</th>")
    '                sb.Append("<th>Error</th>")
    '                'sb.Append("<th>Error ID</th>")
    '                sb.Append("<th>Error Description</th>")
    '                sb.Append("</tr>")

    '                ' Table Rows
    '                For Each r As DataRow In dt1.Rows
    '                    sb.Append("<tr>")
    '                    sb.Append("<td>" & Convert.ToDateTime(r("CreateEDate")).ToString("yyyy-MM-dd") & "</td>")
    '                    sb.Append("<td>" & r("ErrorMsg").ToString() & "</td>")
    '                    'sb.Append("<td>" & r("ErrorId").ToString() & "</td>")
    '                    sb.Append("<td>" & r("ErrorDescription").ToString().Replace(Environment.NewLine, "<br/>") & "</td>")
    '                    sb.Append("</tr>")
    '                Next

    '                sb.Append("</table>")


    '                Dim sent As Boolean = SendEmail("192.168.16.15", 25, False, "", "", "sapint@makevale.com", "Tejas@TechEnable.io", "Daily Integration Error Report", sb.ToString())

    '                    ' Dim sent As Boolean '= SendLocalMail("Daily Error Report Data", sb.ToString()) 'SendEmail("smtp.mail.com", 587, True, "Tejas@TechEnable.io", "Xeoy8105&(", "sapint@makevale.com", "it@makevale.com", "Daily Error Report Data", sb.ToString())



    '                    If sent Then
    '                        ExcuteNonQuery_SciNote(
    '                "UPDATE Mailerrorlog SET EmailSent = 'Y' " &
    '                "WHERE CreateEDate = '" & today & "' AND EmailSent = 'P'",
    '                True
    '                )
    '                    End If
    '                Else
    '                    Exit Sub
    '                End If



    '    Catch ex As Exception
    '        Console.WriteLine("Message: " & ex.Message)

    '    Finally
    '    End Try

    'End Sub

    Public Function SendEmail(
        smtpHost As String,
        smtpPort As Integer,
        enableSsl As Boolean,
        smtpUser As String,
        smtpPassword As String,
        fromEmail As String,
        toEmail As String,
         toEmail1 As String,
         toEmail2 As String,
        subject As String,
        body As String,
        Optional isHtml As Boolean = True,
        Optional attachmentPath As String = ""
    ) As Boolean

        Try
            Dim mail As New MailMessage()
            mail.From = New MailAddress(fromEmail)
            mail.To.Add(toEmail)
            mail.Subject = subject
            mail.Body = body
            mail.IsBodyHtml = isHtml

            ' Attachment (optional)
            If attachmentPath <> "" AndAlso IO.File.Exists(attachmentPath) Then
                mail.Attachments.Add(New Attachment(attachmentPath))
            End If

            Dim smtp As New SmtpClient(smtpHost, smtpPort)
            smtp.Credentials = New NetworkCredential(smtpUser, smtpPassword)
            smtp.EnableSsl = enableSsl

            smtp.Send(mail)

            Return True

        Catch ex As Exception
            ' Log error if needed
            Return False
        End Try

    End Function

    Public Sub initSciNoteDB()
        Dim DBQuery = "Create Database SciNoteDBAll"
        ExcuteNonQuery(DBQuery)


        Dim WOO_Order_errorLog = "Create table Mailerrorlog (ErrorId INT IDENTITY(1,1) PRIMARY KEY,CreateEDate Date Not NULL CAST(GETDATE() AS DATE),CreateETime DATETIME Not NULL DEFAULT GETDATE(),DocumentId VARCHAR(100) NULL,ErrorMsg VARCHAR(200) Not NULL,
                         ErrorDescription NVARCHAR(MAX) NOT NULL,EmailSent CHAR(1) DEFAULT 'N',CompanyName varchar(150),ObjectName varchar(150),Objectid int)"
        ExcuteNonQuery_SciNote(WOO_Order_errorLog)


    End Sub







End Module

