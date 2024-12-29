<%
'=============================================================================
' SQLWatchdog v1.0 - Classic ASP SQL Query Performance Monitor
'=============================================================================
' Author    : Anthony Burak Dursun (badursun)
' Website   : https://github.com/badursun/Classic-ASP-SQLWatchdog
' Email     : badursun@gmail.com
' Created   : 29.12.2023
' License   : MIT License
'-----------------------------------------------------------------------------
' A lightweight SQL query performance monitoring and profiling tool designed 
' specifically for Classic ASP applications. This tool helps developers identify 
' slow queries, monitor SQL performance, and optimize database operations with 
' minimal setup and overhead.
'=============================================================================

Class SQLWatchdog
    Private pConn            ' Original database connection
    Private queryLogs        ' Dictionary to store query logs
    Private threshold        ' Threshold for slow query detection (in seconds)
    Private lastError        ' Stores the last error message
    
    Private Sub Class_Initialize()
        Set queryLogs = Server.CreateObject("Scripting.Dictionary")
        threshold = 0.5      ' Default threshold: 500ms
        lastError = ""
    End Sub
    
    Private Sub Class_Terminate()
        Set queryLogs = Nothing
        Set pConn = Nothing  ' Just remove reference, don't close
    End Sub

    '-------------------------------------------------------------------------
    ' Core Methods
    '-------------------------------------------------------------------------
    
    ' SetConnection: Store reference to original connection
    Public Sub SetConnection(byref connection)
        If Not connection Is Nothing Then
            Set pConn = connection
        End If
    End Sub
    
    ' SetThreshold: Set slow query threshold in seconds
    Public Sub SetThreshold(seconds)
        threshold = CDbl(seconds)
    End Sub

    '-------------------------------------------------------------------------
    ' Query Execution & Monitoring
    '-------------------------------------------------------------------------
    
    ' Execute: Run SQL query and monitor performance
    Public Function Execute(sqlQuery)
        Dim startTime, endTime, duration, queryType
        
        ' Clear previous error
        lastError = ""
        
        ' Start timer
        startTime = Timer()
        
        ' Execute query through original connection
        On Error Resume Next
        Set Execute = pConn.Execute(sqlQuery)
        
        If Err.Number <> 0 Then
            lastError = "SQL Error: " & Err.Description
            Err.Clear
            Exit Function
        End If
        On Error Goto 0
        
        ' Calculate duration
        endTime = Timer()
        duration = endTime - startTime
        
        ' Log query
        Dim logEntry
        Set logEntry = Server.CreateObject("Scripting.Dictionary")
            logEntry.Add "query", sqlQuery
            logEntry.Add "start", startTime
            logEntry.Add "duration", duration
            logEntry.Add "type", GetQueryType(sqlQuery)
        
        queryLogs.Add queryLogs.Count, logEntry
        Set logEntry = Nothing
    End Function
    
    ' ExecuteParams: Execute parameterized query
    Public Function ExecuteParams(sqlQuery, params)
        Dim cmd, i, startTime, endTime, duration
        
        ' Clear previous error
        lastError = ""
        
        ' Create command object
        Set cmd = Server.CreateObject("ADODB.Command")
        With cmd
            .ActiveConnection = pConn
            .CommandText = sqlQuery
            .CommandType = 1 ' adCmdText
            
            ' Add parameters
            If IsArray(params) Then
                For i = 0 To UBound(params)
                    cmd.Parameters.Append cmd.CreateParameter("p" & i, GetParamType(params(i)), 1, -1, params(i))
                Next
            End If
        End With
        
        ' Start timer and execute
        startTime = Timer()
        
        On Error Resume Next
        Set ExecuteParams = cmd.Execute()
        
        If Err.Number <> 0 Then
            lastError = "SQL Error: " & Err.Description
            Err.Clear
            Set cmd = Nothing
            Exit Function
        End If
        On Error Goto 0
        
        ' Calculate duration
        endTime = Timer()
        duration = endTime - startTime
        
        ' Log query
        Dim logEntry
        Set logEntry = Server.CreateObject("Scripting.Dictionary")
        logEntry.Add "query", sqlQuery
        logEntry.Add "start", startTime
        logEntry.Add "duration", duration
        logEntry.Add "type", GetQueryType(sqlQuery)
        
        queryLogs.Add queryLogs.Count, logEntry
        
        Set logEntry = Nothing
        Set cmd = Nothing
    End Function

    '-------------------------------------------------------------------------
    ' Helper Methods
    '-------------------------------------------------------------------------
    
    ' GetQueryType: Determine SQL query type
    Private Function GetQueryType(sqlQuery)
        sqlQuery = UCase(Trim(sqlQuery))
        
        If Left(sqlQuery, 6) = "SELECT" Then
            GetQueryType = "SELECT"
        ElseIf Left(sqlQuery, 6) = "INSERT" Then
            GetQueryType = "INSERT"
        ElseIf Left(sqlQuery, 6) = "UPDATE" Then
            GetQueryType = "UPDATE"
        ElseIf Left(sqlQuery, 6) = "DELETE" Then
            GetQueryType = "DELETE"
        Else
            GetQueryType = "OTHER"
        End If
    End Function
    
    ' GetParamType: Determine parameter type
    Private Function GetParamType(value)
        Select Case VarType(value)
            Case vbInteger, vbLong
                GetParamType = 3 ' adInteger
            Case vbSingle, vbDouble
                GetParamType = 5 ' adDouble
            Case vbDate
                GetParamType = 7 ' adDate
            Case vbString
                GetParamType = 200 ' adVarChar
            Case vbBoolean
                GetParamType = 11 ' adBoolean
            Case Else
                GetParamType = 200 ' adVarChar
        End Select
    End Function

    '-------------------------------------------------------------------------
    ' Transaction Support
    '-------------------------------------------------------------------------
    
    Public Sub BeginTrans()
        On Error Resume Next
        pConn.BeginTrans
        If Err.Number <> 0 Then
            lastError = "Transaction Error: " & Err.Description
            Err.Clear
        End If
        On Error Goto 0
    End Sub
    
    Public Sub CommitTrans()
        On Error Resume Next
        pConn.CommitTrans
        If Err.Number <> 0 Then
            lastError = "Transaction Error: " & Err.Description
            Err.Clear
        End If
        On Error Goto 0
    End Sub
    
    Public Sub RollbackTrans()
        On Error Resume Next
        pConn.RollbackTrans
        If Err.Number <> 0 Then
            lastError = "Transaction Error: " & Err.Description
            Err.Clear
        End If
        On Error Goto 0
    End Sub

    '-------------------------------------------------------------------------
    ' Reporting Methods
    '-------------------------------------------------------------------------
    
    ' GetLastError: Return last error message
    Public Function GetLastError()
        GetLastError = lastError
    End Function
    
    ' ClearLogs: Clear query logs
    Public Sub ClearLogs()
        queryLogs.RemoveAll
    End Sub
    
    ' RenderReport: Generate HTML report of query performance
    Public Function RenderReport(showAll)
        Dim html, key, log, rowClass, statusText
        
        html = "<style>" & _
               "table { width: 100%; border-collapse: collapse; margin: 20px 0; }" & _
               "th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }" & _
               "th { background-color: #f8f9fa; color: #333; }" & _
               ".slow { background-color: #fff3cd; }" & _
               ".error { color: #dc3545; }" & _
               "</style>"
        
        html = html & "<table>" & _
                     "<tr>" & _
                     "<th>Query</th>" & _
                     "<th>Type</th>" & _
                     "<th>Duration</th>" & _
                     "<th>Status</th>" & _
                     "</tr>"
        
        For Each key In queryLogs
            Set log = queryLogs(key)
            
            ' Only show slow queries unless showAll is true
            If showAll Or log("duration") >= threshold Then
                ' Determine row class and status text
                If log("duration") >= threshold Then
                    rowClass = " class='slow'"
                    statusText = "SLOW"
                Else
                    rowClass = ""
                    statusText = "OK"
                End If
                
                ' Build table row
                html = html & "<tr" & rowClass & ">" & _
                             "<td>" & Server.HTMLEncode(log("query")) & "</td>" & _
                             "<td>" & log("type") & "</td>" & _
                             "<td>" & FormatNumber(log("duration"), 3) & "s</td>" & _
                             "<td>" & statusText & "</td>" & _
                             "</tr>"
            End If
        Next
        
        html = html & "</table>"
        RenderReport = html
    End Function

    '-------------------------------------------------------------------------
    ' Connection Property Passthrough
    '-------------------------------------------------------------------------
    
    Public Property Get State()
        If Not pConn Is Nothing Then
            State = pConn.State
        Else
            State = 0 ' adStateClosed
        End If
    End Property

    '-------------------------------------------------------------------------
    ' Connection Methods Passthrough
    '-------------------------------------------------------------------------
    
    Public Sub Close()
        If Not pConn Is Nothing Then
            If pConn.State = 1 Then ' adStateOpen
                pConn.Close
            End If
        End If
    End Sub
End Class
%>
