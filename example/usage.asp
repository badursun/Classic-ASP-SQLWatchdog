<!-- #include file="db.asp" -->
<!-- #include file="../sqlwatchdog.asp" -->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Classic ASPSQLWatchdog Example</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <div class="container">
<%
'=============================================================================
' SQLWatchdog Example Usage
'=============================================================================
' Description: This script demonstrates various SQL queries and their monitoring
' Author     : Anthony Burak DURSUN
' Date       : 29.12.2023
' License    : MIT
'=============================================================================

'-----------------------------------------------------------------------------
' Database Connection & SQLWatchdog Initialization
'-----------------------------------------------------------------------------
' 1. Create and open normal connection
Dim Conn, origConn
Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open connectionString

' 2. Replace connection with SQLWatchdog proxy
Set origConn = Conn                 ' Keep reference to original connection
Set Conn = New SQLWatchdog          ' Create proxy
    Conn.SetConnection origConn     ' Give original connection to proxy
    Conn.SetThreshold 0.3           ' Set slow query threshold (300ms)

' Now all Conn.Execute calls will be monitored automatically

Response.Write "<h2>Classic ASP SQLWatchdog Example Queries</h2>"

'-----------------------------------------------------------------------------
' Example 1: Simple SELECT
'-----------------------------------------------------------------------------
Response.Write "<h3>Example 1: Simple SELECT</h3>"
Set rs = Conn.Execute("SELECT * FROM users WHERE status = 'active'")
If Not rs Is Nothing Then
    Response.Write "<p>Found " & rs.RecordCount & " active users:</p><ul>"
    While Not rs.EOF
        Response.Write "<li>" & rs("username") & " (" & rs("email") & ")</li>"
        rs.MoveNext
    Wend
    Response.Write "</ul>"
    Set rs = Nothing
Else
    Response.Write "<p class='error'>Error in Example 1: " & Conn.GetLastError() & "</p>"
End If

'-----------------------------------------------------------------------------
' Example 2: Slow Query with JOIN (using SLEEP)
'-----------------------------------------------------------------------------
Response.Write "<h3>Example 2: Slow Query with JOIN</h3>"
Set rs = Conn.Execute("SELECT u.username, COUNT(o.id) as order_count, SLEEP(0.4) " & _
                     "FROM users u " & _
                     "LEFT JOIN orders o ON u.id = o.user_id " & _
                     "GROUP BY u.username")
If Not rs Is Nothing Then
    Response.Write "<p>User order statistics:</p><ul>"
    While Not rs.EOF
        Response.Write "<li>" & rs("username") & " has " & rs("order_count") & " orders</li>"
        rs.MoveNext
    Wend
    Response.Write "</ul>"
    Set rs = Nothing
Else
    Response.Write "<p class='error'>Error in Example 2: " & Conn.GetLastError() & "</p>"
End If

'-----------------------------------------------------------------------------
' Example 3: Parameterized Query for High-Value Orders
'-----------------------------------------------------------------------------
Response.Write "<h3>Example 3: Parameterized Query for High-Value Orders</h3>"
Set rs = Conn.ExecuteParams("SELECT u.username, o.total, o.order_date " & _
                           "FROM users u " & _
                           "INNER JOIN orders o ON u.id = o.user_id " & _
                           "WHERE o.total > ? AND o.status = ?", _
                           Array(100, "completed"))
If Not rs Is Nothing Then
    Response.Write "<p>High-value completed orders:</p><ul>"
    While Not rs.EOF
        Response.Write "<li>" & rs("username") & " - $" & rs("total") & _
                      " (Ordered: " & rs("order_date") & ")</li>"
        rs.MoveNext
    Wend
    Response.Write "</ul>"
    Set rs = Nothing
Else
    Response.Write "<p class='error'>Error in Example 3: " & Conn.GetLastError() & "</p>"
End If

'-----------------------------------------------------------------------------
' Example 4: Transaction Example with Multiple Operations
'-----------------------------------------------------------------------------
Response.Write "<h3>Example 4: Transaction Example</h3>"
Conn.BeginTrans
Dim success : success = True
Dim newUserId

' Step 1: Create a new user
Set rs = Conn.ExecuteParams("INSERT INTO users (username, email) VALUES (?, ?)", _
                           Array("test_user", "test@example.com"))
If rs Is Nothing Then
    success = False
    Response.Write "<p class='error'>Error creating user: " & Conn.GetLastError() & "</p>"
End If

' Step 2: Get the new user's ID (with intentional delay)
If success Then
    Set rs = Conn.Execute("SELECT id, SLEEP(0.5) FROM users WHERE username = 'test_user'")
    If Not rs Is Nothing Then
        newUserId = rs("id")
    Else
        success = False
        Response.Write "<p class='error'>Error getting user ID: " & Conn.GetLastError() & "</p>"
    End If
End If

' Step 3: Create orders for the new user
If success Then
    Set rs = Conn.ExecuteParams("INSERT INTO orders (user_id, total, status) VALUES (?, ?, ?)", _
                               Array(newUserId, 99.99, "pending"))
    If rs Is Nothing Then
        success = False
        Response.Write "<p class='error'>Error creating order: " & Conn.GetLastError() & "</p>"
    End If
End If

' Commit or rollback based on success
If success Then
    Conn.CommitTrans
    Response.Write "<p class='success'>Transaction completed successfully!</p>"
Else
    Conn.RollbackTrans
    Response.Write "<p class='error'>Transaction rolled back due to errors.</p>"
End If

'-----------------------------------------------------------------------------
' Example 5: Complex Query with Multiple JOINs and Conditions
'-----------------------------------------------------------------------------
Response.Write "<h3>Example 5: Complex Query</h3>"
Dim complexQuery
complexQuery = "SELECT u.username, " & _
               "       COUNT(DISTINCT o.id) as total_orders, " & _
               "       SUM(o.total) as total_spent, " & _
               "       MAX(o.order_date) as last_order, " & _
               "       SLEEP(0.35) " & _
               "FROM users u " & _
               "LEFT JOIN orders o ON u.id = o.user_id " & _
               "WHERE u.status = 'active' " & _
               "GROUP BY u.username " & _
               "HAVING total_orders > 0 " & _
               "ORDER BY total_spent DESC"

Set rs = Conn.Execute(complexQuery)
If Not rs Is Nothing Then
    Response.Write "<p>Customer spending analysis:</p><ul>"
    While Not rs.EOF
        Response.Write "<li>" & rs("username") & "<br>" & _
                      "Orders: " & rs("total_orders") & "<br>" & _
                      "Total Spent: $" & FormatNumber(rs("total_spent"), 2) & "<br>" & _
                      "Last Order: " & rs("last_order") & "</li>"
        rs.MoveNext
    Wend
    Response.Write "</ul>"
    Set rs = Nothing
Else
    Response.Write "<p class='error'>Error in Example 5: " & Conn.GetLastError() & "</p>"
End If

'-----------------------------------------------------------------------------
' Generate Performance Report
'-----------------------------------------------------------------------------
Response.Write "<h3>Query Performance Report:</h3>"
Response.Write "<p>Threshold for slow queries: 300ms</p>"
Response.Write Conn.RenderReport(True)

'-----------------------------------------------------------------------------
' Clean up
'-----------------------------------------------------------------------------
Conn.ClearLogs()
Conn.Close : Set Conn = Nothing
%>
    </div>
</body>
</html>