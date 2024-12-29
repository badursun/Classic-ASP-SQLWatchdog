<!-- #include file="db.asp" -->
<%
'=============================================================================
' Test Database Creator
'=============================================================================
' Description: This script creates test tables and sample data for SQLWatchdog demos
' Author     : Anthony Burak DURSUN
' Date       : 29.12.2023
' License    : MIT
'=============================================================================

'-----------------------------------------------------------------------------
' Step 1: Database Connection
'-----------------------------------------------------------------------------
' Create a new connection and handle any connection errors
Dim Conn
Set Conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
    Conn.Open connectionString
    
If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Connection Error: " & Err.Description & "</p>"
    Response.End
End If
On Error Goto 0

'-----------------------------------------------------------------------------
' Step 2: Database Creation
'-----------------------------------------------------------------------------
' Ensure the database exists and we're using it
Conn.Execute("CREATE DATABASE IF NOT EXISTS " & DBname)
Conn.Execute("USE " & DBname)

'-----------------------------------------------------------------------------
' Step 3: Create Tables
'-----------------------------------------------------------------------------
On Error Resume Next

' Drop existing tables (orders first due to foreign key constraint)
Conn.Execute("DROP TABLE IF EXISTS orders")
Conn.Execute("DROP TABLE IF EXISTS users")

' Create Users table
' - id: Auto-incrementing primary key
' - username: User's display name
' - email: User's email address
' - created_at: Timestamp of user creation
' - status: User's account status (active/inactive)
Conn.Execute("CREATE TABLE users (" & _
    "id INT AUTO_INCREMENT PRIMARY KEY," & _
    "username VARCHAR(50) NOT NULL," & _
    "email VARCHAR(100) NOT NULL," & _
    "created_at DATETIME DEFAULT CURRENT_TIMESTAMP," & _
    "status ENUM('active', 'inactive') DEFAULT 'active'" & _
    ")")

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error creating users table: " & Err.Description & "</p>"
    Response.End
End If

' Create Orders table
' - id: Auto-incrementing primary key
' - user_id: Foreign key to users table
' - total: Order amount
' - order_date: Timestamp of order creation
' - status: Order status (pending/completed/cancelled)
Conn.Execute("CREATE TABLE orders (" & _
    "id INT AUTO_INCREMENT PRIMARY KEY," & _
    "user_id INT NOT NULL," & _
    "total DECIMAL(10,2) NOT NULL," & _
    "order_date DATETIME DEFAULT CURRENT_TIMESTAMP," & _
    "status ENUM('pending', 'completed', 'cancelled') DEFAULT 'pending'," & _
    "FOREIGN KEY (user_id) REFERENCES users(id)" & _
    ")")

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error creating orders table: " & Err.Description & "</p>"
    Response.End
End If

'-----------------------------------------------------------------------------
' Step 4: Insert Sample Data
'-----------------------------------------------------------------------------
' Insert users one by one for better error handling
' User 1: John Doe - Regular user with multiple orders
Conn.Execute("INSERT INTO users (username, email) VALUES " & _
    "('john_doe', 'john@example.com')")

' User 2: Jane Smith - User with high-value orders
Conn.Execute("INSERT INTO users (username, email) VALUES " & _
    "('jane_smith', 'jane@example.com')")

' User 3: Bob Wilson - User with cancelled order
Conn.Execute("INSERT INTO users (username, email) VALUES " & _
    "('bob_wilson', 'bob@example.com')")

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error inserting users: " & Err.Description & "</p>"
    Response.End
End If

' Insert orders one by one
' Orders for John Doe (user_id = 1)
Conn.Execute("INSERT INTO orders (user_id, total, status) VALUES (1, 150.50, 'completed')")  ' Completed order
Conn.Execute("INSERT INTO orders (user_id, total, status) VALUES (1, 75.25, 'pending')")     ' Pending order

' Order for Jane Smith (user_id = 2)
Conn.Execute("INSERT INTO orders (user_id, total, status) VALUES (2, 200.00, 'completed')")  ' High-value order

' Order for Bob Wilson (user_id = 3)
Conn.Execute("INSERT INTO orders (user_id, total, status) VALUES (3, 50.75, 'cancelled')")   ' Cancelled order

If Err.Number <> 0 Then
    Response.Write "<p style='color:red'>Error inserting orders: " & Err.Description & "</p>"
    Response.End
End If

On Error Goto 0

'-----------------------------------------------------------------------------
' Step 5: Success Report
'-----------------------------------------------------------------------------
Response.Write "<h2>Test Database Created Successfully!</h2>"
Response.Write "<h3>Tables Created:</h3>"
Response.Write "<ul>"
Response.Write "<li>users (id, username, email, created_at, status)</li>"
Response.Write "<li>orders (id, user_id, total, order_date, status)</li>"
Response.Write "</ul>"

Response.Write "<h3>Sample Data Inserted:</h3>"
Response.Write "<ul>"
Response.Write "<li>3 users (John Doe, Jane Smith, Bob Wilson)</li>"
Response.Write "<li>4 orders (2 for John, 1 for Jane, 1 for Bob)</li>"
Response.Write "</ul>"

Response.Write "<h3>Test Scenarios Available:</h3>"
Response.Write "<ul>"
Response.Write "<li>Multiple orders per user (John Doe)</li>"
Response.Write "<li>High-value order (Jane Smith)</li>"
Response.Write "<li>Cancelled order (Bob Wilson)</li>"
Response.Write "<li>Different order statuses (completed, pending, cancelled)</li>"
Response.Write "</ul>"

'-----------------------------------------------------------------------------
' Step 6: Clean up
'-----------------------------------------------------------------------------
Conn.Close
Set Conn = Nothing
%>
