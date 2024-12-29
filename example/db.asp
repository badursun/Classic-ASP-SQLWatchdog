<%
'-----------------------------------------------------------------------------
' Database connection settings
'-----------------------------------------------------------------------------
Const DBdriver = "MySQL ODBC 3.51 Driver" ' "MySQL ODBC 5.2 ANSI Driver"
Dim DBname : DBname     = "demo_test_db"
Dim DBuser : DBuser     = "demo_test_user"
Dim DBpass : DBpass     = "Z_x95n1i7"
Dim DBserver : DBserver = "localhost"
Dim DBPort : DBPort     = "3306"

'-----------------------------------------------------------------------------
' Create connection string with additional parameters
'-----------------------------------------------------------------------------
Dim connectionString
    connectionString = "DRIVER={"& DBdriver &"};" & _
                      "SERVER="& DBserver &";" & _
                      "PORT="& DBPort &";" & _
                      "DATABASE="& DBname &";" & _
                      "UID="& DBuser &";" & _
                      "PWD="& DBpass &";" & _
                      "OPTION=3"               ' Add OPTION=3 for better memory handling
%>