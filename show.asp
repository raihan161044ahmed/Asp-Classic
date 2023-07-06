<!DOCTYPE html>
<html>
<head>
    <title>User Records</title>
     <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            padding: 20px;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
        }
        
        th, td {
            padding: 8px;
            text-align: left;
            border: 1px solid #000;           
        }
        
        th {
            background-color: #d9cdff;
        }
            .btn {
            display: inline-block;
            padding: 6px 16px;
            margin-bottom: 0;
            font-size: 14px;
            font-weight: 500;
            line-height: 1.2;
            text-align: center;
            white-space: nowrap;
            vertical-align: middle;
            border-radius: 4px;
        }
        
    </style>
</head>
<body>
    <h1>User Records</h1>
    <% 
    ' Path to the Access database file
    Dim dbPath
    dbPath = Server.MapPath("crud_db.accdb")

    ' Connection string for Access database
    Dim connStr
    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    ' Create a new connection object
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")

    On Error Resume Next

    ' Open the database connection
    conn.Open connStr

    ' Check for connection errors
    If Err.Number <> 0 Then
        Response.Write "An error occurred while connecting to the database."
        Response.End
    End If

    On Error Goto 0

    ' Prepare the SQL statement to retrieve all user records
    Dim strSQLSelect
    strSQLSelect = "SELECT ID, FirstName, LastName, Email, Phone, Gender FROM [Users]"

    ' Execute the SQL select statement
    Dim rs
    Set rs = conn.Execute(strSQLSelect)

    ' Check if any user records exist
    If rs.EOF Then
        Response.Write "No user records found."
    Else
        ' Display user records in a table
        Response.Write "<table>"
        Response.Write "<tr>"
        Response.Write "<th>ID</th>"
        Response.Write "<th>First Name</th>"
        Response.Write "<th>Last Name</th>"
        Response.Write "<th>Email</th>"
        Response.Write "<th>Phone</th>"
        Response.Write "<th>Gender</th>"
         Response.Write "<th>Actions</th>"
        Response.Write "</tr>"
        
        Do Until rs.EOF
            Response.Write "<tr>"
            Response.Write "<td>" & rs("ID") & "</td>"
            Response.Write "<td>" & rs("FirstName") & "</td>"
            Response.Write "<td>" & rs("LastName") & "</td>"
            Response.Write "<td>" & rs("Email") & "</td>"
            Response.Write "<td>" & rs("Phone") & "</td>"
            Response.Write "<td>" & rs("Gender") & "</td>"
            Response.Write "<td class='text-center'>"
            Response.Write "<a href='edit.asp?id=" & rs("ID") & "' class='btn btn-primary'>Edit</a>"
            Response.Write "<a href='delete.asp?id=" & rs("ID") & "' class='btn btn-danger ml-3'>Delete</a>"
            Response.Write "</td>"
            Response.Write "</tr>"
            rs.MoveNext
        Loop

        Response.Write "</table>"
    End If

    rs.Close
    conn.Close
    Set conn = Nothing
    %>
</body>
</html>
