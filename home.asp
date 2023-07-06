<% 
If Session("LoggedIn") <> True Then
    ' User is not logged in, redirect to login page
    Response.Redirect "login.asp"
Else
    ' User is logged in, retrieve user profile details
    Dim userEmail
    userEmail = Session("UserEmail")

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

    ' Prepare the SQL statement to retrieve the user profile details
    Dim strSQLSelect
    strSQLSelect = "SELECT * FROM [Users] WHERE [Email] = '" & Replace(userEmail, "'", "''") & "'"

    ' Execute the SQL select statement
    Dim rs
    Set rs = conn.Execute(strSQLSelect)

    ' Check if the user profile exists
    If rs.EOF Then
        Response.Write "User profile not found."
    Else
        ' Display user profile details
        Response.Write "<h1>Welcome, " & rs("FirstName") & " " & rs("LastName") & "!</h1>"
        Response.Write "<p>Email: " & rs("Email") & "</p>"
        Response.Write "<p>Phone: " & rs("Phone") & "</p>"
        Response.Write "<p>Gender: " & rs("Gender") & "</p>"
    End If

    rs.Close
    conn.Close
    Set conn = Nothing
End If
%>

<!DOCTYPE html>
<html>
<head>
    <title>User Profile</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            padding: 20px;
        }
        .container {
            max-width: 500px;
            margin: 0 auto;
        }
    </style>
</head>
<body>
    <div class="container mt-5">
        <!-- User profile details will be displayed here -->
    </div>
</body>
</html>
