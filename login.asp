<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' Retrieve form data
    Dim email, password
    email = Request.Form("email")
    password = Request.Form("password")

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

    ' Prepare the SQL statement to retrieve the user from the database
    Dim strSQLSelect
    strSQLSelect = "SELECT * FROM [Users] WHERE [Email] = '" & Replace(email, "'", "''") & "' AND [Password] = '" & Replace(password, "'", "''") & "'"

    ' Execute the SQL select statement
    Dim rs
    Set rs = conn.Execute(strSQLSelect)

    ' Check if the user exists
    If rs.EOF Then
        Response.Write "Invalid email or password. Please try again."
        rs.Close
        conn.Close
        Set conn = Nothing
        Response.End
    Else
        Response.Write "Login successful"
        
        ' Set session variables
        Session("LoggedIn") = True
        Session("UserEmail") = email

        ' Redirect to home page
        Response.Redirect "home.asp"

        rs.Close
        conn.Close
        Set conn = Nothing
        Response.End
    End If
End If
%>
<!DOCTYPE html>
<html>
<head>
    <title>User Login</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            padding: 20px;
        }
        .container {
            max-width: 500px;
            margin: 0 auto;
            padding: 20px;
            border: 1px solid #ccc;
            border-radius: 5px;
            background-color: #f7f7f7;
        }
    </style>
</head>
<body>
    <div class="container mt-5">
        <h1 class="mt-10">User Login</h1>
        <form id="loginForm" method="post">
            <div class="form-group">
                <label for="email">Email:</label>
                <input type="email" class="form-control" id="email" name="email" required>
            </div>
            <div class="form-group">
                <label for="password">Password:</label>
                <input type="password" class="form-control" id="password" name="password" required>
            </div>
            <button type="submit" class="btn btn-primary">Login</button>
        </form>
        <p class="mt-3">Not a user? <a href="registration.asp">Register here</a>.</p>
        
    </div>
</body>
</html>
