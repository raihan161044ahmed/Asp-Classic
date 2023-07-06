<!DOCTYPE html>
<html>
<head>
    <title>User Registration</title>
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
        <h1 class="mt-10">User Registration</h1>
        <form id="registrationForm" method="post">
            <div class="form-group">
                <label for="firstName">First Name:</label>
                <input type="text" class="form-control" id="firstName" name="firstName" required>
            </div>
            <div class="form-group">
                <label for="lastName">Last Name:</label>
                <input type="text" class="form-control" id="lastName" name="lastName" required>
            </div>
            <div class="form-group">
                <label for="email">Email:</label>
                <input type="email" class="form-control" id="email" name="email" required>
                <small id="emailError" class="form-text text-danger"></small>
            </div>
            <div class="form-group">
                <label for="phone">Phone:</label>
                <input type="text" class="form-control" id="phone" name="phone" required>
            </div>
            <div class="form-group">
                <label for="password">Password:</label>
                <input type="password" class="form-control" id="password" name="password" required>
            </div>
            <div class="form-group">
                <label for="gender">Gender:</label>
                <select class="form-control" id="gender" name="gender" required>
                    <option value="">Select Gender</option>
                    <option value="male">Male</option>
                    <option value="female">Female</option>
                    <option value="other">Other</option>
                </select>
            </div>
            <button type="submit" class="btn btn-primary">Register</button>
            <button type="button" class="btn btn-secondary" onclick="location.href='login.asp'">Login</button>
        </form>
    </div>

    <% 
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
        ' Retrieve form data
        Dim firstName, lastName, email, phone, password, gender
        firstName = Request.Form("firstName")
        lastName = Request.Form("lastName")
        email = Request.Form("email")
        phone = Request.Form("phone")
        password = Request.Form("password")
        gender = Request.Form("gender")

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

        ' Prepare the SQL statement to insert data into the database
        Dim strSQLInsert
        strSQLInsert = "INSERT INTO [Users] ([FirstName], [LastName], [Email], [Phone], [Password], [Gender]) VALUES ('" & Replace(firstName, "'", "''") & "', '" & Replace(lastName, "'", "''") & "', '" & Replace(email, "'", "''") & "', '" & Replace(phone, "'", "''") & "', '" & Replace(password, "'", "''") & "', '" & Replace(gender, "'", "''") & "')"

        ' Execute the SQL insert statement
        Response.Write "User Registered Successfully <br>"
        conn.Execute strSQLInsert

        ' Check for any errors during the insert operation
        If Err.Number <> 0 Then
            Response.Write "An error occurred while saving the form data."
            Response.End
        End If

        ' Close the database connection
        conn.Close
        Set conn = Nothing
    End If
    %>
</body>
</html>
