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

' Check if the ID parameter is provided
If Request.QueryString("id") <> "" Then
    ' Get the ID parameter value
    Dim userId
    userId = Request.QueryString("id")
    
    ' Retrieve the user record based on the ID
    Dim strSQLSelect
    strSQLSelect = "SELECT * FROM [Users] WHERE ID=" & userId
    
    ' Execute the SQL select statement
    Dim rs
    Set rs = conn.Execute(strSQLSelect)
    
    ' Check if the user record exists
    If Not rs.EOF Then
        ' Display the edit form with the user details
        If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
            ' Retrieve form data
            Dim firstName, lastName, email, phone, gender
            firstName = Request.Form("firstName")
            lastName = Request.Form("lastName")
            email = Request.Form("email")
            phone = Request.Form("phone")
            gender = Request.Form("gender")
            
            ' Prepare the SQL statement to update the user record
            Dim strSQLUpdate
            strSQLUpdate = "UPDATE [Users] SET [FirstName]='" & Replace(firstName, "'", "''") & "', [LastName]='" & Replace(lastName, "'", "''") & "', [Email]='" & Replace(email, "'", "''") & "', [Phone]='" & Replace(phone, "'", "''") & "', [Gender]='" & Replace(gender, "'", "''") & "' WHERE ID=" & userId
            
            ' Execute the SQL update statement
            conn.Execute(strSQLUpdate)
            
            ' Check for any errors during the update operation
            If Err.Number <> 0 Then
                Response.Write "An error occurred while updating the user record."
                Response.End
            Else
                ' Redirect to show.asp after saving changes
                Response.Redirect "show.asp"
            End If
        Else
            ' Render the edit form
            %>
            <!DOCTYPE html>
            <html>
            <head>
                <title>Edit User</title>
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
                    <h1>Edit User</h1>
                    <form id="editForm" method="post">
                        <input type="hidden" name="id" value="<%= rs("ID") %>">
                        <div class="form-group">
                            <label for="firstName">First Name:</label>
                            <input type="text" class="form-control" id="firstName" name="firstName" value="<%= rs("FirstName") %>" required>
                        </div>
                        <div class="form-group">
                            <label for="lastName">Last Name:</label>
                            <input type="text" class="form-control" id="lastName" name="lastName" value="<%= rs("LastName") %>" required>
                            </div>
                        <div class="form-group">
                            <label for="email">Email:</label>
                            <input type="email" class="form-control" id="email" name="email" value="<%= rs("Email") %>" required>
                        </div>
                        <div class="form-group">
                            <label for="phone">Phone:</label>
                            <input type="text" class="form-control" id="phone" name="phone" value="<%= rs("Phone") %>" required>
                        </div>
                        <div class="form-group">
                            <label for="gender">Gender:</label>
                            <select class="form-control" id="gender" name="gender" required>
                                <option value="male" <%= IIf(rs("Gender") = "male", "selected", "") %>>Male</option>
                                <option value="female" <%= IIf(rs("Gender") = "female", "selected", "") %>>Female</option>
                                <option value="other" <%= IIf(rs("Gender") = "other", "selected", "") %>>Other</option>
                            </select>
                        </div>
                        <button type="submit" class="btn btn-primary">Save Changes</button>
                        <button type="button" class="btn btn-secondary" onclick="location.href='show.asp'">Cancel</button>
                    </form>
                </div>
            </body>
            </html>
            <%
        End If
    Else
        ' User record not found
        Response.Write "User record not found."
    End If
    
    rs.Close
Else
    ' ID parameter not provided
    Response.Write "Invalid request."
End If

' Close the database connection
conn.Close
Set conn = Nothing
%>
