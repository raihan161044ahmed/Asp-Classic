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
    
    ' Delete the user record based on the ID
    Dim strSQLDelete
    strSQLDelete = "DELETE FROM [Users] WHERE ID=" & userId
    
    On Error Resume Next
    
    ' Execute the SQL delete statement
    conn.Execute(strSQLDelete)
    
    If Err.Number <> 0 Then
        Response.Write "An error occurred while deleting the user record."
    Else
        ' Redirect to the home page or a success message
        Response.Redirect "show.asp"
    End If
    
    On Error GoTo 0
Else
    ' ID parameter not provided
    Response.Write "Invalid request."
End If
%>
