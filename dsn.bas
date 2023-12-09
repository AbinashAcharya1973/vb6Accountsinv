Attribute VB_Name = "dsn"
' ODBC API
' -- ODBC Commands
Public Const ODBC_ADD_DSN = 1&
Public Const ODBC_CONFIG_DSN = 2&
Public Const ODBC_REMOVE_DSN = 3&
Public Const ODBC_ADD_SYS_DSN = 4&
Public Const ODBC_CONFIG_SYS_DSN = 5&
Public Const ODBC_REMOVE_SYS_DSN = 6&
Public Const ODBC_REMOVE_DEFAULT_DSN = 7&

' -- ODBC Error Codes
Public Const ODBC_ERROR_GENERAL_ERR = 1
Public Const ODBC_ERROR_INVALID_BUFF_LEN = 2
Public Const ODBC_ERROR_INVALID_HWND = 3
Public Const ODBC_ERROR_INVALID_STR = 4
Public Const ODBC_ERROR_INVALID_REQUEST_TYPE = 5
Public Const ODBC_ERROR_COMPONENT_NOT_FOUND = 6
Public Const ODBC_ERROR_INVALID_NAME = 7
Public Const ODBC_ERROR_INVALID_KEYWORD_VALUE = 8
Public Const ODBC_ERROR_INVALID_DSN = 9
Public Const ODBC_ERROR_INVALID_INF = 10
Public Const ODBC_ERROR_REQUEST_FAILED = 11
Public Const ODBC_ERROR_INVALID_PATH = 12
Public Const ODBC_ERROR_LOAD_LIB_FAILED = 13
Public Const ODBC_ERROR_INVALID_PARAM_SEQUENCE = 14
Public Const ODBC_ERROR_INVALID_LOG_FILE = 15
Public Const ODBC_ERROR_USER_CANCELED = 16
Public Const ODBC_ERROR_USAGE_UPDATE_FAILED = 17
Public Const ODBC_ERROR_CREATE_DSN_FAILED = 18
Public Const ODBC_ERROR_WRITING_SYSINFO_FAILED = 19
Public Const ODBC_ERROR_REMOVE_DSN_FAILED = 20
Public Const ODBC_ERROR_OUT_OF_MEM = 21
Public Const ODBC_ERROR_OUTPUT_STRING_TRUNCATED = 22

'API Command to create a Data Source Name, not used in this example
Public Declare Function SQLCreateDataSource Lib "odbccp32.dll" (ByVal hwnd&, ByVal lpszDS$) As Boolean
'API to modify/Edit/Create a Data Source Name
Public Declare Function SQLConfigDataSource Lib "odbccp32.dll" (ByVal hwnd As Long, ByVal fRequest As Integer, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Boolean

Private Sub CreateSQLODBC()
' This function setups a DSN common for remote Database servers
' Such as SQL or Oracle, keep in mind, this isnt a complete listing of parameters

    Dim DSN As String
    Dim Server As String
    Dim Address As String
    Dim Database As String
    Dim Description As String
    Dim Security As String

    'Basically the DSN Name you want to have
    DSN = "DSN=" & Trim(txtDatabase.Text)
    'The IP Addy of the server you want , if this is a remote connection
    Server = "SERVER=" & Trim(txtServer.Text)
    'Same as above
    Address = "ADDRESS=" & Trim(txtServer.Text)
    'The name of the database as known by the DB Server, such as SQL Server
    Database = "DATABASE=" & Trim(txtDatabase.Text)
    'An Optional Description Feild
    Description = "DESCRIPTION=" & Trim(txtdescription.Text)
    'This is optional, if you require a Security mode check the help files
    Security = "NETWORK=dbmssocn"

    'the next couple lines setup the Driver Text , that defines the type of DB Drivers
    ' you are using, if its anything other than the ones I've listed, check your DB
    ' documentation, or check the ODBC settings to see it's names

    'Also you will notice as each string peice is put together, they are seperated by
    'VbNullChar, this gives it a Null seperated array in a sense so that the API Command
    'can use the Parameters

    If OptSQL.Value = True Then
        SqlDriver = "SQL Server"
        SQLParameter = DSN & vbNullChar & Server & vbNullChar & Address & vbNullChar & Security & vbNullChar & _
                       Database & vbNullChar & Description & vbNullChar & vbNullChar
    ElseIf OptOracle.Value = True Then
        SqlDriver = "Oracle73"
        SQLParameter = DSN & vbNullChar & Server & vbNullChar & Database & vbNullChar & _
                       Description & vbNullChar & vbNullChar
    End If

    'calls SQLConfigDataSource , giving it the forms handle, the command to Add a System DSN
    'giving it the Driver name, and then the Null Seperated Parameter listing

    SQLConfigDataSource 0&, ODBC_ADD_SYS_DSN, SqlDriver, SQLParameter
    'Replace 0& with Me.hwnd if you wish for users to further configure settings such as a long

End Sub

Public Sub CreateAccessODBC(db As String, data_source_name As String, dsn_description As String)
'This common setup for an Access Database

    Dim Driver As String
    Dim DSN As String
    Dim Server As String
    Dim Database As String
    Dim Description As String

    DSN = "DSN=" & Trim(data_source_name)
    Database = "DBQ=" & Trim(db)    'this is your physical path to the *.mdb
    Description = "DESCRIPTION=" & Trim(dsn_description)
    AccessParameter = DSN & vbNullChar & Database & vbNullChar & Description & vbNullChar & vbNullChar
    SQLConfigDataSource 0&, ODBC_ADD_SYS_DSN, "Microsoft Access Driver (*.mdb)", AccessParameter
End Sub




