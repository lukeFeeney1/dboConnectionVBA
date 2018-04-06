VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dboConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------
'Class:     dbConnection
'Purpose:   ADO Abstraction Class
'------------------------------------------------------------
'
Private connectStr As String
Private DB As ADODB.Connection

Public Property Get Connected() As Variant

    If DB.State = adStateOpen Then
        Connected = connectStr
    Else
        Connected = False
    End If
    
End Property

'------------------------------------------------------------
'Function:  Connect
'Purpose:   Accepts connection variables to set up connection (default is SQL Server)
'-------------------------------------------------------------
'
Public Function Connect(ByVal dataSource As String, ByVal initialCatalog As String, ByVal UserName As String, ByVal myPassword As String) As Boolean
    
    connectStr = "Provider=SQLOLEDB;Data Source=" & dataSource & ";Intitial Catalog=" & intialCatalog & ";" & _
                    "User Id=" & UserName & ";Password=" & myPassword & ";Trusted_Connection=No"
    Connect = dbOpen
    
End Function

'------------------------------------------------------------
'Function:  Query
'Purpose:   Generates parameters for query(optional) and either "pulls" data from db or "pushes" data to db
'-------------------------------------------------------------
'
Public Function Query(ByVal querySQL As String, ByVal queryType As String, Optional parameters As Variant) As ADODB.Recordset

    Dim dbQuery As ADODB.Command
    Dim parameter As ADODB.parameter
    Dim output As ADODB.Recordset
    Dim param As Variant
    
    If DB.State <> adStateOpen Then
        Set Query = Nothing
    Else
        Set dbQuery = New ADODB.Command
        dbQuery.ActiveConnection = DB
        dbQuery.CommandText = querySQL
        
        'optional - if parameters are passed in, sets up parameters for query
        If Not IsMissing(parameters) Then
            For Each param In parameters
                'depending on datatype, set size and data type
                If TypeName(param) = "Integer" Then
                    Set parameter = dbQuery.CreateParameter(, adInteger, adParamInput, , param)
                ElseIf TypeName(param) = "Double" Then
                    Set parameter = dbQuery.CreateParameter(, adDouble, adParamInput, , param)
                Else
                    Set parameter = dbQuery.CreateParameter(, adVarChar, adParamInput, Len(param), param)
                End If
                
                dbQuery.parameters.Append parameter
            Next
            
            Set parameter = Nothing
            
        End If
        
        If queryType = "Push" Then
            dbQuery.Execute
        ElseIf queryType = "Pull" Then
            Set output = New ADODB.Recordset
            output.Open dbQuery
            
            If output.EOF Then
                Set Query = Nothing
            Else
                Set Query = output
            End If
            
        End If
        
        Set output = Nothing
        Set parameter = Nothing
        Set dbQuery = Nothing
        
    End If
    
End Function

'------------------------------------------------------------
'Function:  Class_Initialize
'Purpose:   Resets db connection
'-------------------------------------------------------------
'
Private Sub Class_Initialize()

    Set DB = Nothing
    
End Sub


'------------------------------------------------------------
'Procedure:  Class_Terminate
'Purpose:    Runs when class is set to nothing and closes db connection
'-------------------------------------------------------------
'
Private Sub Class_Terminate()

    If Not DB Is Nothing Then
        If DB.State = adStateOpen Then DB.Close
    End If
    
    Set DB = Nothing
    
End Sub

'------------------------------------------------------------
'Procedure:  dbOpen
'Purpose:    Open database using connection string from Connect function
'-------------------------------------------------------------
'
Private Function dbOpen() As Boolean

    On Error Resume Next

    Set DB = New ADODB.Connection
    DB.Open connectStr
    
    dbOpen = (DB.State = adStateOpen)
    
End Function


