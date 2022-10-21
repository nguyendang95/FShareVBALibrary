VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FShareService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_UserEmail As String
Private m_Password As String
Private m_AppKey As String
Private m_UserAgent As String
Private m_Token As String
Private m_SessionId As String
Private m_LoggedIn As Boolean
Private m_LoggedOut As Boolean
Private m_ActiveService As WinHttp.WinHttpRequest

Public Enum CredentialType
    UserInput = 0
    TextFile = 1
End Enum

Public Property Let UserEmail(value As String)
    m_UserEmail = value
End Property

Public Property Get UserEmail() As String
    UserEmail = m_UserEmail
End Property

Public Property Let Password(value As String)
    m_Password = value
End Property

Public Property Get Password() As String
    Password = m_Password
End Property

Public Property Let AppKey(value As String)
    m_AppKey = value
End Property

Public Property Get AppKey() As String
    AppKey = m_AppKey
End Property

Public Property Let UserAgent(value As String)
    m_UserAgent = value
End Property

Public Property Get UserAgent() As String
    UserAgent = m_UserAgent
End Property

Public Property Let Token(value As String)
    m_Token = value
End Property

Public Property Get Token() As String
    Token = m_Token
End Property

Public Property Let SessionId(value As String)
    m_SessionId = value
End Property

Public Property Get SessionId() As String
    SessionId = m_SessionId
End Property

Public Property Let LoggedIn(value As Boolean)
    m_LoggedIn = value
End Property

Public Property Get LoggedIn() As Boolean
    LoggedIn = m_LoggedIn
End Property

Public Property Let LoggedOut(value As Boolean)
    m_LoggedOut = value
End Property

Public Property Get LoggedOut() As Boolean
    LoggedOut = m_LoggedOut
End Property

Private Property Get ActiveService() As WinHttp.WinHttpRequest
    Set ActiveService = m_ActiveService
End Property

Private Property Set ActiveService(value As WinHttp.WinHttpRequest)
    Set m_ActiveService = value
End Property

Public Function Login(Optional CredentialType As CredentialType = UserInput, Optional CredentialFile As String = vbNullString) As FShareFileManager
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim objFileManager As FShareFileManager
    Dim strRequestBody As String
    Dim strUserEmail As String
    Dim strPassword As String
    Dim strAppKey As String
    Dim strUserAgent As String
    Dim objFSO As Scripting.FileSystemObject
    Dim objTextStream As Scripting.TextStream
    If CredentialType = TextFile Then
        Set objFSO = New Scripting.FileSystemObject
        Set objTextStream = objFSO.OpenTextFile(CredentialFile)
        With objTextStream
            strUserEmail = .ReadLine
            strPassword = .ReadLine
            strAppKey = .ReadLine
            strUserAgent = .ReadLine
            .Close
        End With
        UserEmail = strUserEmail
        Password = strPassword
        AppKey = strAppKey
        UserAgent = strUserAgent
    Else
        strUserEmail = UserEmail
        strPassword = Password
        strAppKey = AppKey
        strUserAgent = UserAgent
    End If
    strRequestBody = "{ " & _
                Quote("user_email") & ": " & Quote(UserEmail) & ", " & _
                Quote("password") & ": " & Quote(Password) & ", " & _
                Quote("app_key") & ": " & Quote(AppKey) & " " & _
                "}"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/user/login"
        .SetRequestHeader "User-Agent", UserAgent
        .SetRequestHeader "Accept", "application/json"
        .Send strRequestBody
        If .Status = 200 Then
            Set ActiveService = objWinHTTP
            LoggedIn = True
            LoggedOut = False
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Token = objDict.Item("token")
            SessionId = objDict.Item("session_id")
            Set objFileManager = New FShareFileManager
            With objFileManager
                .SessionId = objDict.Item("session_id")
                .UserAgent = UserAgent
                .Token = objDict.Item("token")
                Set .ActiveService = ActiveService
            End With
        ElseIf .Status <> 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Err.Raise vbObjectError + .Status, , objDict.Item("msg")
        End If
    End With
    Set Login = objFileManager
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objFileManager = Nothing
    Set objFSO = Nothing
    Set objTextStream = Nothing
End Function

Public Sub LogOut()
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/user/logout"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "Accept", "application/json"
        .Send
        LoggedIn = False
        LoggedOut = True
        Token = vbNullString
        SessionId = vbNullString
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Function GetAccountInfo() As FShareAccountInfo
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim objAccountInfo As FShareAccountInfo
    Dim strRequestBody As String
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/user/get"
        .SetRequestHeader "User-Agent", UserAgent
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "Accept", "application/json"
        .Send
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set objAccountInfo = New FShareAccountInfo
            With objAccountInfo
                .Id = objDict.Item("id")
                .Level = objDict.Item("level")
                .Email = objDict.Item("email")
                .Name = objDict.Item("name")
                .Phone = objDict.Item("phone")
                .Birthday = objDict.Item("birthday")
                .Address = objDict.Item("address")
                .IDCard = objDict.Item("id_card")
                .City = objDict.Item("city")
                .Email = objDict.Item("email")
                .JoinDate = objDict.Item("joindate")
                .TotalPoints = CLng(objDict.Item("totalpoints"))
                .ExpireVip = objDict.Item("expiry_vip")
                .Traffic = CDec(BytesToGigabytes(objDict.Item("traffic")))
                .TrafficUsed = CDec(BytesToGigabytes(objDict.Item("traffic_used")))
                .Webspace = CDec(BytesToGigabytes(objDict.Item("webspace")))
                .WebspaceUsed = CDec(BytesToGigabytes(objDict.Item("webspace_used")))
                .WebspaceSecure = CDec(BytesToGigabytes(objDict.Item("webspace_secure")))
                .WebspaceSecureUsed = CDec(BytesToGigabytes(objDict.Item("webspace_secure_used")))
                .Amount = CDec(objDict.Item("amount"))
                .DownloadTimeAvailable = objDict.Item("dl_time_avail")
                .AccountType = objDict.Item("account_type")
                .JobName = objDict.Item("job_name")
                .Occupation = objDict.Item("occupation")
                .Country = objDict.Item("country")
                .CountryName = objDict.Item("country_name")
                .StatusTelesalePrepaid = objDict.Item("status_telesale_prepaid")
            End With
            If objDict.Item("gender") = "1" Then objAccountInfo.Gender = "Male" Else objAccountInfo.Gender = "Female"
            Set GetAccountInfo = objAccountInfo
            Else: Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Function

Private Function Quote(Text As String) As String
    Quote = Chr(34) & Text & Chr(34)
End Function

Private Function BytesToGigabytes(Bytes As Variant) As Variant
    BytesToGigabytes = CDec(Round(CDec(Bytes) / 1024 / 1024 / 1024, 1))
End Function
