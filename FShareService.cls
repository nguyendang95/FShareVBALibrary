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

Public Function Login() As FShareFileManager
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim objFileManager As FShareFileManager
    Dim strRequestBody As String
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
            LoggedIn = True
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Token = objDict.Item("token")
            SessionId = objDict.Item("session_id")
            Set objFileManager = New FShareFileManager
            With objFileManager
                .SessionId = objDict.Item("session_id")
                .UserAgent = UserAgent
            End With
        ElseIf .Status <> 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Err.Raise vbObjectError + CInt(objDict.Item("code")), , objDict.Item("msg")
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objFileManager = Nothing
End Function

Public Sub LogOut()
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                Quote("session_id") & ": " & Quote(SessionId) & " " & _
                "}"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/user/logout"
        .SetRequestHeader "User-Agent", UserAgent
        .SetRequestHeader "Accept", "application/json"
        .Send strRequestBody
        If .Status = 201 Then
            Err.Raise vbObjectError + 201, , "Not logged in yet!"
        ElseIf .Status = 200 Then
            LoggedOut = True
            Token = vbNullString
            SessionId = vbNullString
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Function UserInfo() As FShareAccountInfo
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim objAccountInfo As FShareAccountInfo
    Dim strRequestBody As String
    Set objWinHTTP = New WinHttp.WinHttpRequest
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
                .Name = objDict.Item("name")
                .Phone = objDict.Item("phone")
                .Birthday = objDict.Item("birthday")
                .Gender = objDict.Item("gender")
                .Address = objDict.Item("address")
                .IDCard = objDict.Item("id_card")
                .City = objDict.Item("city")
                .Email = objDict.Item("email")
                .JoinDate = objDict.Item("joindate")
                .TotalPoints = CLng(objDict.Item("totalpoints"))
                .ExpireVip = objDict.Item("expiry_vip")
                .Traffic = CDec(objDict.Item("traffic"))
                .TrafficUsed = CDec(objDict.Item("traffic_used"))
                .Webspace = CDec(objDict.Item("webspace"))
                .WebspaceUsed = objDict.Item("webspace_used")
                .WebspaceSecured = CDec(objDict.Item("webspace_secured"))
                .WebspaceSecuredUsed = CDec(objDict.Item("webspace_secured_used"))
                .Amount = CDec(objDict.Item("amount"))
                .DownloadTimeAvailable = objDict.Item("dl_time_avail")
                .AccountType = objDict.Item("VIP")
            End With
            Set UserInfo = objAccountInfo
            Else: Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Function

Private Function Quote(Text As String) As String
    Quote = Chr(34) & Text & Chr(34)
End Function
