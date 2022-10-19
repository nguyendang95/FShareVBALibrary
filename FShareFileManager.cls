VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FShareFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SessionId As String
Private m_UserAgent As String
Private m_Token As String

Public Enum ChangeFavoriteStatus
    FavoriteEnabled = 1
    FavoriteDisabled = 0
End Enum

Public Enum FileOrFolder
    FileAndFolder = 0
    FolderOnly = 1
End Enum

Public Property Let SessionId(value As String)
    m_SessionId = value
End Property

Public Property Get SessionId() As String
    SessionId = m_SessionId
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

Public Sub DownloadFile(URL As String, Optional Password As String = vbNullString, Optional ZipFlag As Integer = 0)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("url") & ": " & Quote(URL) & ", " & _
                    Quote("password") & ": " & Quote(Password) & ", " & _
                    Quote("url") & ": " & Quote(URL) & ", " & _
                    Quote("token") & ": " & Quote(Token) & ", " & _
                    Quote("zipflag") & ": " & CStr(ZipFlag)
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/session/download"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Call DownloadFileFromURL(objDict.Item("location"))
        Else: Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Private Sub DownloadFileFromURL(URL As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objFileStream As ADODB.Stream
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "GET", URL
        .Send
        If .Status = 200 Then
            Set objFileStream = New ADODB.Stream
            With objFileStream
                .Open
                .Type = adTypeBinary
                .Write objWinHTTP.ResponseBody
                .SaveToFile GetFileNameFromURL(URL), adSaveCreateOverWrite
                .Close
            End With
        End If
    End With
    Set objWinHTTP = Nothing
    Set objFileStream = Nothing
End Sub

Public Function GetUsersFileOrFolderList(PageIndex As Integer, Ext As String, Optional Path As String = vbNullString, Optional DirOnly As FileOrFolder = FileAndFolder, Optional Limit As Integer = 100) As FShareFileOrFolderInfo
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim objFileOrFolderInfo As FShareFileOrFolderInfo
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/fileops/list?ext=" & Ext & "?path=" & Path & "?dirOnly=" & CStr(DirOnly) & "?limit=" & CStr(Limit)
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set objFileOrFolderInfo = New FShareFileOrFolderInfo
            With objFileOrFolderInfo
                .Id = objDict.Item("id")
                .LinkCode = objDict.Item("linkcode")
                .Name = objDict.Item("name")
                .Secure = objDict.Item("secure")
                .DirectLink = objDict.Item("directlink")
                .FileOrFolderType = objDict.Item("type")
                .Path = objDict.Item("path")
                .Size = CDec(objDict.Item("size"))
                .DownloadCount = CDec(objDict.Item("downloadcount"))
                .MimeType = objDict.Item("mimetype")
                .Created = objDict.Item("created")
                .PWD = objDict.Item("pwd")
                .AllowFollow = objDict.Item("allow_follow")
                .NumFollower = CLng(objDict.Item("num_follower"))
            End With
            Set GetUsersFileOrFolderList = objFileOrFolderInfo
        Else: Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objFileOrFolderInfo = Nothing
End Function

Public Function GetTotalFileInFolder(FolderURL As String, Optional HaveFile As Boolean = False) As Long
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    Dim strHaveFile As String
    If HaveFile Then
        strHaveFile = "true"
    Else: strHaveFile = "false"
    End If
    strRequestBody = "{ " & _
                    Quote("token") & ": " & Quote(Token) & ", " & _
                    Quote("url") & ": " & Quote(FolderURL) & ", " & _
                    Quote("have_file") & ": " & Quote(strHaveFile) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/getTotalFileInFolder"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            GetTotalFileInFolder = objDict.Item("total")
        Else: Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Function

Public Function CreateFolder(FolderName As String, Optional InDir As String = "0") As String
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("name") & ": " & Quote(FolderName) & ", " & _
                    Quote("token") & ": " & Quote(Token) & ", " & _
                    Quote("in_dir") & ": " & Quote(InDir) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/createfolder"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            CreateFolder = objDict.Item("folder")("linkcode")
        Else: Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Function

Public Sub RenameFileOrFolder(NewName As String, LinkCode As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("new_name") & ": " & Quote(NewName) & ", " & _
                    Quote("file") & ": " & Quote(LinkCode) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/rename"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub MoveFileOrFolder(FileOrFolderLinkCodes As String, Optional DestinationFileOrFolderLinkCode As String = "0", Optional InDir As String = "0")
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (FileOrFolderLinkCodes) & " ] " & _
                    Quote("token") & ": " & Quote(Token) & ", " & _
                    Quote("to") & ": " & Quote(DestinationFileOrFolderLinkCode) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/move"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub DeleteFileOrFolder(FileOrFolderLinkCodes As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (FileOrFolderLinkCodes) & " ] " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/delete"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub CreateFilePass(FileLinkCodes As String, Optional Password As String = vbNullString)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (FileLinkCodes) & " ] " & _
                    Quote("pass") & ": " & Quote(Password) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/createFilePass"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub ChangeSecure(FileLinkCodes As String, Optional Status As Integer = 0)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (FileLinkCodes) & " ] " & _
                    Quote("status") & ": " & CStr(Status) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/changeSecure"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub SetDirectLink(FileLinkCodes As String, Optional Status As Integer = 0)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (FileLinkCodes) & " ] " & _
                    Quote("status") & ": " & CStr(Status) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/setDirectLink"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub Duplicate(FileToDuplicateLinkCode As String, DestinationFolderLinkCode As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("linkcode") & ": " & Quote(FileToDuplicateLinkCode) & ", " & _
                    Quote("path") & ": " & Quote(DestinationFolderLinkCode) & ", " & _
                    Quote("confirm") & ": true, " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/duplicate"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Function GetPublicFolderList(PageIndex As Integer, URL As String, Optional Path As String = vbNullString, Optional DirOnly As Integer = 0, Optional Limit As Integer = 100) As FShareFileOrFolderInfo
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim objFileOrFolderInfo As FShareFileOrFolderInfo
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("pageIndex") & ": " & Quote(CStr(PageIndex)) & ", " & _
                    Quote("dirOnly") & ": " & Quote(CStr(DirOnly)) & ", " & _
                    Quote("limit") & ": " & Quote(CStr(Limit)) & ", " & _
                    Quote("path") & ": " & Quote(Path) & ", " & _
                    Quote("url") & ": " & Quote(URL) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/getFolderList"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set objFileOrFolderInfo = New FShareFileOrFolderInfo
            With objFileOrFolderInfo
                .Id = objDict.Item("id")
                .LinkCode = objDict.Item("linkcode")
                .Name = objDict.Item("name")
                .Secure = objDict.Item("secure")
                .DirectLink = objDict.Item("directlink")
                .FileOrFolderType = objDict.Item("type")
                .Path = objDict.Item("path")
                .Size = CDec(objDict.Item("size"))
                .DownloadCount = CDec(objDict.Item("downloadcount"))
                .MimeType = objDict.Item("mimetype")
                .Created = objDict.Item("created")
                .PWD = objDict.Item("pwd")
                .AllowFollow = objDict.Item("allow_follow")
                .NumFollower = CLng(objDict.Item("num_follower"))
            End With
            Set GetPublicFolderList = objFileOrFolderInfo
        Else: Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objFileOrFolderInfo = Nothing
End Function

Public Function GetTopFollowMovie() As FShareTopFollowMovie
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim objTopFollowMovie As FShareTopFollowMovie
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/fileops/getTopFollowMovie"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set objTopFollowMovie = New FShareTopFollowMovie
            With objTopFollowMovie
                .Name = objDict.Item("name")
                .LinkCode = objDict.Item("linkcode")
                .ImgLink = objDict.Item("img_link")
                .Followers = objDict.Item("followers")
            End With
            Set GetTopFollowMovie = objTopFollowMovie
        Else: Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objTopFollowMovie = Nothing
End Function

Public Function GetListFollow()
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim objListFollow As FShareListFollow
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/fileops/getListFollow"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set objListFollow = New FShareListFollow
            With objListFollow
                .Name = objDict.Item("name")
                .LinkCode = objDict.Item("linkcode")
                .NewFile = objDict.Item("newfile")
                .Followers = objDict.Item("followers")
                .CreateTime = objDict.Item("create_time")
            End With
            Set GetListFollow = objListFollow
        Else: Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objListFollow = Nothing
End Function

Public Sub FollowFolder(URLFolder As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("link") & ": " & Quote(URLFolder) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/followFolder"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 409 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Err.Raise vbObjectError + 409, , objDict.Item("msg")
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub UnfollowFolder(URLFolder As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("link") & ": " & Quote(URLFolder) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/unfollowFolder"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + 409, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Function ListFavorite() As FShareFileOrFolderInfo
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Scripting.Dictionary
    Dim objFileOrFolderInfo As FShareFileOrFolderInfo
    Dim strRequestBody As String
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/fileops/listFavorite"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set objFileOrFolderInfo = New FShareFileOrFolderInfo
            With objFileOrFolderInfo
                .Id = objDict.Item("id")
                .LinkCode = objDict.Item("linkcode")
                .Name = objDict.Item("name")
                .Secure = objDict.Item("secure")
                .DirectLink = objDict.Item("directlink")
                .FileOrFolderType = objDict.Item("type")
                .Path = objDict.Item("path")
                .Size = CDec(objDict.Item("size"))
                .DownloadCount = CDec(objDict.Item("downloadcount"))
                .MimeType = objDict.Item("mimetype")
                .Created = objDict.Item("created")
                .PWD = objDict.Item("pwd")
                .AllowFollow = objDict.Item("allow_follow")
                .NumFollower = CLng(objDict.Item("num_follower"))
            End With
            Set ListFavorite = objFileOrFolderInfo
        Else: Err.Raise vbObjectError + 201, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objFileOrFolderInfo = Nothing
End Function

Public Sub ChangeFavorite(LinkCodes As String, Optional Status As ChangeFavoriteStatus = FavoriteDisabled)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim strRequestBody As String
    Dim strStatus As String
    Select Case Status
        Case FavoriteEnabled: strStatus = 1
        Case FavoriteDisabled: strStatus = 0
    End Select
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (LinkCodes) & " ] " & _
                    Quote("status") & ": " & strStatus & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = New WinHttp.WinHttpRequest
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/changeFavorite"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + 409, , .ResponseText
        End If
    End With
    Set objWinHTTP = Nothing
End Sub

Private Function AddBulkFileOrFolderLinkCodes(LinkCodes As String) As String
    Dim arrLinkCodes() As String
    Dim i As Long
    arrLinkCodes = Split(LinkCodes, ",")
    For i = 0 To UBound(arrLinkCodes)
        arrLinkCodes(i) = Quote(arrLinkCodes(i))
    Next
    AddBulkFileOrFolderLinkCodes = Join(arrLinkCodes, ",")
End Function

Private Function GetFileNameFromURL(URL As String) As String
    GetFileNameFromURL = Mid(URL, InStrRev(URL, "/") + 1, Len(URL) - InStrRev(URL, "/") + 1)
End Function

Private Function Quote(Text As String) As String
    Quote = Chr(34) & Text & Chr(34)
End Function
