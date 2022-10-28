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
Private m_ActiveService As WinHttp.WinHttpRequest

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

Friend Property Get ActiveService() As WinHttp.WinHttpRequest
    Set ActiveService = m_ActiveService
End Property

Friend Property Set ActiveService(value As WinHttp.WinHttpRequest)
    Set m_ActiveService = value
End Property

Public Sub UploadFile(FileName As String, Optional DestinationFolder As String = "/", Optional Secured As Boolean = False)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim strRequestBody As String
    Dim strSecured As String
    If Secured = False Then strSecured = 0 Else strSecured = 1
    Set objFSO = New Scripting.FileSystemObject
    Set objFile = objFSO.GetFile(FileName)
    strRequestBody = "{ " & _
                    Quote("name") & ": " & Quote(objFile.Name) & ", " & _
                    Quote("size") & ": " & Quote(CStr(objFile.Size)) & ", " & _
                    Quote("path") & ": " & Quote(DestinationFolder) & ", " & _
                    Quote("token") & ": " & Quote(Token) & ", " & _
                    Quote("secured") & ": " & strSecured & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/session/upload"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Call UploadFileToURL(objDict.Item("location"), FileName)
        ElseIf .Status = 201 Then Err.Raise vbObjectError + .Status, , "Not logged in yet!"
        Else: Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objFSO = Nothing
    Set objFile = Nothing
End Sub
 
Private Sub UploadFileToURL(URL As String, FileName As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objFileStream As ADODB.Stream
    Dim objFile() As Byte
    Set objFileStream = New ADODB.Stream
        With objFileStream
            .Open
            .Type = adTypeBinary
            .LoadFromFile FileName
            objFile = .Read
            .Close
        End With
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", URL, True
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send objFile
        .WaitForResponse
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objFileStream = Nothing
End Sub

Public Sub DownloadFile(URL As String, DestinationFolder As String, Optional Password As String = vbNullString, Optional ZipFlag As Integer = 0)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("url") & ": " & Quote(URL) & ", " & _
                    Quote("password") & ": " & Quote(Password) & ", " & _
                    Quote("token") & ": " & Quote(Token) & ", " & _
                    Quote("zipflag") & ": " & CStr(ZipFlag) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/session/download"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Call DownloadFileFromURL(objDict.Item("location"), DestinationFolder)
        Else: Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Private Sub DownloadFileFromURL(URL As String, DestinationFolder As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim decTotalSize As Variant
    Dim decCurrentEndByte As Variant, decChunkSize As Variant, decCurrentStartByte As Variant
    Dim strFile As String
    Dim objRandom As Random
    strFile = DestinationFolder & "\" & GetFileNameFromURL(URLDecode(URL))
    decCurrentStartByte = CDec(decCurrentStartByte)
    decChunkSize = CDec(decChunkSize)
    decChunkSize = 256000000
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "HEAD", URL, True
        .Send
        .WaitForResponse
    End With
    If objWinHTTP.Status = 200 Then
        decTotalSize = CDec(objWinHTTP.GetResponseHeader("Content-Length"))
        Set objWinHTTP = Nothing
        decCurrentEndByte = CDec(decCurrentEndByte)
        decCurrentEndByte = 0
        decCurrentStartByte = 0
        If decTotalSize < decChunkSize Then
            decCurrentEndByte = decTotalSize
        Else: decCurrentEndByte = decChunkSize
        End If
        Do While decCurrentEndByte > decCurrentStartByte
            Set objWinHTTP = ActiveService
            With objWinHTTP
                .Open "GET", URL, True
                .SetRequestHeader "session_id", SessionId
                .SetRequestHeader "User-Agent", UserAgent
                .SetRequestHeader "Content-Type", "application/octet-stream"
                .SetRequestHeader "Range", "bytes=" & CStr(decCurrentStartByte) & "-" & CStr(decCurrentEndByte)
                .Send
                .WaitForResponse
            End With
            Set objRandom = New Random
            objRandom.OpenFile strFile, ForAppending
            objRandom.WriteBytes objWinHTTP.ResponseBody
            objRandom.CloseFile
            Set objWinHTTP = Nothing
            Set objRandom = Nothing
            decCurrentStartByte = decCurrentStartByte + decChunkSize + 1
            decCurrentEndByte = decCurrentEndByte + decChunkSize + 1
            If decCurrentEndByte > decTotalSize Then decCurrentEndByte = decTotalSize
        Loop
    Else: Err.Raise vbObjectError + objWinHTTP.Status, , "Not logged in yet!"
    End If
    Set objWinHTTP = Nothing
End Sub

Private Function URLDecode(ByVal strIn As String) As String
    On Error Resume Next
    Dim sl&, tl&, key$, kl&
    sl = 1:    tl = 1: key = "%": kl = Len(key)
    sl = InStr(sl, strIn, key, 1)
    Do While sl > 0
        If (tl = 1 And sl <> 1) Or tl < sl Then
            URLDecode = URLDecode & Mid(strIn, tl, sl - tl)
        End If
        Dim hh$, hi$, hl$, a$
        Select Case UCase(Mid(strIn, sl + kl, 1))
            Case "U"    'Unicode URLEncode
                a = Mid(strIn, sl + kl + 1, 4)
                URLDecode = URLDecode & ChrW("&H" & a)
                sl = sl + 6
            Case "E"    'UTF-8 URLEncode
                hh = Mid(strIn, sl + kl, 2)
                a = Int("&H" & hh)    'ascii?
                If Abs(a) < 128 Then
                    sl = sl + 3
                    URLDecode = URLDecode & Chr(a)
                Else
                    hi = Mid(strIn, sl + 3 + kl, 2)
                    hl = Mid(strIn, sl + 6 + kl, 2)
                    a = ("&H" & hh And &HF) * 2 ^ 12 Or ("&H" & hi And &H3F) * 2 ^ 6 Or ("&H" & hl And &H3F)
                    If a < 0 Then a = a + 65536
                    URLDecode = URLDecode & ChrW(a)
                    sl = sl + 9
                End If
            Case Else    'Asc URLEncode
                hh = Mid(strIn, sl + kl, 2)    '??
                a = Int("&H" & hh)    'ascii?

                If Abs(a) < 128 Then
                    sl = sl + 3
                Else
                    hi = Mid(strIn, sl + 3 + kl, 2)    '??
                    'a = Int("&H" & hh & hi) '?ascii?
                    a = (Int("&H" & hh) - 194) * 64 + Int("&H" & hi)
                    sl = sl + 6
                End If
                URLDecode = URLDecode & ChrW(a)
        End Select
        tl = sl
        sl = InStr(sl, strIn, key, 1)
    Loop
    URLDecode = URLDecode & Mid(strIn, tl)
End Function

Private Function BytesToGigabytes(Bytes As Variant) As Variant
    BytesToGigabytes = CDec(Round(CDec(Bytes) / 1024 / 1024 / 1024, 1))
End Function

Private Function GetFileNameFromURL(URL As String) As String
    Dim strResult As String
    strResult = Mid(URL, InStrRev(URL, "/") + 1, Len(URL) - InStrRev(URL, "/") + 1)
    GetFileNameFromURL = strResult
End Function

Public Function GetUsersFileOrFolderList(PageIndex As Integer, Ext As String, Optional Path As String = vbNullString, Optional DirOnly As FileOrFolder = FileAndFolder, Optional Limit As Integer = 100) As FShareFileOrFolderInfos
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim objFileOrFolderInfo As FShareFileOrFolderInfo
    Dim colFileOrFolderInfos As FShareFileOrFolderInfos
    Dim i As Long
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/fileops/list?ext=" & Ext & "?path=" & Path & "?dirOnly=" & CStr(DirOnly) & "?limit=" & CStr(Limit)
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set colFileOrFolderInfos = New FShareFileOrFolderInfos
            For i = 1 To objDict.Count
                Set objFileOrFolderInfo = New FShareFileOrFolderInfo
                With objFileOrFolderInfo
                    .Id = objDict.Item(i)("id")
                    .LinkCode = objDict.Item(i)("linkcode")
                    .Name = objDict.Item(i)("name")
                    .Secure = objDict.Item(i)("secure")
                    .DirectLink = objDict.Item(i)("directlink")
                    .FileOrFolderType = objDict.Item(i)("type")
                    .Path = objDict.Item(i)("path")
                    .Size = CDec(objDict.Item(i)("size"))
                    .DownloadCount = CDec(objDict.Item(i)("downloadcount"))
                    If IsNull(objDict.Item(i)("mimetype")) Then .MimeType = vbNullString Else .MimeType = objDict.Item(i)("mimetype")
                    If IsNull(objDict.Item(i)("created")) Then .Created = vbNullString Else .Created = objDict.Item(i)("created")
                    If IsNull(objDict.Item(i)("pwd")) Then .PWD = vbNullString Else .PWD = objDict.Item(i)("pwd")
                    .AllowFollow = objDict.Item(i)("allow_follow")
                    .NumFollower = CLng(objDict.Item(i)("num_follower"))
                End With
                colFileOrFolderInfos.Add objFileOrFolderInfo
            Next
            Set GetUsersFileOrFolderList = colFileOrFolderInfos
        Else: Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objFileOrFolderInfo = Nothing
    Set colFileOrFolderInfos = Nothing
End Function

Public Function GetTotalFileInFolder(FolderURL As String, Optional HaveFile As Boolean = False) As Long
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    Dim strHaveFile As String
    If HaveFile Then
        strHaveFile = "true"
    Else: strHaveFile = "false"
    End If
    strRequestBody = "{ " & _
                    Quote("token") & ":" & Quote(Token) & "," & _
                    Quote("url") & ":" & Quote(FolderURL) & "," & _
                    Quote("have_file") & ":" & strHaveFile & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/getTotalFileInFolder"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            If .ResponseText <> vbNullString Then
                Set objDict = JsonConverter.ParseJson(.ResponseText)
                GetTotalFileInFolder = objDict.Item("total")
            Else: GetTotalFileInFolder = 0
            End If
        Else: Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Function

Public Function CreateFolder(FolderName As String, Optional InDir As String = "0") As String
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                        Quote("name") & ": " & Quote(FolderName) & ", " & _
                        Quote("token") & ": " & Quote(Token) & ", " & _
                        Quote("in_dir") & ": " & Quote(InDir) & _
                        " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/createFolder"
        .SetRequestHeader "Content-Type", "application/json;charset=utf-8"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            CreateFolder = objDict.Item("folder")("linkcode")
        Else: Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Function

Public Sub RenameFileOrFolder(NewName As String, LinkCode As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("new_name") & ": " & Quote(NewName) & ", " & _
                    Quote("file") & ": " & Quote(LinkCode) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/rename"
        .SetRequestHeader "Content-Type", "application/json;charset=utf-8"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub MoveFileOrFolder(FileOrFolderLinkCodes As String, Optional DestinationFileOrFolderLinkCode As String = "0", Optional InDir As String = "0")
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (FileOrFolderLinkCodes) & " ], " & _
                    Quote("token") & ": " & Quote(Token) & ", " & _
                    Quote("to") & ": " & Quote(DestinationFileOrFolderLinkCode) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/move"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub DeleteFileOrFolder(FileOrFolderLinkCodes As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ " & _
                    AddBulkFileOrFolderLinkCodes(FileOrFolderLinkCodes) & " ], " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/delete"
        .SetRequestHeader "Content-Type", "application/json;charset=utf-8"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub CreateFilePass(FileLinkCodes As String, Optional Password As String = vbNullString)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (FileLinkCodes) & " ], " & _
                    Quote("pass") & ": " & Quote(Password) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/createFilePass"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json;charset=utf-8"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub ChangeSecure(FileLinkCodes As String, Optional Status As Integer = 0)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (FileLinkCodes) & " ], " & _
                    Quote("status") & ": " & CStr(Status) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/changeSecure"
        .SetRequestHeader "Content-Type", "application/json;charset=utf-8"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub SetDirectLink(FileLinkCodes As String, Optional Status As Integer = 0)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("items") & ": " & "[ "
                    AddBulkFileOrFolderLinkCodes (FileLinkCodes) & " ], " & _
                    Quote("status") & ": " & CStr(Status) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/setDirectLink"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json;charset=utf-8"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub Duplicate(FileToDuplicateLinkCode As String, DestinationFolderLinkCode As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("linkcode") & ": " & Quote(FileToDuplicateLinkCode) & ", " & _
                    Quote("path") & ": " & Quote(DestinationFolderLinkCode) & ", " & _
                    Quote("confirm") & ": true, " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/duplicate"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json;charset=utf-8"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Function GetPublicFolderList(PageIndex As Integer, URL As String, Optional DirOnly As Integer = 0, Optional Limit As Integer = 100) As FShareFileOrFolderInfos
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim objFileOrFolderInfo As FShareFileOrFolderInfo
    Dim colFileOrFolderInfos As FShareFileOrFolderInfos
    Dim strRequestBody As String
    Dim i As Long
    strRequestBody = "{ " & _
                    Quote("pageIndex") & ": " & CStr(PageIndex) & ", " & _
                    Quote("dirOnly") & ": " & CStr(DirOnly) & ", " & _
                    Quote("limit") & ": " & CStr(Limit) & ", " & _
                    Quote("token") & ": " & Quote(Token) & ", " & _
                    Quote("url") & ": " & Quote(URL) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/getFolderList"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set colFileOrFolderInfos = New FShareFileOrFolderInfos
            For i = 1 To objDict.Count
                Set objFileOrFolderInfo = New FShareFileOrFolderInfo
                With objFileOrFolderInfo
                    .Id = objDict.Item(i)("id")
                    .LinkCode = objDict.Item(i)("linkcode")
                    .Name = objDict.Item(i)("name")
                    .Secure = objDict.Item(i)("secure")
                    .DirectLink = objDict.Item(i)("directlink")
                    .FileOrFolderType = objDict.Item(i)("type")
                    .Path = objDict.Item(i)("path")
                    .Size = CDec(objDict.Item(i)("size"))
                    .DownloadCount = CDec(objDict.Item(i)("downloadcount"))
                    If IsNull(objDict.Item(i)("mimetype")) Then .MimeType = vbNullString Else .MimeType = objDict.Item(i)("mimetype")
                    If IsNull(objDict.Item(i)("created")) Then .Created = vbNullString Else .Created = objDict.Item(i)("created")
                    If IsNull(objDict.Item(i)("pwd")) Then .PWD = vbNullString Else .PWD = objDict.Item(i)("pwd")
                    .AllowFollow = objDict.Item(i)("allow_follow")
                    .NumFollower = CLng(objDict.Item(i)("num_follower"))
                End With
                colFileOrFolderInfos.Add objFileOrFolderInfo
            Next
            Set GetPublicFolderList = colFileOrFolderInfos
        Else: Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objFileOrFolderInfo = Nothing
    Set colFileOrFolderInfos = Nothing
End Function

Public Function GetTopFollowMovie() As FshareTopFollowMovies
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim objTopFollowMovie As FShareTopFollowMovie
    Dim colTopFollowMovies As FShareFileOrFolderInfos
    Dim i As Long
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/fileops/getTopFollowMovie"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set colTopFollowMovies = New FshareTopFollowMovies
            For i = 1 To objDict.Count
                Set objTopFollowMovie = New FShareTopFollowMovie
                With objTopFollowMovie
                    .Name = objDict.Item(i)("name")
                    .LinkCode = objDict.Item(i)("linkcode")
                    .ImgLink = objDict.Item(i)("img_link")
                    .Followers = objDict.Item(i)("followers")
                End With
                colTopFollowMovies.Add objTopFollowMovie
            Next
            Set GetTopFollowMovie = colTopFollowMovies
        Else: Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objTopFollowMovie = Nothing
    Set colTopFollowMovies = Nothing
End Function

Public Function GetListFollow() As FShareListFollow
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim objListFollow As FShareListFollow
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/fileops/getListFollow"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send
        If .Status = 200 And .ResponseText <> "[]" Then
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
        ElseIf .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        Else: Set GetListFollow = Nothing
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objListFollow = Nothing
End Function

Public Sub FollowFolder(URLFolder As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("link") & ": " & Quote(URLFolder) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/followFolder"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 409 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Err.Raise vbObjectError + .Status, , objDict.Item("msg")
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Sub UnfollowFolder(URLFolder As String)
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim strRequestBody As String
    strRequestBody = "{ " & _
                    Quote("link") & ": " & Quote(URLFolder) & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/unfollowFolder"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
End Sub

Public Function ListFavorite() As FShareFileOrFolderInfos
    Dim objWinHTTP As WinHttp.WinHttpRequest
    Dim objDict As Object
    Dim objFileOrFolderInfo As FShareFileOrFolderInfo
    Dim colFileOrFolderInfos As FShareFileOrFolderInfos
    Dim strRequestBody As String
    Dim i As Long
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "GET", "https://api.fshare.vn/api/fileops/listFavorite"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status = 200 Then
            Set objDict = JsonConverter.ParseJson(.ResponseText)
            Set colFileOrFolderInfos = New FShareFileOrFolderInfos
            For i = 1 To objDict.Count
                Set objFileOrFolderInfo = New FShareFileOrFolderInfo
                With objFileOrFolderInfo
                    .Id = objDict.Item(i)("id")
                    .LinkCode = objDict.Item(i)("linkcode")
                    .Name = objDict.Item(i)("name")
                    .Secure = objDict.Item(i)("secure")
                    .DirectLink = objDict.Item(i)("directlink")
                    .FileOrFolderType = objDict.Item(i)("type")
                    .Path = objDict.Item(i)("path")
                    .Size = CDec(objDict.Item(i)("size"))
                    .DownloadCount = CDec(objDict.Item(i)("downloadcount"))
                    If IsNull(objDict.Item(i)("mimetype")) Then .MimeType = vbNullString Else .MimeType = objDict.Item(i)("mimetype")
                    If IsNull(objDict.Item(i)("created")) Then .Created = vbNullString Else .Created = objDict.Item(i)("created")
                    If IsNull(objDict.Item(i)("pwd")) Then .PWD = vbNullString Else .PWD = objDict.Item(i)("pwd")
                    .AllowFollow = objDict.Item(i)("allow_follow")
                    .NumFollower = CLng(objDict.Item(i)("num_follower"))
                End With
                colFileOrFolderInfos.Add objFileOrFolderInfo
            Next
            Set ListFavorite = colFileOrFolderInfos
        Else: Err.Raise vbObjectError + .Status, , .StatusText
        End If
    End With
    Set objWinHTTP = Nothing
    Set objDict = Nothing
    Set objFileOrFolderInfo = Nothing
    Set colFileOrFolderInfos = Nothing
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
                    AddBulkFileOrFolderLinkCodes (LinkCodes) & " ], " & _
                    Quote("status") & ": " & strStatus & ", " & _
                    Quote("token") & ": " & Quote(Token) & _
                    " }"
    Set objWinHTTP = ActiveService
    With objWinHTTP
        .Open "POST", "https://api.fshare.vn/api/fileops/changeFavorite"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "session_id", SessionId
        .SetRequestHeader "User-Agent", UserAgent
        .Send strRequestBody
        If .Status <> 200 Then
            Err.Raise vbObjectError + .Status, , .StatusText
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

Private Function Quote(Text As String) As String
    Quote = Chr(34) & Text & Chr(34)
End Function
