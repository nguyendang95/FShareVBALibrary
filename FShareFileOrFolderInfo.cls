VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FShareFileOrFolderInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_LinkCode As String
Private m_Name As String
Private m_Secure As String
Private m_DirectLink As String
Private m_FileOrFolderType As String
Private m_Path As String
Private m_Size As Variant
Private m_DownloadCount As Variant
Private m_MimeType As String
Private m_Created As String
Private m_PWD As String
Private m_AllowFollow As String
Private m_NumFollower As Long

Public Property Let Id(value As String)
    m_Id = value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let LinkCode(value As String)
    m_LinkCode = value
End Property

Public Property Get LinkCode() As String
    LinkCode = m_LinkCode
End Property

Public Property Let Name(value As String)
    m_Name = value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Secure(value As String)
    m_Secure = value
End Property

Public Property Get Secure() As String
    Secure = m_Secure
End Property

Public Property Let DirectLink(value As String)
    m_DirectLink = value
End Property

Public Property Get DirectLink() As String
    DirectLink = m_DirectLink
End Property

Public Property Let FileOrFolderType(value As String)
    m_FileOrFolderType = value
End Property

Public Property Get FileOrFolderType() As String
    FileOrFolderType = m_FileOrFolderType
End Property

Public Property Let Path(value As String)
    m_Path = value
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Size(value As Variant)
    m_Size = value
End Property

Public Property Get Size() As Variant
    Size = m_Size
End Property

Public Property Let DownloadCount(value As Variant)
    m_DownloadCount = value
End Property

Public Property Get DownloadCount() As Variant
    DownloadCount = m_DownloadCount
End Property

Public Property Let MimeType(value As String)
    m_MimeType = value
End Property

Public Property Get MimeType() As String
    MimeType = m_MimeType
End Property

Public Property Let Created(value As String)
    m_Created = value
End Property

Public Property Get Created() As String
    Created = m_Created
End Property

Public Property Let PWD(value As String)
    m_PWD = value
End Property

Public Property Get PWD() As String
    PWD = m_PWD
End Property

Public Property Let AllowFollow(value As String)
    m_AllowFollow = value
End Property

Public Property Get AllowFollow() As String
    AllowFollow = m_AllowFollow
End Property

Public Property Let NumFollower(value As Long)
    m_NumFollower = value
End Property

Public Property Get NumFollower() As Long
    NumFollower = m_Id
End Property
