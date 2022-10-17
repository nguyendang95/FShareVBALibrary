VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FShareListFollow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_LinkCode As String
Private m_NewFile As String
Private m_Followers As String
Private m_CreateTime As String

Public Property Let Name(value As String)
    m_Name = value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let LinkCode(value As String)
    m_LinkCode = value
End Property

Public Property Get LinkCode() As String
    LinkCode = m_LinkCode
End Property

Public Property Let NewFile(value As String)
    m_NewFile = value
End Property

Public Property Get NewFile() As String
    NewFile = m_NewFile
End Property

Public Property Let Followers(value As String)
    m_Followers = value
End Property

Public Property Get Followers() As String
    Followers = m_Followers
End Property

Public Property Let CreateTime(value As String)
    m_CreateTime = value
End Property

Public Property Get CreateTime() As String
    CreateTime = m_CreateTime
End Property

