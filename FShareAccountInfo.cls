VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FShareAccountInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Level As String
Private m_Name As String
Private m_Phone As String
Private m_Birthday As String
Private m_Gender As String
Private m_Address As String
Private m_IDCard As String
Private m_City As String
Private m_Email As String
Private m_JoinDate As String
Private m_TotalPoints As Long
Private m_ExpireVip As String
Private m_Traffic As Variant
Private m_TrafficUsed As Variant
Private m_Webspace As Variant
Private m_WebspaceUsed As Variant
Private m_WebspaceSecured As Variant
Private m_WebspaceSecuredUsed As Variant
Private m_Amount As Variant
Private m_DownloadTimeAvailable As Long
Private m_AccountType As String
Private m_Occupation As String
Private m_JobName As String
Private m_StatusTelesalePrepaid As String
Private m_Country As String
Private m_CountryName As String
Private m_ActiveService As WinHttp.WinHttpRequest

Public Property Let CountryName(value As String)
    m_CountryName = value
End Property

Public Property Get CountryName() As String
    CountryName = m_CountryName
End Property

Public Property Let Country(value As String)
    m_Country = value
End Property

Public Property Get Country() As String
    Country = m_Country
End Property

Public Property Let StatusTelesalePrepaid(value As String)
    m_StatusTelesalePrepaid = value
End Property

Public Property Get StatusTelesalePrepaid() As String
    StatusTelesalePrepaid = m_StatusTelesalePrepaid
End Property

Public Property Let JobName(value As String)
    m_JobName = value
End Property

Public Property Get JobName() As String
    JobName = m_JobName
End Property

Public Property Let Occupation(value As String)
    m_Occupation = value
End Property

Public Property Get Occupation() As String
    Occupation = m_Occupation
End Property

Public Property Let Id(value As String)
    m_Id = value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Level(value As String)
    m_Level = value
End Property

Public Property Get Level() As String
    Level = m_Level
End Property

Public Property Let Name(value As String)
    m_Name = value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Phone(value As String)
    m_Phone = value
End Property

Public Property Get Phone() As String
    Phone = m_Phone
End Property

Public Property Let Birthday(value As String)
    m_Birthday = value
End Property

Public Property Get Birthday() As String
    Birthday = m_Birthday
End Property

Public Property Let Gender(value As String)
    m_Gender = value
End Property

Public Property Get Gender() As String
    Gender = m_Gender
End Property

Public Property Let Address(value As String)
    m_Address = value
End Property

Public Property Get Address() As String
    Address = m_Address
End Property

Public Property Let IDCard(value As String)
    m_IDCard = value
End Property

Public Property Get IDCard() As String
    IDCard = m_IDCard
End Property

Public Property Let City(value As String)
    m_City = value
End Property

Public Property Get City() As String
    City = m_City
End Property

Public Property Let Email(value As String)
    m_Email = value
End Property

Public Property Get Email() As String
    Email = m_Email
End Property

Public Property Let JoinDate(value As String)
    m_JoinDate = value
End Property

Public Property Get JoinDate() As String
    JoinDate = m_JoinDate
End Property

Public Property Let TotalPoints(value As String)
    m_TotalPoints = value
End Property

Public Property Get TotalPoints() As String
    TotalPoints = m_TotalPoints
End Property

Public Property Let ExpireVip(value As String)
    m_ExpireVip = value
End Property

Public Property Get ExpireVip() As String
    ExpireVip = m_ExpireVip
End Property

Public Property Let Traffic(value As Variant)
    m_Traffic = value
End Property

Public Property Get Traffic() As Variant
    Traffic = m_Traffic
End Property

Public Property Let TrafficUsed(value As Variant)
    m_TrafficUsed = value
End Property

Public Property Get TrafficUsed() As Variant
    TrafficUsed = m_TrafficUsed
End Property

Public Property Let Webspace(value As Variant)
    m_Webspace = value
End Property

Public Property Get Webspace() As Variant
    Webspace = m_Webspace
End Property

Public Property Let WebspaceSecured(value As Variant)
    m_WebspaceSecured = value
End Property

Public Property Get WebspaceSecured() As Variant
    WebspaceSecured = m_WebspaceSecured
End Property

Public Property Let WebspaceSecuredUsed(value As Variant)
    m_WebspaceSecuredUsed = value
End Property

Public Property Get WebspaceSecuredUsed() As Variant
    WebspaceSecuredUsed = m_WebspaceSecuredUsed
End Property

Public Property Let Amount(value As Variant)
    m_Amount = value
End Property

Public Property Get Amount() As Variant
    Amount = m_Amount
End Property

Public Property Let DownloadTimeAvailable(value As Long)
    m_DownloadTimeAvailable = value
End Property

Public Property Get DownloadTimeAvailable() As Long
    DownloadTimeAvailable = m_DownloadTimeAvailable
End Property

Public Property Let AccountType(value As String)
    m_AccountType = value
End Property

Public Property Get AccountType() As String
    AccountType = m_AccountType
End Property

Public Property Let WebspaceUsed(value As String)
    m_WebspaceUsed = value
End Property

Public Property Get WebspaceUsed() As String
    WebspaceUsed = m_WebspaceUsed
End Property

Public Sub ExportToTextFile(DestinationFolder As String)
    Dim objFSO As Scripting.FileSystemObject
    Dim objTextFile As Scripting.TextStream
    Set objFSO = New Scripting.FileSystemObject
    Set objTextFile = objFSO.CreateTextFile(DestinationFolder & "\" & "InformationAbout" & Id & ".txt", True, True)
    With objTextFile
        .WriteLine "Fshare Account Information: "
        .WriteLine
        .WriteLine "Id: " & Id
        .WriteLine "Level: " & Level
        .WriteLine "Account type: " & AccountType
        .WriteLine "Name: " & Name
        .WriteLine "Phone: " & Phone
        .WriteLine "Birthday: " & Birthday
        .WriteLine "Gender: " & Gender
        .WriteLine "Address: " & Address
        .WriteLine "City: " & City
        .WriteLine "ID card: " & IDCard
        .WriteLine "E-mail: " & Email
        .WriteLine "Occupation: " & Occupation
        .WriteLine "Job name: " & JobName
        .WriteLine "Country: " & Country
        .WriteLine "Language: " & CountryName
        .WriteLine "Join date: " & JoinDate
        .WriteLine "Total points: " & CStr(TotalPoints)
        .WriteLine "Expire VIP: " & ExpireVip
        .WriteLine "Traffic: " & CStr(Traffic) & " GB"
        .WriteLine "Traffic used: " & CStr(TrafficUsed) & " GB"
        .WriteLine "Webspace: " & CStr(Webspace) & " GB"
        .WriteLine "Webspace used: " & CStr(WebspaceUsed) & " GB"
        .WriteLine "Webspace secured: " & CStr(WebspaceSecured) & " GB"
        .WriteLine "Webspace secured used: " & CStr(WebspaceSecuredUsed) & " GB"
        .WriteLine "Amount: " & CStr(Amount) & " Fxu"
        .WriteLine "Download time available: " & CStr(DownloadTimeAvailable)
        .WriteLine "Telesale prepaid status: " & StatusTelesalePrepaid
        .Close
    End With
    Set objFSO = Nothing
    Set objTextFile = Nothing
End Sub

