VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FShareFileOrFolderInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colFshareFileOrFolderInfos As Collection

Private Sub Class_Initialize()
    Set colFshareFileOrFolderInfos = New Collection
End Sub

Public Sub Add(Item As FShareFileOrFolderInfo)
    colFshareFileOrFolderInfos.Add Item
End Sub

Public Sub Remove(Index As Long)
    colFshareFileOrFolderInfos.Remove Index
End Sub

Public Function Count() As Long
    Count = colFshareFileOrFolderInfos.Count
End Function

Public Function Item(Index As Long) As FShareFileOrFolderInfo
    Set Item = colFshareFileOrFolderInfos.Item(Index)
End Function

Private Sub Class_Terminate()
    Set colFshareFileOrFolderInfos = Nothing
End Sub
