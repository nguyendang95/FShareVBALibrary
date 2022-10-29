VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FshareTopFollowMovies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colFshareTopFollowMovies As Collection

Private Sub Class_Initialize()
    Set colFshareTopFollowMovies = New Collection
End Sub

Public Sub Add(Item As FShareTopFollowMovie)
    colFshareTopFollowMovies.Add Item
End Sub

Public Sub Remove(Index As Long)
    colFshareTopFollowMovies.Remove Index
End Sub

Public Function Count() As Long
    Count = colFshareTopFollowMovies.Count
End Function

Public Function Item(Index As Long) As FShareTopFollowMovie
    Set Item = colFshareTopFollowMovies.Item(Index)
End Function

Private Sub Class_Terminate()
    Set colFshareTopFollowMovies = Nothing
End Sub

