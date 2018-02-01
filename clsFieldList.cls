VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFieldList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dicFields As Scripting.Dictionary

Private Sub Class_Initialize()
    Set dicFields = New Scripting.Dictionary
    dicFields.CompareMode = TextCompare
End Sub

Public Sub Add(Fields As String)
    Dim arr
    Dim i As Long
    
    arr = Split(Fields, ",")
    For i = 0 To UBound(arr)
        arr(i) = Trim(LCase(arr(i)))
        If Not dicFields.Exists(arr(i)) Then
            dicFields.Add arr(i), Empty
        End If
    Next
End Sub

Public Function ToString() As String
    Dim sAux As String
    Dim item
    
    sAux = ""
    For Each item In dicFields
        sAux = sAux & "," & item
    Next
    If sAux <> "" Then sAux = Mid(sAux, 2)
    ToString = sAux
End Function

Public Function Keys()
    Keys = dicFields.Keys
End Function
