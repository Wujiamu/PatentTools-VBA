VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStringMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mNames As Collection
Private mItems As Collection

Private Sub Class_Initialize()
    Set mNames = New Collection
    Set mItems = New Collection
End Sub

Private Function FindIndex(ByVal key As String) As Long
    Dim i As Long
    For i = 1 To mNames.Count
        If StrComp(CStr(mNames(i)), key, vbTextCompare) = 0 Then
            FindIndex = i
            Exit Function
        End If
    Next i
End Function

Public Function Exists(ByVal key As Variant) As Boolean
    Exists = (FindIndex(CStr(key)) > 0)
End Function

Public Sub Add(ByVal key As Variant, ByVal value As Variant)
    Dim name As String
    name = CStr(key)
    If Exists(name) Then Err.Raise 457, "clsStringMap", "This key is already associated with an element of this collection"
    mNames.Add name
    mItems.Add value
End Sub

Public Property Get Item(ByVal key As Variant) As Variant
    Dim index As Long
    index = FindIndex(CStr(key))
    If index = 0 Then Err.Raise 5, "clsStringMap", "Key not found"
    Item = mItems(index)
End Property

Public Property Let Item(ByVal key As Variant, ByVal value As Variant)
    Dim index As Long
    index = FindIndex(CStr(key))
    If index = 0 Then
        Add key, value
    Else
        mItems.Remove index
        If index <= mItems.Count Then
            InsertValue index, value
        Else
            mItems.Add value
        End If
    End If
End Property

Private Sub InsertValue(ByVal index As Long, ByVal value As Variant)
    Dim i As Long
    Dim oldValues As Collection
    Set oldValues = New Collection

    For i = 1 To mItems.Count
        If i = index Then oldValues.Add value
        oldValues.Add mItems(i)
    Next i

    Set mItems = oldValues
End Sub

Public Property Get Keys() As Variant
    Dim result() As Variant
    Dim i As Long

    If mNames.Count = 0 Then
        Keys = Array()
        Exit Property
    End If

    ReDim result(0 To mNames.Count - 1)
    For i = 1 To mNames.Count
        result(i - 1) = mNames(i)
    Next i
    Keys = result
End Property

Public Property Get Count() As Long
    Count = mNames.Count
End Property
