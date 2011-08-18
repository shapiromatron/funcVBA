Attribute VB_Name = "z_VBA"
Public Function Dict_CreateEmpty() As Variant
    '---------------------------------------------------------------------------------------------------------
    ' Dict_CreateEmpty   - Creates an empty dictionary
    '                    - In :
    '                    - Out: Dictionary as Variant
    '                    - Last Updated: 8/18/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim TempDict As Variant
    Set TempDict = CreateObject("Scripting.Dictionary")
    TempDict.RemoveAll
    Set Dict_CreateEmpty = TempDict
End Function

Public Function Dict_AddOrUpdate(ThisDict As Variant, Key As String, Value As Variant) As Variant
    '---------------------------------------------------------------------------------------------------------
    ' Dict_AddOrUpdate   - Adds a new key or updates existing value in dictionary
    '                    - In :
    '                    - Out: Dictionary as Variant
    '                    - Last Updated: 8/18/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    If ThisDict.Exists(Key) Then
        ThisDict.item(Key) = Value
    Else
        ThisDict.Add Key, Value
    End If
    Set Dict_AddOrUpdate = ThisDict
End Function
