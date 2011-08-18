Attribute VB_Name = "z_VBA"
'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|        DICTIONARY FUNCTIONS             |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************

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


Public Function Coll_ReturnUniqueCollFromColl(FullCollection As Collection) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Coll_ReturnUniqueCollFromColl     - Returns a collection of unique values from a full collection
    '                                   - In : FullCollection As Collection
    '                                   - Out: UniqueCollection of values, or error
    '                                   - Last Updated: 8/7/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    Dim UniqueCollection As New Collection
    Dim eachItem As Variant, eachUnique As Variant
    Dim MatchFound As Boolean
    On Error GoTo isError:
        For Each eachItem In FullCollection
            MatchFound = False
            For Each eachUnique In UniqueCollection
                If eachItem = eachUnique Then
                    MatchFound = True
                    Exit For
                End If
            Next
            If MatchFound = False Then UniqueCollection.Add eachItem
        Next
        Set Coll_ReturnUniqueCollFromColl = UniqueCollection
    Exit Function
isError:
    Coll_ReturnUniqueCollFromColl = CVErr(xlErrNA)
    Debug.Print "Error in Coll_ReturnUniqueCollFromColl: " & Err.Number & ": " & Err.Description
End Function

'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|        COLLECTION FUNCTIONS             |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************

Public Function Coll_ReturnListFromCollection(FullCollection As Collection) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Coll_ReturnListFromCollection     - Returns a string list from a collection
    '                                       If 3 or greater: "alpha, beta, and zeta"
    '                                       If 2: "alpha and beta"
    '                                       If 1: "alpha"
    '                                   - In : FullCollection As Collection
    '                                   - Out: List in string, or error
    '                                   - Last Updated: 8/7/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo isError
    Dim eachItem As Variant
    Dim FullString As String
    Dim FullArray() As Variant
    Dim Counter As Integer
    
    FullArray = Coll_ToArray(FullCollection)
    
    For Counter = 1 To UBound(FullArray)
        Select Case Counter
            Case 1
                FullString = FullArray(Counter)
            Case UBound(FullArray)
                If Counter = 2 Then
                    FullString = FullString & " and " & FullArray(Counter)
                Else
                    FullString = FullString & ", and " & FullArray(Counter)
                End If
            Case Else
                FullString = FullString & ", " & FullArray(Counter)
        End Select
    Next
    Coll_ReturnListFromCollection = FullString
    Exit Function
isError:
    Coll_ReturnListFromCollection = CVErr(xlErrNA)
    Debug.Print "Error in Coll_ReturnListFromCollection: " & Err.Number & ": " & Err.Description
End Function

Public Function Coll_ToArray(FullCollection As Collection) As Variant
    Dim eachItem As Variant
    Dim Counter As Integer
    Dim FullArray() As Variant
    On Error GoTo isError:
    Counter = 1
    For Each eachItem In FullCollection
        ReDim Preserve FullArray(1 To Counter)
        FullArray(Counter) = CStr(eachItem)
        Counter = Counter + 1
    Next
    Coll_ToArray = FullArray
    Exit Function
isError:
        Coll_ToArray = CVErr(xlErrNA)
        Debug.Print "Error in Coll_ToArray: " & Err.Number & ": " & Err.Description
End Function


