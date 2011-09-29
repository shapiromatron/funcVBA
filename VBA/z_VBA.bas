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
    On Error GoTo IsError:
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
IsError:
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

Public Function Coll_ReturnStringListFromCollection(FullCollection As Collection) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Coll_ReturnStringListFromCollection     - Returns a string list from a collection
    '                                       If 3 or greater: "alpha, beta, and zeta"
    '                                       If 2: "alpha and beta"
    '                                       If 1: "alpha"
    '                                   - In : FullCollection As Collection
    '                                   - Out: List in string, or error
    '                                   - Last Updated: 8/7/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
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
    Coll_ReturnStringListFromCollection = FullString
    Exit Function
IsError:
    Coll_ReturnStringListFromCollection = CVErr(xlErrNA)
    Debug.Print "Error in Coll_ReturnStringListFromCollection: " & Err.Number & ": " & Err.Description
End Function

Public Function Coll_ToArray(FullCollection As Collection) As Variant
    Dim eachItem As Variant
    Dim Counter As Integer
    Dim FullArray() As Variant
    On Error GoTo IsError:
    Counter = 1
    For Each eachItem In FullCollection
        ReDim Preserve FullArray(1 To Counter)
        FullArray(Counter) = CStr(eachItem)
        Counter = Counter + 1
    Next
    Coll_ToArray = FullArray
    Exit Function
IsError:
        Coll_ToArray = CVErr(xlErrNA)
        Debug.Print "Error in Coll_ToArray: " & Err.Number & ": " & Err.Description
End Function

Public Function Coll_ExistsInCollection(SearchString As String, Coll As Collection) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    ' Coll_ExistsInCollection     - Returns TRUE if value exists in collection, FALSE if otherwise
    '                             - In : SearchString As String, Coll As Collection
    '                             - Last Updated: 9/28/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    Dim eachVal As Variant
    For Each eachVal In Coll
        If SearchString = CStr(eachVal) Then
            Coll_ExistsInCollection = True
            Exit Function
        End If
    Next
    Coll_ExistsInCollection = False
End Function

Public Function Coll_AddIfUnique(Addition As String, Coll As Collection)
    '-----------------------------------------------------------------------------------------------------------
    ' Coll_AddIfUnique     - Adds to collection if value doesn't already exist in collection
    '                      - In : Addition as string, Coll as collection
    '                      - Last Updated: 9/28/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    If Coll_ExistsInCollection(Addition, Coll) = False Then
        Coll.Add Addition
    End If
    Set Coll_AddIfUnique = Coll
End Function

Public Function Coll_Sort(ByVal c As Collection) As Collection
' This routine uses the "heap sort" algorithm to sort a VB collection.
' It returns the sorted collection.
' Author: Christian d'Heureuse (www.source-code.biz)
' http://www.source-code.biz/snippets/vbasic/6.htm
   Dim n As Long: n = c.Count
   If n = 0 Then Set SortCollection = New Collection: Exit Function
   ReDim Index(0 To n - 1) As Long                    ' allocate index array
   Dim i As Long, m As Long
   For i = 0 To n - 1: Index(i) = i + 1: Next         ' fill index array
   For i = n \ 2 - 1 To 0 Step -1                     ' generate ordered heap
      Heapify c, Index, i, n
      Next
   For m = n To 2 Step -1                             ' sort the index array
      Exchange Index, 0, m - 1                        ' move highest element to top
      Heapify c, Index, 0, m - 1
      Next
   Dim c2 As New Collection
   For i = 0 To n - 1: c2.Add c.item(Index(i)): Next  ' fill output collection
   Set Coll_Sort = c2
   End Function

Private Sub Heapify(ByVal c As Collection, Index() As Long, ByVal i1 As Long, ByVal n As Long)
   ' Heap order rule: a[i] >= a[2*i+1] and a[i] >= a[2*i+2]
   ' http://www.source-code.biz/snippets/vbasic/6.htm
   Dim nDiv2 As Long: nDiv2 = n \ 2
   Dim i As Long: i = i1
   Do While i < nDiv2
      Dim k As Long: k = 2 * i + 1
      If k + 1 < n Then
         If c.item(Index(k)) < c.item(Index(k + 1)) Then k = k + 1
         End If
      If c.item(Index(i)) >= c.item(Index(k)) Then Exit Do
      Exchange Index, i, k
      i = k
      Loop
   End Sub

Private Sub Exchange(Index() As Long, ByVal i As Long, ByVal j As Long)
   ' http://www.source-code.biz/snippets/vbasic/6.htm
   Dim Temp As Long: Temp = Index(i)
   Index(i) = Index(j)
   Index(j) = Temp
   End Sub

''*********************************************
''*/-----------------------------------------\*
''*|                                         |*
''*|  VBA OBJECT IMPORT/EXPORT FUNCTIONS     |*
''*|                                         |*
''*\-----------------------------------------/*
''*********************************************
'Public Function ExportVBComponent(VBComp As vbide.VBComponent, _
'                                  FolderName As String, _
'                                  Optional FileName As String, _
'                                  Optional ByVal Extension As String, _
'                                  Optional OverwriteExisting As Boolean = True) As Variant
'    '-----------------------------------------------------------------------------------------------------------
'    ' ExportVBComponent   - This function exports the code module of a VBComponent to a text
'    '                       file. If FileName is missing, the code will be exported to
'    '                       a file with the same name as the VBComponent followed by the
'    '                       appropriate extension.
'    '                     - Last Updated: 8/27/11 by AJS, created by GH
'    '-----------------------------------------------------------------------------------------------------------
'    Dim FName As String
'    On Error GoTo IsError
'
'    'get extension (if not passed)
'    If Trim(Extension) = vbNullString Then
'        Extension = z_VBA.GetVBAFileExtension(VBComp:=VBComp)
'    End If
'
'    'get filename and extension
'    If Trim(FileName) = vbNullString Then
'        FName = VBComp.Name & Extension
'    Else
'        FName = FileName
'        FName = FName & "." & Extension
'    End If
'
'    'get full directory for export
'    If Right(FolderName, 1) = "\" Then
'        FName = FolderName & FName
'    Else
'        FName = FolderName & "\" & FName
'    End If
'
'    'overwrite if needed
'    If Len(Dir(FName)) > 0 Then
'        If OverwriteExisting = True Then
'            Kill FName
'        Else
'            ExportVBComponent = ""
'            Exit Function
'        End If
'    End If
'
'    'export component; return filename
'    VBComp.Export FileName:=FName
'    ExportVBComponent = FName
'    Exit Function
'IsError:
'    ExportVBComponent = CVErr(xlErrNA)
'    Debug.Print "Error in ExportVBComponent: " & Err.Number & ": " & Err.Description
'End Function
'
'Public Function GetVBAFileExtension(VBComp As vbide.VBComponent) As String
'    '-----------------------------------------------------------------------------------------------------------
'    ' GetVBAFileExtension   - This returns the appropriate file extension based on the Type of
'    '                         the VBComponent.
'    '                       - Last Updated: 8/27/11 by AJS, created by GH
'    '-----------------------------------------------------------------------------------------------------------
'    On Error GoTo IsError:
'    Select Case VBComp.Type
'        Case vbext_ct_ClassModule, vbext_ct_Document
'            GetFileExtension = ".cls"
'        Case vbext_ct_MSForm
'            GetFileExtension = ".frm"
'        Case vbext_ct_StdModule
'            GetFileExtension = ".bas"
'        Case Else
'            GetFileExtension = ".bas"
'    End Select
'IsError:
'    ExportVBComponent = CVErr(xlErrNA)
'    Debug.Print "Error in ExportVBComponent: " & Err.Number & ": " & Err.Description
'End Function
