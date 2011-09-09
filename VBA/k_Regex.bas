Attribute VB_Name = "Regex"
Option Explicit

'#==========================================================================#
'#   Shortcut functions for working with regular expressions.               #
'#                                                                          #
'#   These require a reference to "Microsoft VBScript Regular Expressions   #
'#   5.5", which is available on any computer with Internet Explorer 5.5+   #
'#   on it. The patterns follow the same syntax as used in VBScript,        #
'#   described on the following pages:                                      #
'#                                                                          #
'#       http://www.regular-expressions.info/vbscript.html                  #
'#       http://msdn.microsoft.com/en-us/library/ms974570.aspx              #
'#==========================================================================#

Public Function RegexSearch(ByVal Pattern As String, _
                            ByVal SourceString As String, _
                            Optional ByVal IgnoreCase As Boolean = False, _
                            Optional ByVal Glbl As Boolean = True, _
                            Optional ByVal Multiline As Boolean = False _
                            ) As MatchCollection
'Perform a regular expression search.
'
'Parameters
'----------
'  Pattern:
'    The regular expression pattern to search for
'  SourceString:
'    The string to search within
'  IgnoreCase (default=True):
'    Whether to be case-insensitive
'  Glbl (default=True):
'    Whether to search the entire SourceString (otherwise, stop after
'    the first match)
'  Multiline (default=False):
'    Whether to activate multiline mode. This makes the carat (^) and
'    dollar ($) match at the beginning and end of each line, rather
'    than the beggining and end of the entire string.
'
'Returns
'-------
'  MatchCollection object. A collection of Match objects, with `Count`
'  and `Item` properties. Note that `Items` is indexed by zero. Each
'  Match object has `FirstIndex`, `Length`, `Value`, and `SubMatches`
'  attributes. See:
'  http://msdn.microsoft.com/en-us/library/ms974619.aspx#scripting12_topic4
'
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    With re
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .Global = Glbl
        .Multiline = Multiline
    End With
    
    Set RegexSearch = re.Execute(SourceString)
End Function

Public Function RegexReplace(ByVal Pattern As String, _
                             ByVal Replacement As String, _
                             ByVal SourceString As String, _
                             Optional ByVal IgnoreCase As Boolean = False, _
                             Optional ByVal Glbl As Boolean = True, _
                             Optional ByVal Multiline As Boolean = False _
                             ) As String
'Perform a regular expression replacement
'
'Parameters
'----------
'  Pattern:
'    The regular expression pattern to search for
'  Replacement:
'    The string or pattern to be used as a replacement
'  SourceString:
'    The string to search within
'  IgnoreCase (default=True):
'    Whether to be case-insensitive
'  Glbl (default=True):
'    Whether to replace all matches in `SourceString` (otherwise,
'    stop after the first replacement)
'  Multiline (default=False):
'    Whether to activate multiline mode. This makes the carat (^) and
'    dollar ($) match at the beginning and end of each line, rather
'    than the beggining and end of the entire string.
'
'Returns
'-------
'  String, with `Pattern` replaced with `Replacement`, if found
'  in `SourceString`. If `Pattern` is not found, returns `SourceString`
'  as-is.
'
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    With re
        .Pattern = Pattern
        .IgnoreCase = IgnoreCase
        .Global = Glbl
        .Multiline = Multiline
    End With
    
    RegexReplace = re.Replace(SourceString, Replacement)
End Function


Private Sub RegexDemo()
'Demonstration usage
    Dim Matches As MatchCollection
    Dim M As Match
    Dim foo As String
    Dim bar As String
    
    foo = "123 abc 456"
    
    Set Matches = Regex.RegexSearch("\d+", foo)
    Debug.Print Matches.Count '--> 2
    For Each M In Matches
        Debug.Print M.FirstIndex
        Debug.Print M.Value
    Next M
    
    Set Matches = Regex.RegexSearch("blahblah", foo)
    Debug.Print Matches.Count '--> 0
    For Each M In Matches
        Debug.Print "There are no matches so you'll never see this"
    Next M
    
    bar = Regex.RegexReplace("\d+", "#", foo)
    Debug.Print bar '--> # abc #
    
    Set Matches = Regex.RegexSearch("\d+", foo, Glbl:=False)
    Debug.Print Matches.Count
    
End Sub
