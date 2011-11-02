VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "z_Word"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'In order to run this VBA Macro, go to Tools->References-> MS Word 11.0 Object Library
Public appWD As Word.Application
Public ActDoc As Word.Document

'///----------------------------\\\
'||| TO USE IN EXCEL (SETUP):   |||
'\\\----------------------------///
'DIM WRD AS ICFWRD
'SET WRD = NEW ICFWRD

'   /-------------------------------------------------------------------------------\
'   |   NEW FUNCTIONS TO BE PROPERLY QA/QC AND CATEGORIZED                          |
'   \-------------------------------------------------------------------------------/


Function AddTOC(UpperHeadingLvl As Integer, LowerHeadingLvl As Integer) As Boolean
        With ActDoc
                        .TablesOfContents.Add Range:=appWD.Selection.Range, _
                        RightAlignPageNumbers:=True, _
                        UseHeadingStyles:=True, _
                        UpperHeadingLevel:=UpperHeadingLvl, _
                        LowerHeadingLevel:=LowerHeadingLvl, _
                        IncludePageNumbers:=True, _
                        AddedStyles:="", _
                        UseHyperlinks:=True, _
                        HidePageNumbersInWeb:=True, _
                        UseOutlineLevels:=True
                .TablesOfContents(1).TabLeader = wdTabLeaderDots
                .TablesOfContents.Format = wdIndexIndent
        End With
        AddTOC = True
End Function

Function UpdateTOC() As Boolean
        ActDoc.TablesOfContents(1).Update
        UpdateTOC = True
End Function

        Public Sub StandardParagraphFormat()
                'Standard paragraph formating (change as needed)
                With appWD.Selection.ParagraphFormat
                        .LeftIndent = Application.InchesToPoints(0)
                        .RightIndent = Application.InchesToPoints(0)
                        .SpaceBefore = 0
                        .SpaceBeforeAuto = False
                        .SpaceAfter = 0
                        .SpaceAfterAuto = False
                        .LineSpacingRule = wdLineSpaceSingle
        '        .Alignment = wdAlignParagraphLeft
                        .WidowControl = True
        '        .KeepWithNext = False
        '        .KeepTogether = False
        '        .PageBreakBefore = False
        '        .NoLineNumber = False
        '               .Hyphenation = True
                        .FirstLineIndent = Application.InchesToPoints(0)
                        .OutlineLevel = wdOutlineLevelBodyText
                        .CharacterUnitLeftIndent = 0
                        .CharacterUnitRightIndent = 0
                        .CharacterUnitFirstLineIndent = 0
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                        .MirrorIndents = False
        '        .TextboxTightWrap = wdTightNone
                End With
        End Sub
        
Public Function SearchFoundInWordDoc(SearchTerm As String) As Boolean
        appWD.Selection.Find.ClearFormatting
        With appWD.Selection.Find
                .Text = SearchTerm
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
        End With
        appWD.Selection.Find.Execute
        If appWD.Selection.Text = SearchTerm Then
                SearchFoundInWordDoc = True
        Else
                SearchFoundInWordDoc = False
        End If
End Function

Public Function SearchReplaceInWordDoc(SearchTerm As String, ReplaceTerm As String) As Boolean
        appWD.Selection.Find.ClearFormatting
        With appWD.Selection.Find
                .Text = SearchTerm
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
        End With
        appWD.Selection.Find.Execute
        If appWD.Selection.Text = SearchTerm Then
                appWD.Selection.Text = ReplaceTerm
                SearchReplaceInWordDoc = True
        Else
                SearchReplaceInWordDoc = False
        End If
End Function

Public Sub AddBookmark(BookmarkName As String)
        ActDoc.Bookmarks.Add Name:=BookmarkName, Range:=appWD.Selection.Range
End Sub

Public Sub AddCrossReferenceToBookmark(BookmarkName As String, Optional CrossRefText As String)
        Dim wdFld As Word.Field
        Dim BMText As String
        BMText = appWD.Documents(ActDoc).Bookmarks(BookmarkName).Range.Text
        appWD.Selection.InsertCrossReference "Bookmark", wdNumberFullContext, BookmarkName, True, False, False, " "
        If CRText <> "" Then
                For Each wdFld In appWD.Selection.Range.Paragraphs(1).Range.Fields
                        If wdFld.Type = wdFieldRef Then
                                If InStr(1, wdFld.Code.Text, BookmarkName, vbTextCompare) > 0 Then
                                        wdFld.Select
                                        appWD.Selection.Find.Execute FindText:=appWD.Selection.Text, replaceWith:=CRText
                                End If
                        End If
                Next
        End If
End Sub

'/-------------------------------------------------------------------------------\
'|   WORD SETUP AND MANIPULATION                                                 |
'|-------------------------------------------------------------------------------|
'|                       |                                                       |
'|                       |                                                       |
'|                       |                                                       |
'\-------------------------------------------------------------------------------/
        Public Sub WordDocument_SelectActive()
                Set appWD = Word.Application
                Set ActDoc = appWD.ActiveDocument
        End Sub
        
        Public Sub WordDocument_Load(WordFullPath As String, Optional ReadOnlyToggle As Boolean = False)
                Set appWD = CreateObject("Word.Application")
                appWD.Application.DisplayAlerts = False
                appWD.Documents.Open FileName:=WordFullPath, ReadOnlyToggle:=ReadOnly
                Set ActDoc = appWD.ActiveDocument
        End Sub
        
        Public Sub WordDocument_Add(FileDir As String, FileNameNoExt As String, Optional WordFileExtension As String = "docx", Optional FullTemplatePath As String = "")
                '----------------------------------------------------------------
                ' AddWordDocument       - Creates a word document of the designated type or template
                '                       - In : ByVal XValue As String, XRange As Range, YRange As Range
                '                       - Out: Linear interpolation as string, may include < or > if greater than bounds of range
                '                       - Last Updated: 3/26/11 by AJS
                '----------------------------------------------------------------
                Dim FileFormat As Integer
                
                Select Case UCase(WordFileExtension)
                        'http://msdn.microsoft.com/en-us/library/bb238158(v=office.12).aspx
                        Case "DEFAULT"
                                FileFormat = 16
                        Case "DOC", ".DOC", "2003"
                                FileFormat = 0
                        Case "DOCX", ".DOCX", "2007", "2010"
                                FileFormat = 12
                        Case "DOCM", ".DOCM"
                                FileFormat = 13
                        Case Else
                                MsgBox "Unknown word file extension- " & WordFileExtension, vbCritical, "Error in Function ""AddWordDocument"""
                                Exit Sub
                End Select
                
                'CREATE NEW INSTANCE OF WORD
                Set appWD = CreateObject("Word.Application")
                appWD.Application.DisplayAlerts = False
                
                'CREATE A NEW DOCUMENT AND SAVE
                If FullTemplatePath = "" Then
                        appWD.Documents.Add DocumentType:=wdNewBlankDocument
                Else
                        appWD.Documents.Add DocumentType:=wdNewBlankDocument, Template:=FullTemplatePath
                End If
                appWD.ChangeFileOpenDirectory FileDir
                appWD.ActiveDocument.SaveAs FileName:=FileNameNoExt, _
                        FileFormat:=FileFormat, LockComments:=False, Password:="", _
                        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
                        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
                        SaveAsAOCELetter:=False
                Set ActDoc = appWD.ActiveDocument
        End Sub
        
        Public Sub SaveFile()
                appWD.Documents(ActDoc).Save
        End Sub
        
        Public Sub WordCloseFile(SaveChg As Boolean)
                appWD.Documents(ActDoc).Close SaveChanges:=SaveChg
                appWD.Quit
        End Sub
        
        Public Sub Visible(Visible As Boolean)
                appWD.Visible = Visible
        End Sub
        
        Public Sub WordOptionsManual()
                appWD.Options.BackgroundSave = False
                appWD.Options.CheckSpellingAsYouType = False
                appWD.Options.CheckGrammarAsYouType = False
                appWD.Documents(ActDoc).ShowSpellingErrors = False
        End Sub
        
        Public Sub WordOptionsAutomatic()
                appWD.Options.BackgroundSave = True
                appWD.Options.CheckSpellingAsYouType = True
                appWD.Options.CheckGrammarAsYouType = True
                appWD.Documents(ActDoc).ShowSpellingErrors = True
        End Sub
        
        Function PaperSize(Width As Double, Height As Double)
                appWD.Documents(ActDoc).PageSetup.PageWidth = appWD.InchesToPoints(Width)
                appWD.Documents(ActDoc).PageSetup.PageHeight = appWD.InchesToPoints(Height)
        End Function
        
        Public Sub PageLandscape()
                appWD.Documents(ActDoc).Sections(appWD.Selection.Information(wdActiveEndSectionNumber)).PageSetup.Orientation = wdOrientLandscape
        End Sub
        
        Public Sub PagePortrait()
                appWD.Documents(ActDoc).PageSetup.Orientation = wdOrientPortrait
        End Sub

'   /-------------------------------------------------------------------------------\
'   |   WORD NAVIGATION AND MOVEMENT                                                |
'   |-------------------------------------------------------------------------------|
'   |                       |                                                       |
'   |                       |                                                       |
'   |                       |                                                       |
'   \-------------------------------------------------------------------------------/
        Public Sub GoToHeader()
        'Goes to header from body
                If appWD.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
                        appWD.ActiveWindow.Panes(2).Close
                End If
                If appWD.ActiveWindow.ActivePane.View.Type = wdNormalView Or _
                        appWD.ActiveWindow.ActivePane.View.Type = wdOutlineView Then
                        appWD.ActiveWindow.ActivePane.View.Type = wdPrintView
                End If
                appWD.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        End Sub
        
        Public Sub GoToMainTextBody()
        'Returns to body from header
                appWD.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        End Sub
        
        Public Sub GoToEnd()
        'Goes to the end of document
                appWD.Selection.EndKey unit:=wdStory
        End Sub
        
        Public Sub GoToBeginning()
        'Returns to beginning of document
                appWD.Selection.HomeKey unit:=wdStory
        End Sub

        Public Sub GoToBookmark(BookmarkName As String)
                appWD.Selection.Goto What:=wdGoToBookmark, Name:=BookmarkName
        End Sub
        
        Public Sub DeleteAllBM()
                Dim eachBM As Bookmark
                For Each eachBM In appWD.Documents(ActDoc).Bookmarks
                        eachBM.Delete
                Next
        End Sub

'   /-------------------------------------------------------------------------------\
'   |   WORD TEXT EDITING                                                           |
'   |-------------------------------------------------------------------------------|
'   |                       |                                                       |
'   |                       |                                                       |
'   |                       |                                                       |
'   \-------------------------------------------------------------------------------/
        Public Sub PrintText(TextString As String, Optional StyleName As String = "Normal")
                AddBookmark "Temp1"
                appWD.Selection.TypeText Text:=TextString
                AddBookmark "Temp2"
                SelectTextBetweenBookmarks "Temp1", "Temp2"
                If StyleName <> "" Then appWD.Selection.Style = StyleName
                appWD.Selection.MoveRight unit:=wdCharacter, Count:=1
        End Sub
        
        Public Sub NextLine()
                'Type new line
                appWD.Selection.TypeParagraph
        End Sub
        
        Public Sub BreakSectionNextPage()
                'Insert page break
                appWD.Selection.InsertBreak Type:=wdSectionBreakNextPage
        End Sub
        
        Public Sub BreakPage()
                'Insert page break
                appWD.Selection.InsertBreak Type:=wdPageBreak
        End Sub
        
        Public Sub PageMargins(LeftMargin As Double, RightMargin As Double, Top As Double, Bottom As Double)
                ActDoc.PageSetup.LeftMargin = appWD.InchesToPoints(LeftMargin)
                ActDoc.PageSetup.RightMargin = appWD.InchesToPoints(RightMargin)
                ActDoc.PageSetup.TopMargin = appWD.InchesToPoints(Top)
                ActDoc.PageSetup.BottomMargin = appWD.InchesToPoints(Bottom)
        End Sub
        
        Public Function SelectTextBetween(SearchText As String, StartField As String, EndField As String) As String
                Dim CropLeft As String
                If InStr(1, SearchText, EndField, vbTextCompare) = 0 Then
                        SelectTextBetween = "ERROR- End field not found (" & """" & EndField & """" & " not not found in " & """" & SearchText & """" & ")"
                        MsgBox FindTextBetween
                ElseIf InStr(1, SearchText, StartField, vbTextCompare) = 0 Then
                        MsgBox FindTextBetween
                        SelectTextBetween = "ERROR- Start field not found (" & """" & StartField & """" & " not not found in " & """" & SearchText & """" & ")"
                Else
                        CropLeft = Left(SearchText, InStr(1, SearchText, EndField, vbTextCompare) - 1)
                        SelectTextBetween = Right(CropLeft, Len(CropLeft) - (InStr(1, SearchText, StartField, vbTextCompare) + Len(StartField) - 1))
                End If
        End Function
        
        Public Function SelectTextBetweenBookmarks(BM1 As String, BM2 As String)
                ActDoc.Range( _
                        ActDoc.Bookmarks(BM1).Range.Start, _
                        ActDoc.Bookmarks(BM2).Range.Start).Select
        End Function

        
'   /-------------------------------------------------------------------------------\
'   |   WORD TABLE MANIPULATION                                                     |
'   |-------------------------------------------------------------------------------|
'   |                       |                                                       |
'   |                       |                                                       |
'   |                       |                                                       |
'   \-------------------------------------------------------------------------------/
        Public Sub AddTable(NumCols As Integer, NumRows As Integer)
            ActDoc.Tables.Add Range:=appWD.Selection.Range, _
                NumRows:=NumRows, NumColumns:=NumCols, _
                DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitFixed
        End Sub
        
        Public Sub Table_PasteFromExcel(PasteRange As Range)
                PasteRange.Copy
                appWD.Selection.PasteExcelTable False, False, False
        End Sub
        
        Public Sub ClearBorders()
                On Error Resume Next
                With appWD.Selection
                        .Borders(wdBorderLeft).Visible = False
                        .Borders(wdBorderRight).Visible = False
                        .Borders(wdBorderTop).Visible = False
                        .Borders(wdBorderBottom).Visible = False
                        .Borders(wdBorderHorizontal).Visible = False
                        .Borders(wdBorderVertical).Visible = False
                End With
                On Error GoTo 0
        End Sub

'   /-------------------------------------------------------------------------------\
'   |   WORD CHARTS AND PICTURES MANIPULATION                                       |
'   |-------------------------------------------------------------------------------|
'   |                       |                                                       |
'   |                       |                                                       |
'   |                       |                                                       |
'   \-------------------------------------------------------------------------------/
Function AddPictureInline(PictureDirectory As String, Optional ScaleFactor As Double = 1, Optional DrawBorder = False)
        With appWD.Documents(ActDoc).InlineShapes.AddPicture(FileName:=PictureDirectory, LinkToFile:=False, SaveWithDocument:=True, Range:=appWD.Selection.Range)
                If DrawBorder = True Then
                        .Line.Visible = msoTrue    'draw border around object
                        .Line.Weight = 1#          'weight of line around object
                End If
'        .ConvertToShape
                .Height = .Height * ScaleFactor
                .Width = .Width * ScaleFactor
        End With
End Function

Function AddPictureShape(PictureDirectory As String, Optional ScaleFactor As Double = 1, Optional DrawBorder = False)
        With appWD.Documents(ActDoc).Shapes.AddPicture(FileName:=PictureDirectory, LinkToFile:=False, SaveWithDocument:=True, Anchor:=appWD.Selection.Range)
                If DrawBorder = True Then
                        .Line.Visible = msoTrue    'draw border around object
                        .Line.Weight = 1#          'weight of line around object
                End If
'        .ConvertToInlineShape
                .Height = .Height * ScaleFactor
                .Width = .Width * ScaleFactor
        End With
End Function

Function AddChartInline(ChartPaste As Variant) 'can be ChartObject or Chart
        Dim WRDRange As Word.Range
        ChartPaste.Chart = AutoScaling = False
        ChartPaste.ChartArea.AutoScaleFont = False
        ChartPaste.Copy
        Set WRDRange = app.Selection.Range
        With WRDPasteRange
                .inlineInlineShapes(1).LockAspectRatio = msoTrue
                If ChartWidthInches <> 0 Then
                        .Shapes(1).Width = appWD.InchesToPoints(ChartWidthInches)
                End If
                If ChartHeightInches <> 0 Then
                        .InlineShapes(1).Height = appWD.InchesToPoints(ChartHeightInches)
                End If
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
End Function

Function AddChartShape(Optional ChartWidthInches As Double, Optional ChartHeightInches As Double)
        Dim WRDRange As Word.Range
        ChartPaste.Chart = AutoScaling = False
        ChartPaste.ChartArea.AutoScaleFont = False
        ChartPaste.Copy
'       appWD.Selection.PasteSpecial DataType:=wdPasteEnhancedMetafile, Placement:=wdFloatOverText
        Set WRDRange = app.Selection.Range
        With WRDPasteRange
                .Shapes(1).LockAspectRatio = msoTrue
                If ChartWidthInches <> 0 Then
                        .Shapes(1).Width = Application.InchesToPoints(ChartWidthInches)
                        
                End If
                If ChartHeightInches <> 0 Then
                        .Shapes(1).Height = Application.InchesToPoints(ChartHeightInches)
                End If
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
End Function


Function AddShapeInline(PasteShape As Shape, Optional WidthInches As Double, Optional HeightInches As Double)
        Dim WRDRange As Word.Range
        
        If HeightInches <> 0 Then PasteShape.Height = Application.InchesToPoints(HeightInches)
        If WidthInches <> 0 Then PasteShape.Width = Application.InchesToPoints(WidthInches)
        PasteShape.Copy
        Set WRDRange = appWD.Selection.Range
        With WRDRange
                appWD.Selection.Paste
        End With
End Function


