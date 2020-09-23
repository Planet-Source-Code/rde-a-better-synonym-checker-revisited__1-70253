VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMsSpell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MsWord As Object

Private Sub Class_Initialize()
    Set MsWord = CreateObject("Word.Basic")
    MsWord.AppMinimize
    MsWord.FileNewDefault
End Sub

Private Sub Class_Terminate()
    MsWord.FileExit 2&
    Set MsWord = Nothing
End Sub

' Displays the Spell Check dialog box
Function SpellCheck(sText As String) As String
    On Error GoTo ErrHandler
    SpellCheck = sText
    With MsWord
        .EditSelectAll
        .EditClear
        .Insert sText
        .StartOfDocument
        .ToolsSpelling ' Checks spelling in the current selection or, if there isn't a selection, checks spelling from the location of the insertion point to the end of the document
        .EditSelectAll
        SpellCheck = TrimCR(.Selection$)
    End With
ErrHandler:
End Function

' Displays the Thesaurus dialog box
Function SynonymCheck(sWord As String) As String
    On Error GoTo ErrHandler
    SynonymCheck = sWord
    With MsWord
        .EditSelectAll
        .EditClear
        .Insert sWord
        .StartOfDocument
        .ToolsThesaurus ' Displays the Thesaurus dialog box which lists alternatives (synonyms) for the selected word
        .EditSelectAll
        SynonymCheck = TrimCR(.Selection$)
    End With
ErrHandler:
End Function

' Displays the Grammar dialog box
Function GrammarCheck(sText As String) As String
    On Error GoTo ErrHandler
    GrammarCheck = sText
    With MsWord
        .EditSelectAll
        .EditClear
        .Insert sText
        .StartOfDocument
        .ToolsGrammar ' Displays the Grammar dialog box and begins checking grammar in the active document
        .EditSelectAll
        GrammarCheck = TrimCR(.Selection$)
    End With
ErrHandler:
End Function

' Does NOT display a dialog box
Function SpellCheckResults(sWord As String, asResults() As String, Optional ByVal cResults As Long = 5) As Long
    ' Returns the number of possible replacements contained in asResults, which is
    ' redimmed as a one based array, and may be less then cResults. If the word is
    ' spelled correctly, zero is returned and asResults contains cResults empty strings.
    Dim rc As Long
    On Error GoTo ErrHandler
    ReDim asResults(1& To cResults) As String
    rc = MsWord.ToolsGetSpelling(asResults, sWord)
    If rc <> 0& Then
        SpellCheckResults = rc
        If rc < cResults Then
            ReDim Preserve asResults(1& To rc) As String
        End If
    End If
ErrHandler:
End Function

' Does NOT display a dialog box
Function SynonymCheckResults(sWord As String, asResults() As String, Optional ByVal cResults As Long = 5&) As Long
    ' Returns the number of possible alternatives (synonyms) contained in asResults, which
    ' is redimmed as a one based array, and may be less then cResults. If there are no
    ' synonyms available, zero is returned and asResults contains cResults empty strings.
    Dim rc As Long
    On Error GoTo ErrHandler
    ReDim asResults(1& To cResults) As String
    rc = MsWord.ToolsGetSynonyms(asResults, sWord)
    If rc <> 0& Then
        rc = cResults
        Do While asResults(rc) = vbNullString
            rc = rc - 1&
        Loop
        SynonymCheckResults = rc
        If rc < cResults Then
            ReDim Preserve asResults(1& To rc) As String
        End If
    End If
ErrHandler:
End Function

' Support function
Private Function TrimCR(sText As String) As String
    If Right$(sText, 1&) = vbCr Then
        TrimCR = Left$(sText, Len(sText) - 1&)
    Else
        TrimCR = sText
    End If
End Function
