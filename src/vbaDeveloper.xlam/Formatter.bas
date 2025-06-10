Attribute VB_Name = "Formatter"
Option Private Module
Option Explicit

Private Const INDENT_ON_ERROR As Boolean = True
Private Const INDENT_PRECOMPILATION_SHARP As Boolean = True

Private Const END_ON_ERROR As String = "On Error GoTo 0"
Private Const BEG_ON_ERROR As String = "On Error"

Private Const PRECOMP_BEG_IF As String = "#If"
Private Const PRECOMP_BEG_END_ELSE As String = "#Else"
Private Const PRECOMP_BEG_END_ELSEIF As String = "#ElseIf"
Private Const PRECOMP_END_IF As String = "#End If"

Private Const BEG_SUB As String = "Sub"
Private Const END_SUB As String = "End Sub"
Private Const BEG_PB_SUB As String = "Public Sub"
Private Const BEG_PV_SUB As String = "Private Sub"
Private Const BEG_FR_SUB As String = "Friend Sub"
Private Const BEG_PB_ST_SUB As String = "Public Static Sub"
Private Const BEG_PV_ST_SUB As String = "Private Static Sub"
Private Const BEG_FR_ST_SUB As String = "Friend Static Sub"

Private Const BEG_FUN As String = "Function"
Private Const END_FUN As String = "End Function"
Private Const BEG_PB_FUN As String = "Public Function"
Private Const BEG_PV_FUN As String = "Private Function"
Private Const BEG_FR_FUN As String = "Friend Function"
Private Const BEG_PB_ST_FUN As String = "Public Static Function"
Private Const BEG_PV_ST_FUN As String = "Private Static Function"
Private Const BEG_FR_ST_FUN As String = "Friend Static Function"

Private Const BEG_PROP As String = "Property"
Private Const END_PROP As String = "End Property"
Private Const BEG_PB_PROP As String = "Public Property"
Private Const BEG_PV_PROP As String = "Private Property"
Private Const BEG_FR_PROP As String = "Friend Property"
Private Const BEG_PB_ST_PROP As String = "Public Static Property"
Private Const BEG_PV_ST_PROP As String = "Private Static Property"
Private Const BEG_FR_ST_PROP As String = "Friend Static Property"

Private Const BEG_ENUM As String = "Enum"
Private Const END_ENUM As String = "End Enum"
Private Const BEG_PB_ENUM As String = "Public Enum"
Private Const BEG_PV_ENUM As String = "Private Enum"

Private Const BEG_IF As String = "If"
Private Const END_IF As String = "End If"
Private Const BEG_WITH As String = "With"
Private Const END_WITH As String = "End With"

Private Const BEG_SELECT As String = "Select Case"
Private Const END_SELECT As String = "End Select"

Private Const BEG_FOR As String = "For"
Private Const END_FOR As String = "Next"
Private Const BEG_WHILE As String = "While"
Private Const END_WHILE As String = "Wend"
Private Const BEG_DO As String = "Do"
Private Const END_DO As String = "Loop"

Private Const BEG_TYPE As String = "Type"
Private Const END_TYPE As String = "End Type"
Private Const BEG_PB_TYPE As String = "Public Type"
Private Const BEG_PV_TYPE As String = "Private Type"

' Single words that need to be handled separately
Private Const BEG_END_ELSE As String = "Else"
Private Const BEG_END_ELSEIF As String = "ElseIf"
Private Const BEG_END_CASE As String = "Case"

Private Const THEN_KEYWORD As String = "Then"
Private Const LINE_CONTINUATION As String = " _"

Private Const INDENT As String = "    "

Private words As Dictionary ' Keys are Strings, Value is an Integer indicating change in indentation

Private indentation(0 To 20) As Variant ' Prevent repeatedly building the same strings by looking them up in here

' 3-state data type for checking if part of code is within a string or not
Private Enum StringStatus
    InString
    MaybeInString
    NotInString
End Enum

Private Sub initialize()
    initializeWords
    initializeIndentation
End Sub

Private Sub initializeIndentation()
    Dim indentString As String
    indentString = ""
    Dim i As Integer
    For i = 0 To UBound(indentation)
        indentation(i) = indentString
        indentString = indentString & INDENT
    Next
End Sub

Private Sub initializeWords()
    Dim w As Dictionary
    Set w = New Dictionary
    
    With w
        If INDENT_PRECOMPILATION_SHARP Then
            .Add PRECOMP_BEG_IF, 1
            .Add PRECOMP_END_IF, -1
        End If
        
        .Add BEG_SUB, 1
        .Add END_SUB, -1
        .Add BEG_PB_SUB, 1
        .Add BEG_PV_SUB, 1
        .Add BEG_FR_SUB, 1
        .Add BEG_PB_ST_SUB, 1
        .Add BEG_PV_ST_SUB, 1
        .Add BEG_FR_ST_SUB, 1
        
        .Add BEG_FUN, 1
        .Add END_FUN, -1
        .Add BEG_PB_FUN, 1
        .Add BEG_PV_FUN, 1
        .Add BEG_FR_FUN, 1
        .Add BEG_PB_ST_FUN, 1
        .Add BEG_PV_ST_FUN, 1
        .Add BEG_FR_ST_FUN, 1
        
        .Add BEG_PROP, 1
        .Add END_PROP, -1
        .Add BEG_PB_PROP, 1
        .Add BEG_PV_PROP, 1
        .Add BEG_FR_PROP, 1
        .Add BEG_PB_ST_PROP, 1
        .Add BEG_PV_ST_PROP, 1
        .Add BEG_FR_ST_PROP, 1
        
        .Add BEG_ENUM, 1
        .Add END_ENUM, -1
        .Add BEG_PB_ENUM, 1
        .Add BEG_PV_ENUM, 1
        
        .Add BEG_IF, 1
        .Add END_IF, -1
        'because any following 'Case' indents to the left we jump two
        .Add BEG_SELECT, 2
        .Add END_SELECT, -2
        .Add BEG_WITH, 1
        .Add END_WITH, -1
        
        .Add BEG_FOR, 1
        .Add END_FOR, -1
        .Add BEG_DO, 1
        .Add END_DO, -1
        .Add BEG_WHILE, 1
        .Add END_WHILE, -1
        
        .Add BEG_TYPE, 1
        .Add END_TYPE, -1
        .Add BEG_PB_TYPE, 1
        .Add BEG_PV_TYPE, 1
        
        If INDENT_ON_ERROR Then
            ' Keep Order
            .Add END_ON_ERROR, -1
            .Add BEG_ON_ERROR, 1
        End If
        
    End With
    
    Set words = w
End Sub

Private Function CreateMiddleWords() As Dictionary
    
    Dim MiddleWords As Dictionary
    Set MiddleWords = New Dictionary
    
    With MiddleWords
        
        If INDENT_PRECOMPILATION_SHARP Then
            .Add PRECOMP_BEG_END_ELSE, Empty
            .Add PRECOMP_BEG_END_ELSEIF, Empty
        End If
        
        .Add BEG_END_ELSE, Empty
        .Add BEG_END_ELSEIF, Empty
        .Add BEG_END_CASE, Empty
    End With
    
    Set CreateMiddleWords = MiddleWords
End Function

Private Property Get vbaWords() As Dictionary
    If words Is Nothing Then
        initialize
    End If
    Set vbaWords = words
End Property

Private Function GetMiddleWords() As Dictionary
    Static MiddleWords As Dictionary
    
    If MiddleWords Is Nothing Then
        Set MiddleWords = CreateMiddleWords()
    End If
    
    Set GetMiddleWords = MiddleWords
End Function

Private Function IsMiddleWord(Line As String) As Boolean
    
    Dim MiddleWords As Dictionary
    Set MiddleWords = GetMiddleWords()
    
    Dim MiddleWord As Variant
    For Each MiddleWord In MiddleWords.Keys
        If lineStartsWith(MiddleWord, Line) Then
            IsMiddleWord = True
            Exit Function
        End If
    Next
    
End Function

Private Function IsKeyWord(Line As String) As Boolean
    Static KeyWords As Dictionary
    If KeyWords Is Nothing Then
        Set KeyWords = GetKeyWords()
    End If
    
    Dim KeyWord As Variant
    For Each KeyWord In KeyWords.Keys
        If lineStartsWith(KeyWord, Line) Then
            IsKeyWord = True
            Exit Function
        End If
    Next
End Function

Private Function GetKeyWords() As Dictionary
    Dim KeyWords As Dictionary
    Set KeyWords = New Dictionary
    
    Dim Key As Variant
    For Each Key In GetMiddleWords().Keys
        KeyWords.Add Key, Empty
    Next
    
    For Each Key In vbaWords().Keys
        KeyWords.Add Key, Empty
    Next
    
    Set GetKeyWords = KeyWords
End Function

Public Sub testFormatting()
    If words Is Nothing Then
        initialize
    End If
    'Debug.Print Application.VBE.ActiveCodePane.codePane.Parent.Name
    'Debug.Print Application.VBE.ActiveWindow.caption
    
    Dim projName As String, moduleName As String
    projName = "vbaDeveloper"
    moduleName = "Test"
    Dim vbaProject As VBProject
    Set vbaProject = Application.VBE.VBProjects(projName)
    Dim code As codeModule
    Set code = vbaProject.VBComponents(moduleName).codeModule
    
    'removeIndentation code
    'formatCode code
    formatProject vbaProject
End Sub

Public Sub formatProject(vbaProject As VBProject)
    
    Dim vbProjectFileName As String
    On Error Resume Next
        'this can throw if the workbook has never been saved.
        vbProjectFileName = vbaProject.fileName
    On Error GoTo 0
    
    Debug.Print "Formatting VBA project " & vbaProject.name, vbProjectFileName
    
    Dim component As Variant
    For Each component In vbaProject.VBComponents
        Dim codePane As codeModule
        Set codePane = component.codeModule
        'Debug.Print "Formatting " & component.Name
        formatCode codePane
    Next
End Sub

Public Sub formatActiveCodePane()
    formatCode Application.VBE.ActiveCodePane.codeModule
End Sub


Public Sub formatCode(codePane As codeModule)
    'On Error GoTo formatCodeError
    Dim lineCount As Integer
    lineCount = codePane.CountOfLines
    
    Dim isPrevLineContinuated As Boolean
    isPrevLineContinuated = False
    
    Dim IndentLevel As Integer
    IndentLevel = 0
    
    Dim FormatIncorrecLines As String
    FormatIncorrecLines = ""
    
    Dim lineNr As Integer
    For lineNr = 1 To lineCount
        
        Dim Line As String
        Line = Trim(codePane.lines(lineNr, 1))
        
        Dim LineWithoutComments As String
        LineWithoutComments = TrimComments(Line)
        
        Dim IsCurrentLineContinuated As Boolean
        IsCurrentLineContinuated = IsLineContinuated(LineWithoutComments)
        
        Dim levelChange As Integer
        levelChange = 0
        
        If LineWithoutComments = "" Then
            levelChange = 0
        ElseIf IsMiddleWord(LineWithoutComments) Then
            ' Case, Else, ElseIf need to jump to the left, and Next Indent
            levelChange = 1
            IndentLevel = IndentLevel - 1
        ElseIf isPrevLineContinuated And IsNamedArgument(LineWithoutComments) Then
            levelChange = 0
        ElseIf isLabel(LineWithoutComments) Then
            ' Labels don't have indentation
            levelChange = IndentLevel
            IndentLevel = 0
        ElseIf endsWithThen(LineWithoutComments) Then
            levelChange = 1
        ElseIf isOneLineIfStatemt(LineWithoutComments) Then
            ' check for one line If statemts
            levelChange = 0
        ElseIf isSingleNextToMultipleFor(LineWithoutComments) Then
            levelChange = -CountForNextClosures(LineWithoutComments)
        Else
            levelChange = indentChange(LineWithoutComments)
        End If
        
        ' Update Indentation for current line
        Dim CurrentIndentLevel As Integer
        If levelChange < 0 Then
            CurrentIndentLevel = IndentLevel + levelChange
        Else
            CurrentIndentLevel = IndentLevel
        End If
        
        If CurrentIndentLevel < 0 Then
            ' Reset CurrentIndentLevel and LevelChange
            CurrentIndentLevel = 0
            If levelChange < 0 Then levelChange = 0
            
            FormatIncorrecLines = FormatIncorrecLines & ", " & lineNr
        End If
        
        ' Update Code Line
        Line = indentation(CurrentIndentLevel) + Line
        codePane.ReplaceLine lineNr, Line
        
        ' Treate Indentantion for LineContinuation (lines ending with  " _")
        If IsCurrentLineContinuated And Not isPrevLineContinuated Then
            levelChange = levelChange + 2
        ElseIf isPrevLineContinuated And Not IsCurrentLineContinuated Then
            levelChange = levelChange - 2
        End If
        
        ' Update  Variables for next iteration
        IndentLevel = IndentLevel + levelChange
        isPrevLineContinuated = IsCurrentLineContinuated
    Next
    
    If FormatIncorrecLines <> "" Then
        Debug.Print "Format not correct in module: " & codePane & _
                "; Lines: " & FormatIncorrecLines
    ElseIf IndentLevel <> 0 Then
        Debug.Print "Format not correct in the end of module: " & codePane
    End If
    
    Exit Sub
formatCodeError:
    Debug.Print "Error while formatting " & codePane.Parent.name
    Debug.Print Err.Number & " " & Err.Description
    Debug.Print " on line " & lineNr & ": " & Line
    Debug.Print "IndentLevel: " & IndentLevel & " , levelChange: " & levelChange
End Sub

Public Sub removeIndentation(codePane As codeModule)
    Dim lineCount As Integer
    lineCount = codePane.CountOfLines
    
    Dim lineNr As Integer
    For lineNr = 1 To lineCount
        Dim Line As String
        Line = codePane.lines(lineNr, 1)
        Line = Trim(Line)
        Call codePane.ReplaceLine(lineNr, Line)
    Next
End Sub

Private Function indentChange(ByVal Line As String) As Integer
    indentChange = 0
    
    Dim w As Dictionary
    Set w = vbaWords()
    
    Dim word As Variant
    For Each word In w.Keys
        If lineStartsWith(word, Line) Then
            indentChange = vbaWords(word)
            GoTo hell
        End If
    Next
hell:
End Function

' Returns true if both strings are equal, ignoring case
Private Function isEqual(first As String, second As String) As Boolean
    isEqual = (StrComp(first, second, vbTextCompare) = 0)
End Function

' Returns True if strToCheck begins with begin, ignoring case
Private Function lineStartsWith(ByVal begin As String, ByVal strToCheck As String) As Boolean
    
    strToCheck = GetLeftToColon(strToCheck)
    
    ' Add Space on the right to check exact word.
    ' This avoids cases where variable or procedure start with Keywords, E.g. NextLevel
    AddSpaceOnTheRight begin
    AddSpaceOnTheRight strToCheck
    
    lineStartsWith = False
    Dim beginLength As Integer
    beginLength = Len(begin)
    
    If Len(strToCheck) >= beginLength Then
        Dim StringToCompare As String
        StringToCompare = Left(strToCheck, beginLength)
        lineStartsWith = isEqual(begin, StringToCompare)
    End If
End Function

Private Function GetLeftToColon(Line As String) As String
    
    If Line = "" Then Exit Function
    
    Dim ColonArray As Variant
    ColonArray = Split(Line, ":")
    
    Dim FirstWordLeftToColon As String
    GetLeftToColon = ColonArray(0)
    
End Function

Private Sub AddSpaceOnTheRight(ByRef Text As String)
    If Right(Text, 1) <> " " Then
        Text = Text & " "
    End If
End Sub
' Returns True if strToCheck ends with ending, ignoring case
Private Function lineEndsWith(ByVal ending As String, ByVal strToCheck As String) As Boolean
    
    ' Add Space on the right to check exact word.
    ' This avoids cases where variable or procedure start with Keywords, E.g. DoSomethingThen
    AddSpaceOnTheLeft ending
    AddSpaceOnTheLeft strToCheck
    
    lineEndsWith = False
    Dim Length As Integer
    Length = Len(ending)
    If Len(strToCheck) >= Length Then
        Dim StringToCompare As String
        StringToCompare = Right(strToCheck, Length)
        lineEndsWith = isEqual(ending, StringToCompare)
    End If
End Function

Private Sub AddSpaceOnTheLeft(ByRef Text As String)
    If Left(Text, 1) <> " " Then
        Text = " " & Text
    End If
End Sub

Private Function isSingleNextToMultipleFor(Line As String) As Boolean
    isSingleNextToMultipleFor = lineStartsWith(END_FOR, Line) And CountForNextClosures(Line) > 0
End Function

' Count how many For are closed by a single next, e.g. Next i,j,k
Private Function CountForNextClosures(Line As String) As Integer
    CountForNextClosures = UBound(Split(Line, ",")) + 1
End Function

Private Sub TestIsLabel()
    Debug.Print isLabel("AnyWord:")
    Debug.Print isLabel("AnyWord: ' Comment")
    
    Debug.Print Not isLabel("On Error GoTo nochange:")
    Debug.Print Not isLabel("Next:")
    Debug.Print Not isLabel("Next")
    Debug.Print Not isLabel("AnyWord")
    Debug.Print Not isLabel("AnyWord ' Comment")
    
    Debug.Print Not isLabel("Next: ' Comment")
    Debug.Print Not isLabel("""Anyword: ' Comment")
End Sub

Private Function isLabel(Line As String) As Boolean
    
    Dim ColonArray As Variant
    ColonArray = Split(Line, ":")
    
    If UBound(ColonArray) = 0 Then
        isLabel = False
        Exit Function
    End If
    
    Dim FirstWordLeftToColon As String
    FirstWordLeftToColon = ColonArray(0)
    
    'it must end with a colon: and may not contain a space.
    isLabel = Not IsKeyWord(FirstWordLeftToColon) And (InStr(FirstWordLeftToColon, " ") < 1) And (InStr(FirstWordLeftToColon, """") < 1)
    
End Function

Private Function IsNamedArgument(Line As String) As Boolean
    IsNamedArgument = InStr(Line, ":=") > 0
End Function

Private Function endsWithThen(Line As String) As Boolean
    endsWithThen = lineEndsWith(THEN_KEYWORD, Line)
End Function

Private Function isOneLineIfStatemt(Line As String) As Boolean
    ' If * then in same line
    ' If * then * Else in same line
    isOneLineIfStatemt = (lineStartsWith(BEG_IF, Line) _
            And (Not lineEndsWith(THEN_KEYWORD, Line)))
End Function


Private Function IsLineContinuated(Line As String) As Boolean
    IsLineContinuated = lineEndsWith(LINE_CONTINUATION, Line)
End Function


' Trims trailing comments (and whitespace before a comment) from a line of code
Private Function TrimComments(ByVal Line As String) As String
    Dim c               As Long
    Dim inQuotes        As StringStatus
    Dim inComment       As Boolean
    
    inQuotes = NotInString
    inComment = False
    For c = 1 To Len(Line)
        If Mid(Line, c, 1) = Chr(34) Then
            ' Found a double quote
            Select Case inQuotes
                Case NotInString:
                    inQuotes = InString
                Case InString:
                    inQuotes = MaybeInString
                Case MaybeInString:
                    inQuotes = InString
            End Select
        Else
            ' Resolve uncertain string status
            If inQuotes = MaybeInString Then
                inQuotes = NotInString
            End If
        End If
        ' Now know as much about status inside double quotes as possible, can test for comment
        If inQuotes = NotInString And Mid(Line, c, 1) = "'" Then
            inComment = True
            Exit For
        End If
    Next c
    If inComment Then
        TrimComments = Trim(Left(Line, c - 1))
    Else
        TrimComments = Line
    End If
End Function

