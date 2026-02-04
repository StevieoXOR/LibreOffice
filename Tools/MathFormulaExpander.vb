REM ** Creation Date: 16Jan2026-19Jan2026
REM ** Macro Programming Language: LibreOffice BASIC
REM ** MACRO NAME:     LibreOffice Math Formula Snippet Expander
REM ** MACRO PURPOSE:  Expands shorthand codes into full formula markup,
REM **   but only within formula objects or when a formula object is selected
REM ** Usage: Place cursor in formula, type shorthand (e.g., %mat2x2), run macro
' Developer Note: This is nearly the ***least*** efficient way to do this task.
'   This entire macro needs to be re-implemented with a DFA (Deterministic Finite Automaton)
'   , like how FLex and RegEx string replacement algos work, scanning
'   through the entire formula box 1 time instead of NumSubstitutionRules times.
' * FLex = Fast Lexical Analyzer, RegEx = Regular Expression


Option Explicit

Sub Main_ExpandFormulaShortcuts()
    REM ** Main entry point - call this macro to expand shortcuts in current formula
    
    Dim oDoc         As Object
    Dim oViewCursor  As Object
    Dim oTextContent As Object,  oEmbeddedObj As Object
    Dim sFormula     As String,  sNewFormula  As String
    Dim sRulesUsed   As String
    Dim sPrefix As String,  sMacroName As String,  sRulesUsedPg1 As String,  sRulesUsedPg2 As String
    Dim iIndex26thEntry As Integer,  iIndex51stEntry As Integer,  iNumChanged As Integer

    
    sMacroName = "Macro - Expand Formula Shortcuts"

    ' Get the current document
    oDoc = ThisComponent
    
    ' Scenario 1: We're INSIDE the formula editor (editing mode)
    ' Location: Writer document -> Formula Box -> Formula Editor (where you can type anything in the box at the bottom of the screen)
    If oDoc.supportsService("com.sun.star.formula.FormulaProperties") Then
        oEmbeddedObj = oDoc
    
    ' Scenario 2: We're OUTSIDE in Writer with formula selected, not selected, cursor near it or far away
    ElseIf oDoc.supportsService("com.sun.star.text.TextDocument") Then
        oEmbeddedObj = GetFormulaObject(oDoc, False)
        
        If (oEmbeddedObj Is Nothing) Then
            MsgBox "Please select a formula object, then place cursor inside, then run the macro again.", MB_ICONEXCLAMATION,  sMacroName & " - No Formula Object Found"
            Exit Sub
        End If
    
    Else
        MsgBox "This macro only works in Writer documents, inside Math formula editors.", MB_ICONEXCLAMATION,  sMacroName & " - Unsupported Document"
        Exit Sub
    End If
    
    ' Get current formula markup text
    sFormula = oEmbeddedObj.Formula
    
    ' Perform all substitutions (i.e., apply each shortcut replacement)
    sNewFormula = sFormula
    iNumChanged = 0
    sRulesUsed = ""
    sPrefix = ""
    sNewFormula = ReplaceAllShortcuts(sNewFormula, iNumChanged, sRulesUsed)
    
    
    ' Update formula if any changes were made
    If iNumChanged > 0 Then
        oEmbeddedObj.Formula = sNewFormula

        sPrefix = "Expanded " & iNumChanged & " formula shortcuts." & Chr(10)

        ' You can't force the case sensitivity to matter for some reason.
        ' If you force it to matter by appending an argument with value int 0, it can't find the substring.
        iIndex26thEntry = InStr(sRulesUsed, "26) {'")
        iIndex51stEntry = InStr(sRulesUsed, "51) {'")


        ' Get everything before the start of the 26th element, put it on the 1st page.
        ' Put everything afterward on the 2nd page.
        If iIndex26thEntry > 0 Then
            sRulesUsedPg1 = sPrefix & "Substitution Rules Used: " & Chr(10) & Left(sRulesUsed, iIndex26thEntry - 1)
            sRulesUsedPg2 = sPrefix & "Substitution Rules Used: " & Chr(10) &  Mid(sRulesUsed, iIndex26thEntry)
        Else    ' There are 25 elements or fewer.
            sRulesUsedPg1 = sPrefix & "Substitution Rules Used: " & Chr(10) & sRulesUsed
            sRulesUsedPg2 = ""
        End If

        If sRulesUsedPg2 = "" Then
            MsgBox sRulesUsedPg1, MB_ICONINFORMATION, sMacroName
        Else
            MsgBox sRulesUsedPg1, MB_ICONINFORMATION,  sMacroName & " Pg 1/2"
            MsgBox sRulesUsedPg2, MB_ICONINFORMATION,  sMacroName & " Pg 2/2"
        End If

        If iIndex51stEntry > 0 Then
            MsgBox "Too many (50+) rules were utilized to display effectively on 2 pages." &Chr(10)& "Using more pages to display utilized rules has not been implemented.", MB_ICONEXCLAMATION,  sMacroName
        End If

    Else
        MsgBox "No shortcuts found to expand."&Chr(10)&"> Did you forget to add a space after your term to substitute?", MB_ICONEXCLAMATION,  sMacroName & " - No Changes Made"
    End If
    
    ' MsgBox "FormulaObject.Implementation Name: " & oEmbeddedObj.getImplementationName(), MB_ICONINFORMATION
    ' It's com.sun.star.comp.Math.FormulaDocument
End Sub




Sub Main_ExpandFormulaShortcutsQuiet()
    ' Silent version that doesn't show message boxes unless significant error (useful for event triggers)
    Dim oDoc         As Object
    Dim oViewCursor  As Object
    Dim oTextContent As Object,  oEmbeddedObj As Object
    Dim sFormula     As String,  sNewFormula  As String
    Dim sRulesUsed   As String,  sMacroName   As String
    Dim iNumChanged  As Integer

    
    sMacroName = "Macro - Expand Formula Shortcuts"

    ' Get the current document
    oDoc = ThisComponent
    
    ' Scenario 1: We're INSIDE the formula editor (editing mode)
    ' Location: Writer document -> Formula Box -> Formula Editor (where you can type anything in the box at the bottom of the screen)
    If oDoc.supportsService("com.sun.star.formula.FormulaProperties") Then
        oEmbeddedObj = oDoc
    
    ' Scenario 2: We're OUTSIDE it in Writer with formula selected or not, cursor near or far away
    ElseIf oDoc.supportsService("com.sun.star.text.TextDocument") Then
        oEmbeddedObj = GetFormulaObject(oDoc, False)
        
        If (oEmbeddedObj Is Nothing) Then
            MsgBox "Please select a formula object, then place cursor inside, then run the macro again.", MB_ICONEXCLAMATION,  sMacroName & " - No Formula Object Found"
            Exit Sub
        End If
    Else
        MsgBox "This macro only works in Writer documents, inside Math formula editors.", MB_ICONEXCLAMATION,  sMacroName & " - Unsupported Document"
        Exit Sub
    End If
    
    ' Get current formula markup text
    sFormula = oEmbeddedObj.Formula
    
    ' Perform all substitutions (i.e., apply each shortcut replacement)
    sNewFormula = sFormula
    iNumChanged = 0
    sRulesUsed = ""
    sNewFormula = ReplaceAllShortcuts(sNewFormula, iNumChanged, sRulesUsed)
    
    ' Update formula if any changes were made
    If iNumChanged > 0 Then
        oEmbeddedObj.Formula = sNewFormula
    End If
End Sub




Function GetFormulaObject(oDoc As Object, bQuiet As Boolean) As Object
    REM ** Helper function to get formula object when in Writer document
    REM ** Tries multiple methods: selection, view cursor position, etc.
    REM ** Returns the formula object or Null if not found
    
    Dim oSelection   As Object
    Dim oViewCursor  As Object
    Dim oTextContent As Object
    Dim oSelectedObj As Object
    Dim oEmbedded    As Object
    Dim msgContents  As String

    
    
    oSelection = oDoc.CurrentController.getSelection()

    
    ' ' Get Selection->(Text)EmbeddedObject(Formula)
    ' If Not (oSelection Is Nothing) Then
    '     MsgBox "Case where Inside FormulaEditor:  Reached Writer->Selection->(Text)EmbeddedObject(Formula)", MB_ICONINFORMATION
    '     ' Check if selection itself (not one of its objects) is a TextEmbeddedObject
    '     If oSelection.supportsService("com.sun.star.text.TextEmbeddedObject") Then
    '         ' Selection (not Selection->ObjectList) IS the TextEmbeddedObject, get its EmbeddedObject property
    '         oEmbedded = oSelection.EmbeddedObject
    '         MsgBox "Case where Inside FormulaEditor:  VALID PROPERTY: Writer->Selection->(Text)EmbeddedObject", MB_ICONINFORMATION
            
    '         If Not (oEmbedded Is Nothing) Then
    '             If oEmbedded.supportsService("com.sun.star.formula.FormulaProperties") Then
    '                 msgContents = "Case where Inside FormulaEditor:" & Chr(10)
    '                 msgContents = msgContents & "> Writer->Selection->(Text)EmbeddedObject(Formula)" & Chr(10)
    '                 msgContents = msgContents & "> You are inside 'Writer->Formula Box->Formula Editor'."
    '                 MsgBox msgContents, MB_ICONINFORMATION, "Inside FormulaEditor :)"

    '                 GetFormulaObject = oEmbedded
    '                 Exit Function
    '             Else
    '                 MsgBox "Case where Inside FormulaEditor:"&Chr(10)&"> oEmbedded.FormulaProperties service not supported (so not inside FormulaEditor...)", MB_ICONINFORMATION
    '             End If
    '         Else
    '             MsgBox "Case where Inside FormulaEditor:"&Chr(10)&"> oSelection.EmbeddedObject is Null (so not inside FormulaEditor...)", MB_ICONINFORMATION
    '         End If
    '     Else
    '         MsgBox "Case where Inside FormulaEditor:"&Chr(10)&"> oSelection.TextEmbeddedObject service not supported (so not inside FormulaEditor...)", MB_ICONINFORMATION
    '     End If
    ' Else
    '     MsgBox "Case where Inside FormulaEditor:"&Chr(10)&"> oSelection is Null", MB_ICONINFORMATION
    ' End If

    ' Equivalent to above block, but no popup messages for debugging in this block
    If Not (oSelection Is Nothing) Then
        If oSelection.supportsService("com.sun.star.text.TextEmbeddedObject") Then
            oEmbedded = oSelection.EmbeddedObject
            If Not (oEmbedded Is Nothing) Then
                If oEmbedded.supportsService("com.sun.star.formula.FormulaProperties") Then
                    GetFormulaObject = oEmbedded
                    Exit Function
                End If
            End If
        End If
    End If



    ' ' Also seems to be equivalent to the above block. I'm not sure though.
    ' ' Get Selection->(OLE2Shape)EmbeddedObject(Formula)
    ' If Not bQuiet Then
    '     If Not (oSelection is Nothing) Then
    '         MsgBox "v1) Reached Writer->Selection->(OLE2Shape)EmbeddedObject(Formula)", MB_ICONINFORMATION
    '         ' Check if selection itself (not its object) is a (OLE2Shape)EmbeddedObject
    '         If oSelection.supportsService("com.sun.star.drawing.OLE2Shape") Then
    '             ' Selection (not Selection->Object) IS the EmbeddedObject, get its EmbeddedObject property
    '             oEmbedded = oSelection.EmbeddedObject
    '                 MsgBox "VALID PROPERTY: Writer->Selection->(OLE2Shape)EmbeddedObject(Formula)", MB_ICONINFORMATION
                
    '             If Not (oEmbedded is Nothing) Then
    '                 If oEmbedded.supportsService("com.sun.star.formula.FormulaProperties") Then
    '                     GetFormulaObject = oEmbedded
    ' 
    '                     msgContents = "Case where Formula Box Selected v1):"&Chr(10)
    '                     msgContents = msgContents & "> Writer->Selection->(OLE2Shape)EmbeddedObject(Formula)"&Chr(10)
    '                     msgContents = msgContents & "> You are at 'Writer->Formula Box Selected (not entered)'."
    '                     MsgBox msgContents, MB_ICONINFORMATION, "Writer->Formula Box Selected"
    '                     Exit Function
    '                 Else
    '                     MsgBox "Case where Formula Box Selected v1):"&Chr(10)&"> oEmbedded.FormulaProperties service not supported", MB_ICONINFORMATION
    '                 End If
    '             Else
    '                 MsgBox "Case where Formula Box Selected v1):"&Chr(10)&"> oSelection.EmbeddedObject is Null", MB_ICONINFORMATION
    '             End If
    '         Else
    '             MsgBox "Case where Formula Box Selected v1):"&Chr(10)&"> oSelection.OLE2Shape service not supported", MB_ICONINFORMATION
    '         End If
    '     Else
    '         MsgBox "Case where Formula Box Selected v1):"&Chr(10)&"> oSelection is Null", MB_ICONINFORMATION
    '     End If
    ' Else
    '     MsgBox "v2) Reached Writer->Selection->(OLE2Shape)EmbeddedObject(Formula)", MB_ICONINFORMATION
    '     If oSelection.supportsService("com.sun.star.drawing.OLE2Shape") Then
    '         oEmbedded = oSelection.EmbeddedObject
    '         If oEmbedded.supportsService("com.sun.star.formula.FormulaProperties") Then
    '             GetFormulaObject = oEmbedded
    '             Exit Function
    '         End If
    '     End If
    ' End If

    ' Nothing found
End Function








Function ReplaceShortcut(ByRef sSrcText As String,  sShortcut As String,  sExpansionOfShortcut As String,  ByRef iNumChanged As Integer, ByRef sRulesUsed As String) As String
    ' Helper function to replace shortcuts and track if changes were made
    ' https://help.libreoffice.org/latest/en-US/text/sbasic/shared/replace.html?&DbPAR=BASIC&System=WIN
    Dim sResult As String
    sResult = sSrcText
    
    If InStr(sSrcText, sShortcut) > 0 Then
        sResult = Replace(sSrcText, sShortcut, sExpansionOfShortcut, 1, -1, False)

        ' Modifies the passed-in variables
        If iNumChanged = 0 Then
            sRulesUsed = "1) {'" & sShortcut & "'->" & sExpansionOfShortcut & "}"
        Else
            sRulesUsed = sRulesUsed & Chr(10) & (iNumChanged+1) & ") {'" & sShortcut & "'->" & sExpansionOfShortcut & "}"
        End If

        iNumChanged = iNumChanged + 1
    End If
    
    ReplaceShortcut = sResult
End Function



Function ReplaceAllShortcuts(ByRef sNewFormula As String, ByRef iNumChanged As Integer, ByRef sRulesUsed As String) As String
    REM ** Perform all string substitutions on contents of (i.e., text inside) formula object.

    REM ** Longest strings must occur first (i.e., above other rules that are shorter) ("longest substring/prefix matching")
    REM ** Otherwise, rule:{%lim -> %limit} would have "%limit" become "%limitit"

    REM ** All terms need to avoid cyclic dependencies.
    REM ** This is because this macro may be run on formulas that already
    REM **   contain correct symbols, hence why some of these rules require the
    REM **   space appended afterward, like "%the " instead of "%the".
    REM ** Otherwise, rule:{%the -> %theta} would cause "%theta" to become
    REM **   "%thetata" (replaces "[%the]ta" with "[%theta]ta")
    REM **   , then "%thetatata", then "%thetatatata", ...

    REM ** All terms need to avoid accidental chained dependencies
    REM **   (caused by the ordering of the rules).
    REM ** This is why some of these rules require the substring-pattern
    REM **   -breaking "\" character as an intermediate rule.
    REM **   You could also SOMETIMES just reorder the rules instead, at the
    REM **   cost of readability and rule extensibility.
    REM ** If you don't prevent the input string from being contained in the
    REM **   output string, then the set of two rules 1) "%The" -> "%THETA"
    REM **   followed by 2) "%THE" -> "%THETA" (in that exact ordering of
    REM **   rules) would have "%The" become "%THETA" (Rule 1),
    REM **   then "THETATA" ("[THETA]TA") (Rule 2).
    REM **
    REM ** If you run the macro again, then Rule 2 would get executed, eliminating
    REM **   the chained dependency based on rule-ordering issue, but still having
    REM **   the cyclic dependency issue due only to Rule 2, each macro execution creating
    REM **   the next expansion: "%THETA" -> "%THETATA" -> "%THETATATA" -> "%THETATATATA", ...


    Dim saTemp(10) As String    ' sa = String array
    Dim sTemp      As String
    Dim vbNewLine  As String
    vbNewLine = Chr(10)

    ' These rules must occur before the rules of the form %mat2, %mat3
    sNewFormula = ReplaceShortcut(sNewFormula, "%mat2x3",   "matrix{a # b # c ## d # e # f}",   iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%mat3x2",   "matrix{a # b ## c # d ## e # f}",  iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%mat2",     "left [%nmatrix{%n   a # b%n## c # d%n}%nright ]",           iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%mat3",     "left [%nmatrix{%n   a # b # c%n## d # e # f%n## g # h # i%n}%nright ]", iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%mat4",     "left [%nmatrix{%n   a # b # c # d%n## e # f # g # h%n## i # j # k # l%n## m # n # o # p}%nright ]", iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%idmat2",   "left [%nmatrix{%n   1 # 0%n## 1 # 0%n}%nright ]",           iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%idmat3",   "left [%nmatrix{%n   1 # 0 # 0%n## 0 # 1 # 0%n## 0 # 0 # 1}%nright ]", iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%idmat4",   "left [%nmatrix{%n   1 # 0 # 0 # 0%n## 0 # 1 # 0 # 0%n## 0 # 0 # 1 # 0%n## 0 # 0 # 0 # 1}%nright ]", iNumChanged,    sRulesUsed)

    ' Column vectors (cvec)
    sNewFormula = ReplaceShortcut(sNewFormula, "%cvec2",    "stack{a # b}",                     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%cvec3",    "stack{a # b # c}",                 iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%integral", "int from{a} to{b} f(x) dx",        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%sum",      "sum from{i=1} to{n} a_i",          iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%lim",      "lim from{x toward infinity} f(x)", iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%inf",      "infinity",                         iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%frac",     "{{a} over {b}}",                   iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%rt",       "sqrt{x}",                          iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%sqrt",     "sqrt{x}",                          iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%irt2",     "frac{1}{sqrt{2}}",                 iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%invrt2",   "sqrt{frac{1}{2}}",                 iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%irt3",     "frac{1}{sqrt{3}}",                 iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%invrt3",   "sqrt{frac{1}{3}}",                 iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%pw2",          "%cases2",                      iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%piecewise2",   "%cases2",                      iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%cases2",       "left {%n  stack{a, x>0 # b, x <= 0}%n right none",   iNumChanged,    sRulesUsed)

    saTemp(0) = "%% Piecewise Function 4"
    saTemp(1) = "stack{%theta`=` # ` # `}"
    saTemp(2) = "size *3.75{\lbrace}"
    saTemp(3) = "  stack{"
    saTemp(4) = "    {x,```i>0}"
    saTemp(5) = "  # {y,```i=0}"
    saTemp(6) = "  # {z,```i<0}"
    saTemp(7) = "  # {%alpha,`i notin setR}"
    saTemp(8) = "  # {size *2.5{~}}"
    saTemp(9) = "  }"
    saTemp(10) = "``````"
    ' Add %n in between each array element while concatenating the array elements into a single String
    sTemp = Join(saTemp, "%n")
    sNewFormula = ReplaceShortcut(sNewFormula, "%pw4",          sTemp,   iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%piecewise4",   sTemp,   iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%cases4",       sTemp,   iNumChanged,    sRulesUsed)


    ' FORMATTING
    sNewFormula = ReplaceShortcut(sNewFormula, "%aligneqn",   "alignl{a &= b + c ## &= d}",       iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%newline",    "newline",                          iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%bigspace",   "~~~~",                             iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%medspace",   "~~",                               iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%space ",     "~~",                               iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%sp ",        "~~",                               iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%smolspace",  "~",                                iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%tinyspace",  "`",                                iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%unary=",     "`=`",                              iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%unary<",     "`lt`",                             iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%unary>",     "`gt`",                             iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%unary<=",    "`le`",                             iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%unary>=",    "`ge`",                             iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%comment",    "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%annotation", "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%description","%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%noshow",     "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%hide",       "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%hidetext",   "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%/*",         "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%//",         "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%--",         "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%<!--",       "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%<--",        "%\comment",                        iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\comment",   "%% This is a comment. You may want phantom{IamHidden} instead.",  iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%binom",      "binom{n}{k}",                      iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%deriv",      "{{df} over {dx}}",                 iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%part ",      "{{partial f} over {partial x}}",   iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%partial",    "{{partial f} over {partial x}}",   iNumChanged,    sRulesUsed)


    ' QUANTUM
    sNewFormula = ReplaceShortcut(sNewFormula, "%ket ",       "left lline <?> right rangle",      iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%ketpsi",     "left lline %psi right rangle",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%ketPsi",     "left lline %PSI right rangle",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%ketPSI",     "left lline %PSI right rangle",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%q ",         "%\qubit1",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%qubit ",     "%\qubit1",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q1 ",        "%\qubit1",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q1exp ",     "%\qubit1",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q1ket ",     "%\qubit1",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q1dirac ",   "%\qubit1",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%qubit1 ",    "%\qubit1",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\qubit1",    "%% |q> Dirac/Ket/Expanded%n%alpha left lline 0 right rangle + %beta left lline 1 right rangle",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%q2 ",        "%\qubit2",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q2exp ",     "%\qubit2",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q2ket ",     "%\qubit2",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q2dirac ",   "%\qubit2",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%qubit2 ",    "%\qubit2",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\qubit2",    "%% |qq> Dirac/Ket/Expanded%n  %alpha left lline 00 right rangle%n+ %beta left lline 01 right rangle%n+ %gamma left lline 10 right rangle%n+ %delta left lline 11 right rangle",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%q3 ",        "%\qubit3",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q3exp ",     "%\qubit3",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q3ket ",     "%\qubit3",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q3dirac ",   "%\qubit3",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%qubit3 ",    "%\qubit3",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\qubit3",    "%% |qqq> Dirac/Ket/Expanded%n  %alpha left lline 000 right rangle%n+ %beta left lline 001 right rangle%n+ %gamma left lline 010 right rangle%n+ %delta left lline 011 right rangle%n+ w left lline 000 right rangle%n+ x left lline 001 right rangle%n+ y left lline 010 right rangle%n+ z left lline 011 right rangle",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%q1col ",     "%% |q> ColVec%nleft(  stack{%alpha # %beta}  right)",                                     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q2col ",     "%% |qq> ColVec%nleft(  stack{%alpha # %beta # %gamma # %delta}  right)",                  iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q3col ",     "%% |qqq> ColVec%nleft(  stack{%alpha # %beta # %gamma # %delta # w # x # y # z}  right)", iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q4col ",     "%% |qqqq>%n%% ColVec, NOT matrix!!!%nleft(%nstack{%n  a0 # b1 # c2 # d3   #%n  e4 # f5 # g6 # h7   #%n  i8 # j9 # k10 # l11   #%n  m12 # n13 # o14 # p15%n}%nright)", iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%qp ",        "%\qplus",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q+ ",        "%\qplus",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\qplus",     "%%|+>%nfrac{1}{sqrt{2}} left(`%n    left lline 0 right rangle  +  left lline 1 right rangle%n`right)", iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%qn ",        "%\qminus",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%qm ",        "%\qminus",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%q- ",        "%\qminus",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\qminus",    "%%|->%nfrac{1}{sqrt{2}} left(`%n    left lline 0 right rangle  -  left lline 1 right rangle%n`right)", iNumChanged,    sRulesUsed)


    ' Greek letters NEED a % prefix
    sNewFormula = ReplaceShortcut(sNewFormula, "%al ",        "%\alpha",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%alp ",       "%\alpha",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\alpha",     "%alpha",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%Al ",        "%\ALPHA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%Alp ",       "%\ALPHA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%AL ",        "%\ALPHA",    iNumChanged,    sRulesUsed) 
    sNewFormula = ReplaceShortcut(sNewFormula, "%ALP ",       "%\ALPHA",    iNumChanged,    sRulesUsed) 
    sNewFormula = ReplaceShortcut(sNewFormula, "%\ALPHA",     "%ALPHA",     iNumChanged,    sRulesUsed) 
    

    sNewFormula = ReplaceShortcut(sNewFormula, "%be ",        "%\beta",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%bet ",       "%\beta",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\beta",      "%beta",      iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%Be ",        "%\BETA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%Bet ",       "%\BETA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%BE ",        "%\BETA",     iNumChanged,    sRulesUsed) 
    sNewFormula = ReplaceShortcut(sNewFormula, "%BET ",       "%\BETA",     iNumChanged,    sRulesUsed) 
    sNewFormula = ReplaceShortcut(sNewFormula, "%\BETA",      "%BETA",      iNumChanged,    sRulesUsed) 


    sNewFormula = ReplaceShortcut(sNewFormula, "%ga ",        "%\gamma",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%gam ",       "%\gamma",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\gamma",     "%gamma",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%Ga ",        "%\GAMMA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%Gam ",       "%\GAMMA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%GA ",        "%\GAMMA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%GAM ",       "%\GAMMA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\GAMMA",     "%GAMMA",     iNumChanged,    sRulesUsed)


    sNewFormula = ReplaceShortcut(sNewFormula, "%del ",       "%delta",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%DEL ",       "%DELTA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%Del ",       "%DELTA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%tri ",       "%DELTA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%triangle1",  "%DELTA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%changein",   "%DELTA",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%si ",        "%\sigma",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%sig ",       "%\sigma",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\sigma",     "%sigma",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%Si ",        "%\SIGMA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%Sig ",       "%\SIGMA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%SI ",        "%\SIGMA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%SIG ",       "%\SIGMA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\SIGMA",     "%SIGMA",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%ome ",       "%omega",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%ohm ",       "%omega",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%Ohm ",       "%\OMEGA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%Ome ",       "%\OMEGA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%OHM ",       "%\OMEGA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%OME ",       "%\OMEGA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\OMEGA",     "%OMEGA",     iNumChanged,    sRulesUsed)


    sNewFormula = ReplaceShortcut(sNewFormula, "%omi ",       "%omicron",   iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%Omi ",       "%\OMICRON",  iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%OMI ",       "%\OMICRON",  iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\OMICRON",   "%OMICRON",   iNumChanged,    sRulesUsed)
 

    sNewFormula = ReplaceShortcut(sNewFormula, "%the ",       "%theta",     iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%The ",       "%\THETA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%THE ",       "%\THETA",    iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\THETA",     "%THETA",     iNumChanged,    sRulesUsed)


    sNewFormula = ReplaceShortcut(sNewFormula, "%ze ",        "%\zeta",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%zet ",       "%\zeta",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\zeta",      "%zeta",      iNumChanged,    sRulesUsed)

    sNewFormula = ReplaceShortcut(sNewFormula, "%Ze ",        "%\ZETA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%Zet ",       "%\ZETA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%ZE ",        "%\ZETA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%ZET ",       "%\ZETA",     iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%\ZETA",      "%ZETA",      iNumChanged,    sRulesUsed)


    ' nabla should NOT have a % prefix.
    sNewFormula = ReplaceShortcut(sNewFormula, "%nab",        "nabla",      iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%gradientof", "nabla",      iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%gradient",   "nabla",      iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%grad",       "nabla",      iNumChanged,    sRulesUsed)
    sNewFormula = ReplaceShortcut(sNewFormula, "%triangle2",  "nabla",      iNumChanged,    sRulesUsed)

    ' Add more shortcuts here as needed using the same pattern:
    ' sNewFormula = ReplaceShortcut(sNewFormula, "%yourShortcutTextInsideTheFormulaEditor", "your desired formula markup", iNumChanged,    sRulesUsed)


    ' For your own convenience, leave this "newline" substitution line as the last substitution to perform.
    ' MsgBox "Pre-%n #rulesApplied=" & iNumChanged,  MB_ICONINFORMATION
    sNewFormula = ReplaceShortcut(sNewFormula, "%n",        vbNewLine,    iNumChanged,    sRulesUsed)
    ' MsgBox "Post-%n #rulesApplied=" & iNumChanged,  MB_ICONINFORMATION

	ReplaceAllShortcuts = sNewFormula
End Function




Sub ListAvailableShortcuts()
    ' Helper macro to display all available shortcuts
    ' This list is likely not a 100% accurate reflection of the full set of substitution rules.
    
    ' As you add more substitutions, you must increase the storage size of the arrays.
    Dim sMessagePg1       As String
    Dim sMessagePg2       As String
    Dim sMessagePg3       As String
    Dim saMsgFormat(20)   As String
    Dim saMsgMatrix(8)   As String
    Dim saMsgVector(4)   As String
    Dim saMsgCases(3)    As String
    Dim saMsgAlgebra(7)  As String
    Dim saMsgCalculus(5) As String
    Dim saMsgSymbol(20)  As String
    Dim saMsgQuantum(14)  As String

    Dim vbNewLine As String
    Dim vbTab As String
    vbNewLine = Chr(10) ' Keyword: vbNewLine
    vbTab = Chr(9)      ' Keyword: vbTab
    ' Why so many arrays? Because it's:
    ' 1) Far fewer time-consuming memory allocations than appending every individual String, and
    ' 2) Is more unitizable (can easily swap entire categories in/out if needed)
    
    sMessagePg1 = "<> Available Formula Shortcuts" & vbNewLine
    sMessagePg1 = sMessagePg1 & "<> (apostrophes indicate a required space character)" & vbNewLine & vbNewLine

    saMsgFormat(0) = "Formatting Shortcuts"
    saMsgFormat(1) = "> %aligneqn -->  Aligned equation template"
    saMsgFormat(2) = "> %newline  -->  Inserts new visual line (newline) in displayed equation"
    saMsgFormat(3) = "> %n"+vbTab+"--> Inserts new line (vbNewLine) inside the formula editor"
    saMsgFormat(4) = "> %comment, %annotation, %description  -->  %% This is a comment"
    saMsgFormat(5) = "> %noshow, %hide, %hidetext            -->  %% This is a comment"
    saMsgFormat(6) = "> %--,   %<!--,   %<--,   %/*,   %//   -->  %% This is a comment"+vbNewLine
    saMsgFormat(7)  = "Get rid of red ? when using multi-line equations"
    saMsgFormat(8)  = "> %unary=    -->   `=`"
    saMsgFormat(9)  = "> %unary<    -->   `lt`"
    saMsgFormat(10)  = "> %unary>    -->   `gt`"
    saMsgFormat(11)  = "> %unary<=  -->   `le`"
    saMsgFormat(12) = "> %unary>=  -->   `ge`"+vbNewLine
    saMsgFormat(13) = "Spacing Shortcuts"
    saMsgFormat(14) = "> %bigspace"+vbTab+vbTab+"-->   ~~~~"
    saMsgFormat(15) = "> '%medspace '  -->   ~~"
    saMsgFormat(16) = "> '%space '"+vbTab+vbTab+vbTab+"-->   ~~"
    saMsgFormat(17) = "> '%sp '"+vbTab+vbTab+vbTab+vbTab+"-->   ~~"
    saMsgFormat(18) = "> %smolspace"+vbTab+"-->   ~"
    saMsgFormat(19) = "> %tinyspace"+vbTab+vbTab+"-->   `"
    saMsgFormat(20) = vbNewLine + vbNewLine

    saMsgMatrix(0) = "Matrix Shortcuts"
    saMsgMatrix(1) = "> %mat2   - 2x2 matrix"
    saMsgMatrix(2) = "> %mat3   - 3x3 matrix"
    saMsgMatrix(3) = "> %mat4   - 4x4 matrix"
    saMsgMatrix(4) = "> %mat2x3 - 2x3 matrix"
    saMsgMatrix(5) = "> %mat3x2 - 3x2 matrix"
    saMsgMatrix(6) = "> %idmat2 - 2x2 identity matrix"
    saMsgMatrix(7) = "> %idmat3 - 3x3 identity matrix"
    saMsgMatrix(8) = "> %idmat4 - 4x4 identity matrix"
    'saMsgMatrix(9) = vbNewLine

    saMsgVector(0) = "Vector Shortcuts"
    saMsgVector(1) = "> %binom - Binomial coefficient ('choose')"
    saMsgVector(2) = "> %cvec2 - 2D vector (column)"
    saMsgVector(3) = "> %cvec3 - 3D vector (column)"
    saMsgVector(4) = vbNewLine

    saMsgCases(0) = "Cases (i.e., piecewise/multi-domain functions) Shortcuts"
    saMsgCases(1) = "> %cases2, %piecewise2, %pw2 - Piecewise function w/ 2 partitions"
    saMsgCases(2) = "> %cases4, %piecewise4, %pw4 - Piecewise function  w/ 4 partitions"
    saMsgCases(3) = vbNewLine

    saMsgAlgebra(0) = "Algebra Shortcuts"
    saMsgAlgebra(1) = "> %frac"+vbTab+vbTab+"- fraction"
    saMsgAlgebra(2) = "> %sqrt, %rt - square root"
    saMsgAlgebra(3) = "> %irt2"+vbTab+"- Inverse Root2 (1 over sqrt{2})"
    saMsgAlgebra(4) = "> %invrt2"+vbTab+"- Inverse Root2 (sqrt{1/2})"
    saMsgAlgebra(5) = "> %irt3"+vbTab+"- Inverse Root3 (1 over sqrt{3})"
    saMsgAlgebra(6) = "> %invrt3"+vbTab+"- Inverse Root3 (sqrt{1/3})"
    saMsgAlgebra(7) = vbNewLine

    saMsgCalculus(0) = "Calculus Shortcuts"
    saMsgCalculus(1) = "> %integral - definite integral"
    saMsgCalculus(2) = "> %sum   - summation Σ"
    saMsgCalculus(3) = "> %lim   - limit"
    saMsgCalculus(4) = "> %deriv - derivative  df/dx"
    saMsgCalculus(5) = "> %partial, %part - partial derivative  ∂f/∂x"
    ' saMsgCalculus(6) = vbNewLine

    saMsgSymbol(0)  = "Greek Symbol Shortcuts"
    saMsgSymbol(1)  = "> '%al ', '%alp '"+vbTab+vbTab+vbTab+vbTab+vbTab+vbTab+"-> %alpha"
    saMsgSymbol(2)  = "> %Al, %Alp, '%AL ', '%ALP ' -> %ALPHA"
    saMsgSymbol(3)  = "> '%be ', '%bet '"+vbTab+vbTab+vbTab+vbTab+vbTab+vbTab+"-> %beta"
    saMsgSymbol(4)  = "> %Be, %Bet, '%BE ', '%BET ' -> %BETA"
    saMsgSymbol(5)  = "> '%ga ', '%gam '"+vbTab+vbTab+vbTab+vbTab+vbTab+vbTab+"-> %gamma"
    saMsgSymbol(6)  = "> %Ga, %Gam, '%GA ', '%GAM ' -> %GAMMA"
    saMsgSymbol(7)  = "> %del"+vbTab+vbTab+vbTab+vbTab+vbTab+"-> %delta"
    saMsgSymbol(8)  = "> %Del, '%DEL '"  +vbTab+"-> %DELTA"
    saMsgSymbol(9)  = "> '%si ', '%sig '"+vbTab+vbTab+vbTab+vbTab+vbTab+vbTab+"-> %sigma"
    saMsgSymbol(10) = "> %Si, %Sig, '%SI ', '%SIG '"+vbTab+"-> %SIGMA"
    saMsgSymbol(11) = "> %ohm, '%ome '"  +vbTab+vbTab+vbTab+vbTab+vbTab+vbTab+vbTab+"-> %omega"
    saMsgSymbol(12) = "> %Ohm, %Ome, %OHM, '%OME ' -> %OMEGA"
    saMsgSymbol(13) = "> '%omi '"+vbTab+vbTab+vbTab+vbTab+"-> %omicron"
    saMsgSymbol(14) = "> %Omi, '%OMI '"+vbTab+"-> %OMICRON"
    saMsgSymbol(15) = "> '%the '"+vbTab+vbTab+vbTab+vbTab+"-> %theta"
    saMsgSymbol(16) = "> %The, '%THE '"  +vbTab+"-> %THETA"
    saMsgSymbol(17) = "> '%ze ', '%zet '"+vbTab+vbTab+vbTab+vbTab+vbTab+vbTab+"-> %zeta"
    saMsgSymbol(18) = "> '%Ze ', '%Zet ', '%ZE ', '%ZET ' -> %ZETA"
    saMsgSymbol(19) = "> %nab, %grad, %gradient, %gradientof -> nabla"
    saMsgSymbol(20) = vbNewLine

    saMsgQuantum(0) = "Quantum Shortcuts"
    saMsgQuantum(1) = "|q〉"
    saMsgQuantum(2) = "* > '%q1 ', '%q1exp ', '%q1ket ', '%q1dirac ', '%qubit1 ', '%qubit ', '%q '"+vbNewLine+"  -->  α|0〉 + β|1〉"
    saMsgQuantum(3) = "|qq〉"
    saMsgQuantum(4) = "* > '%q2 ', '%q2exp ', '%q2ket ', '%q2dirac ', '%qubit2 '"+vbNewLine+"  -->  α|00〉 + β|01〉 + γ|10〉 + δ|11〉"
    saMsgQuantum(5) = "|qqq〉"
    saMsgQuantum(6) = "* > '%q3 ', '%q3exp ', '%q3ket ', '%q3dirac ', '%qubit3 '"+vbNewLine+"  -->  α|000〉 + β|001〉 + γ|010〉 + δ|011〉 + w|100〉 + x|101〉 + y|110〉 + z|111〉"+vbNewLine
    saMsgQuantum(7) = "> '%q1col '"+vbTab+"--> (columnVector2{α # β})"+vbTab+vbTab+vbTab +"--   |q〉"
    saMsgQuantum(8) = "> '%q2col '"+vbTab+"--> (columnVector4{α # β # γ # δ})"+vbTab+vbTab+"--   |qq〉"
    saMsgQuantum(9) = "> '%q3col '"+vbTab+"--> (columnVector8{α # β # γ # δ # w # x # y # z})   --   |qqq〉"
    saMsgQuantum(10) = "> '%q+ ', '%qp '"+vbTab+vbTab+vbTab+vbTab+vbTab+"--> 1/sqrt2 ( |0〉+|1〉 )   --   |+〉"
    saMsgQuantum(11) = "> '%q- ', '%qn ', '%qm '"+vbTab+vbTab+"--> 1/sqrt2 ( |0〉 -|1〉 )   --   | -〉"
    saMsgQuantum(12) = "> '%ket '"+vbTab+vbTab+vbTab+"-->  |?〉"
    saMsgQuantum(13) = "> %ketpsi"+vbTab+vbTab+vbTab+"-->  |ψ〉"
    saMsgQuantum(14) = "> %ketPsi, %ketPSI"+vbTab+"-->  |Ψ〉"
    ' saMsgQuantum(15) = vbNewLine

    sMessagePg1 = sMessagePg1 & Join(saMsgFormat,  vbNewLine) & Join(saMsgMatrix,   vbNewLine)

    sMessagePg2 = Join(saMsgVector,  vbNewLine) & Join(saMsgCases,    vbNewLine)
    sMessagePg2 = sMessagePg2 & Join(saMsgAlgebra, vbNewLine) & Join(saMsgCalculus, vbNewLine)

    sMessagePg3 = Join(saMsgSymbol, vbNewLine) & Join(saMsgQuantum, vbNewLine)
    
    MsgBox sMessagePg1, MB_ICONINFORMATION, "Formula Shortcuts (1/3) - Formatting, Matrix"
    MsgBox sMessagePg2, MB_ICONINFORMATION, "Formula Shortcuts (2/3) - Vector, Cases, Algebra, Calculus"
    MsgBox sMessagePg3, MB_ICONINFORMATION, "Formula Shortcuts (3/3) - Greek Symbol, Quantum"
    
End Sub
