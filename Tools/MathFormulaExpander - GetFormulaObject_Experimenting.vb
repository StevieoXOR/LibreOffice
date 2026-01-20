' This file contains GetFormulaObject() and no other functions/subroutines so as to prevent mixing of stale versions of code.

' This is a development/testing-only file, for trying to get "cursor clicked on
'   formula box, but is not inside it" to work correctly and replace the formula box's contents 


Option Explicit


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

    
    
    ' Method 1: Try to get selected object
    oSelection = oDoc.CurrentController.getSelection()

    ' DIAGNOSTIC BLOCK to see what's selected by the mouse/cursor
    If Not bQuiet Then
        Dim vbNewLine As String
        vbNewLine = Chr(10)

        Dim sInfo As String
        Dim saInfo(7) As String

        saInfo(0) = "Selection obtained: " & Not IsNull(oSelection) & vbNewLine
        
        If Not (oSelection is Nothing) Then
            ' On Error GoTo ErrorHandler
            saInfo(1) = "Implementation Name: " & oSelection.getImplementationName() & vbNewLine

            ' LibreOffic Basic does NOT support oDoc.CurrentController.getSelection().getCount() ???
            ' saInfo(2) = "Selection count: " & oSelection.getCount()
            ' If oSelection.getCount() > 0 Then
            '     oSelectedObj = oSelection.getByIndex(0)

            saInfo(2) = "Supports oSelection.TextEmbeddedObject: " & oSelection.supportsService("com.sun.star.text.TextEmbeddedObject")
            saInfo(3) = "Supports oSelection.TextRange: "          & oSelection.supportsService("com.sun.star.text.TextRange")
            saInfo(4) = "Supports oSelection.TextContent: "        & oSelection.supportsService("com.sun.star.text.TextContent")
            

            ' Check if it has TextContent property
            Dim oTC As Object
            ' oTC = Null    ' Can do this line in VBA, not in LibreOffice Basic. Causes error "Object variable not set."
            Set oTC = Nothing

            If oSelection.supportsService("com.sun.star.text.TextEmbeddedObject") _
            Or oSelection.supportsService("com.sun.star.text.TextRange") _
            Or oSelection.supportsService("com.sun.star.text.TextContent") Then
                ' If the .TextContent property didn't actually exist, ignore it
                On Error Resume Next
                oTC = oSelection.TextContent
                On Error GoTo 0 ' Does not jump anywhere. Disables the current error handler and restores default error handling.
            End If
            saInfo(5) = sInfo & "Has TextContent property: " &  Not (oTC Is Nothing)
            
            ' If Not IsNull(oTC) Then   ' IsNull(x) is for Variants, not Objects
            If Not (oTC Is Nothing)  Then
                saInfo(6) = "TextContent is EmbeddedObject: " & oTC.supportsService("com.sun.star.text.TextEmbeddedObject")
            End If
        Else
            MsgBox "oSelection is Null", MB_ICONINFORMATION, "Selection Diagnostic Info"
        End If
        sInfo = Join(saInfo, vbNewLine)
        MsgBox sInfo, MB_ICONINFORMATION, "Selection Diagnostic Info"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

    ' Get Selection->SelectedObject->TextEmbeddedObject->Formulaobject
    ' Multiple objects are selected?
    If Not (oSelection is Nothing) Then
        Dim n As Long
        n = -1

        On Error Resume Next
        n = oSelection.getCount()
        If Err = 0 Then ' Err is LO Basic keyword
            ' oSelection is index-accessible, so .getCount() and .getByIndex() are actual methods
            If n > 0 Then
                oSelectedObj = oSelection.getByIndex(0)
                MsgBox "Reached Writer->Selection->SelectedObject->(Text)EmbeddedObject(Formula)", MB_ICONINFORMATION
                
                If Not (oSelectedObj Is Nothing) Then
                    If oSelectedObj.supportsService("com.sun.star.text.TextEmbeddedObject") Then
                        oEmbedded = oSelectedObj.EmbeddedObject ' Instead of the below line
                        ' oEmbedded = oSelection.EmbeddedObject
                        If Not bQuiet Then
                            MsgBox "VALID PROPERTY: Writer->Selection->SelectedObject->(Text)EmbeddedObject", MB_ICONINFORMATION
                        End If
                        
                        If Not (oEmbedded Is Nothing) Then
                            If oEmbedded.supportsService("com.sun.star.formula.FormulaProperties") Then
                                GetFormulaObject = oEmbedded    ' Return the formula object
                                If Not bQuiet Then
                                    msgContents = "YAY: Writer->Selection->SelectedObject->(Text)EmbeddedObject(Formula)"&Chr(10)
                                    msgContents = msgContents & "> You are inside 'Writer->Formula Box->Formula Editor'."
                                    MsgBox msgContents, MB_ICONINFORMATION, "Inside FormulaEditor :)"
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Else
            MsgBox "Case where Multiple Formula Boxes Selected:  oSelection.getCount() is an invalid method", MB_ICONINFORMATION
        End If
        On Error GoTo 0 ' Clear the error flag
        ' Err = 0 ' Reset the Sticky Error Flag to NoError. Does the same thing as the above line, but above must be used directly after On Error [...].
    End If


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

    

    ' Method 2: Try view cursor position

    ' Try nearby shapes in paragraph
    Dim oShapes As Object, oCursor As Object
    Dim i As Long
    oCursor = oDoc.CurrentController.getViewCursor()
    ' Get the cursor position, the formula boxes position, and see if they're relatively physically close in coordinates

    ' Get ALL text embedded objects (formula boxes)
    oShapes = oDoc.getText().createEnumeration()  ' returns XEnumeration
    Do While oShapes.hasMoreElements()
        Dim oElem As Object
        oElem = oShapes.nextElement()
        
        ' Check if it is a formula box
        If oElem.supportsService("com.sun.star.text.TextEmbeddedObject") _
        or oElem.supportsService("com.sun.star.drawing.OLE2Shape") _
        or oElem.supportsService("com.sun.star.draw.OLE2Shape") Then
            oEmbedded = oElem.EmbeddedObject
            If oEmbedded.supportsService("com.sun.star.formula.FormulaProperties") Then
                ' This is a formula box
                GetFormulaObject = oEmbedded    ' VarToReturn
                If Not bQuiet Then
                    msgContents = "YAY: Writer->AllShapesPresentInDoc->(Text)EmbeddedObject(Formula)"&Chr(10)
                    msgContents = msgContents & "> You are at 'Writer->Some Formula Box Selected (not entered)'."
                    MsgBox msgContents, MB_ICONINFORMATION, "Writer->Some Formula Box Selected"
                End If
                Exit Function
            Else
                MsgBox "Case where Formula Box Nearby: oDoc.getText().createEnumeration().nextElement().EmbeddedObject.FormulaProperties  service not supported", MB_ICONINFORMATION
            End If
        Else
            MsgBox "Case where Formula Box Nearby: oDoc.getText().createEnumeration().nextElement().[text.TextEmbeddedObject, or draw.OLE2Shape, or drawing.OLE2Shape]  service not supported", MB_ICONINFORMATION
        End If
    Loop


    

    ' Nothing found
    MsgBox "Cursor is Null", MB_ICONINFORMATION, "Error"
    ' GetFormulaObject = Null
    GetFormulaObject = Nothing
    
    ErrorHandler:
        ' Only executes if an error occurred above, thanks to everything else 'Exit Function'ing first
        MsgBox "Could not access formula object: " & Error$ & Erl, MB_ICONEXCLAMATION
    Exit Function
End Function


