Attribute VB_Name = "DocumentTools"
Option Compare Database

Public Function Variable_Replacement(app As Word.Application, doc As Word.Document, tag As String, text As String, process As Integer, LOG_FILE As TextStream) As String
    'Replaces a tag denoted $tagname with text provided
    ' This function uses Two methods to updated
    ' The first method uses a find and replace all for small strings under a length of 255.
    ' This method finds all cases and replaces them. In the event a string is over 255 the second method
    ' is applied. This method only locates it once and applies it. This could change to a while loop,
    ' I might think about making it a while loop
    
    'Selecting the whole doccment
    doc.Range.Select
    app.Selection.Find.ClearFormatting
    app.Selection.Find.Replacement.ClearFormatting
    app.Selection.Find.Execute Replace:=wdReplaceAll
    
    'Application Error Handling
    On Error GoTo ExitFunction
    
    'locating tag
        With app.Selection.Find
            .ClearFormatting
            .Execute FindText:=tag
            If .found = True Then
            'Tag found moving to replace it
                If (Len(text) < 255) Then
                ' Checking text length for Method
                    With app.Selection.Find
                    'Text length less than 255 applying Method 1 (mass replace)
                        .text = tag
                        .Replacement.text = text
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                        app.Selection.Find.Execute Replace:=wdReplaceAll
                Else
                    With app.Selection
                    ' Text Length over 255 only replacing first instance
                        .WholeStory
                        .Find.ClearFormatting
                        .Find.Execute FindText:=tag
                        Options.ReplaceSelection = True
                        .TypeText text:=text
                    End With
                End If
            Else
                Call Log_Line("Filed to Find Tag", 1, LOG_FILE)
                Call Log_Line("Cannont find " + tag + " In the document", 3, LOG_FILE)
            End If
        End With
Done:
        Exit Function
ExitFunction:
        On Error Resume Next

        Call Log_Line("Function Failed Exiting", 3, LOG_FILE)
        Call UI_update("Function Failed Exiting", process)
        Call UI_update("ERROR: " & Err.Description, process)
        doc.Close
        app.Quit
        LOG_FILE.Close
        If (process = 0) Then
            MsgBox ("Failed to Run, re-run")
        End If
End Function
Public Function Variable_Remove(app As Word.Application, doc As Word.Document, tag As String, LOG_FILE As TextStream)
    'Removes a tag denoted %tagname entirely, including trailing LFCR
    '
    doc.Range.Select
    app.Selection.Find.ClearFormatting
    app.Selection.Find.Replacement.ClearFormatting
    With app.Selection.Find
        .ClearFormatting
        .Execute FindText:=tag
        If .found = True Then
            With app.Selection.Find
                .text = tag
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
        End With
        app.Selection.Find.Execute
        app.Selection.Delete
        app.Selection.HomeKey wdLine
        app.Selection.EndKey wdLine, wdExtend
        app.Selection.Delete
        Else
            Call Log_Line("Filed to Find Tag", 1, LOG_FILE)
            Call Log_Line("Cannont find " + tag + " In the Body", 3, LOG_FILE)
        End If
    End With
End Function

Public Function Header_Footer(app As Word.Application, doc As Word.Document, key As String, value As String, LOG_FILE As TextStream) As String
    doc.sections(1).Footers(wdHeaderFooterPrimary).Range.Select
    With app.Selection.Find
        .ClearFormatting
        .Execute FindText:=key
        If .found = True Then
            With app.Selection
                .WholeStory
                .Find.ClearFormatting
                .Find.Execute FindText:=key
                Options.ReplaceSelection = True
                .TypeText text:=value
            End With
            Call Log_Line("Updated in header: " + tagFindString + " Updating: " + tagReplaceString, 1, LOG_FILE)
        End If
    End With
    
    doc.sections(1).Headers(wdHeaderFooterPrimary).Range.Select
    With app.Selection.Find
        .ClearFormatting
        .Execute FindText:=key
        If .found = True Then
            With app.Selection
                .WholeStory
                .Find.ClearFormatting
                .Find.Execute FindText:=key
                Options.ReplaceSelection = True
                .TypeText text:=value
            End With
            Call Log_Line("Updated in header: " + tagFindString + " Updating: " + tagReplaceString, 1, LOG_FILE)
        End If
    End With

End Function
Public Function Find_Table_By_Title(doc As Word.Document, title As String, method As String) As Integer
    'Returns table number for table with given title within document, -1 if not found.
    Dim i As Integer
    Dim y As Integer
    If method = "update" Then
        'Finding first instance of the table
        Find_Table_By_Title = -1
        For i = 1 To doc.Tables.Count
            If doc.Tables(i).title = title Then
                Find_Table_By_Title = i
                Exit For
            End If
        Next
    ElseIf method = "template" Then
        ' Locating second instance of the template'
        y = 0
        Find_Table_By_Title = -1
            For i = 1 To doc.Tables.Count
                If y < 2 Then
                    If doc.Tables(i).title = title Then
                        Find_Table_By_Title = i
                        y = y + 1
                    End If
                Else
                    Exit For
                End If
        Next
    Else
        
    End If
    
    
End Function

Public Function Table_Instantiation(app As Word.Application, doc As Word.Document, rs As ADODB.Recordset, tableTitle As String, LOG_FILE As TextStream)
' This function iterates over a table in the database if the table can be found.
' Function exectution
'     1. Query is checked for rows      -> If recordset has no rows function exites
'     2. Table is located in template   -> if Table is not founc function exites
'     3. Table rows are inserted
'     4. Table Cells are populated
'     5. Function exits
'
' Inputs
'     - app (word.application)
'     - doc (word.Document)
'     - tableTitle (String)
'     - rs (ADODB.RecordSet)
'     - LOG_FILE (TextStream)
' Outputs
'     - Boolean (not yet used)

queryRows = rs.RecordCount
    Call Log_Line("Update Table: Starting Table Instantiation", 2, LOG_FILE)
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)

    'Write Revision Control Table into SSFS
    If queryRows > 0 Then
        tableNo = Find_Table_By_Title(doc, tableTitle, "update")
        If tableNo > 0 Then
        Call Log_Line("Update Table: TableFound", 2, LOG_FILE)
            If queryRows > 1 Then
                doc.Tables.Item(tableNo).Rows(2).Select
                app.Selection.InsertRowsBelow queryRows - 1
                Call Log_Line("Update Table: " & tableTitle & " - Inserted " & queryRows - 1 & " rows.", 3, LOG_FILE)
            End If

            rs.MoveFirst
            i = 1
            Do Until rs.EOF = True
                For j = 1 To rs.Fields.Count
                    doc.Tables.Item(tableNo).Cell(i + 1, j).Range.text = rs.Fields(j - 1)
                Next
                rs.MoveNext
                i = i + 1
            Loop

            Call Log_Line("Update Table: " & tableTitle & " - Populated " & queryRows & " rows.", 3, LOG_FILE)
        Else
            Call Log_Line("Update Table: " & tableTitle & " - Table not found.", 3, LOG_FILE)
        End If
        Call Log_Line("Update Table: " & tableTitle & " - Completed.", 3, LOG_FILE)
        Call Log_Line_Break(LOG_FILE, 1)
    End If

    End Function
    
    Public Function Table_Instantiation_UpDown(app As Word.Application, doc As Word.Document, siteID As String, rsUp As ADODB.Recordset, rsDown, tableTitle As String, LOG_FILE As TextStream)
' This function iterates over a table in the database if the table can be found.
' Function exectution
'     1. Query is checked for rows      -> If recordset has no rows function exites
'     2. Table is located in template   -> if Table is not founc function exites
'     3. Table rows are inserted
'     4. Table Cells are populated
'     5. Function exits
'
' Inputs
'     - app (word.application)
'     - doc (word.Document)
'     - tableTitle (String)
'     - rs (ADODB.RecordSet)
'     - LOG_FILE (TextStream)
' Outputs
'     - Boolean (not yet used)

Dim rs As New ADODB.Recordset
Dim upRowsCnt As Integer
Dim downRowsCnt As Integer

'Determining number of rows that need to be inserted
upRowsCnt = rsUp.RecordCount
downRowsCnt = rsDown.RecordCount
If (upRowsCnt > downRowsCnt) Then
    queryRows = upRowsCnt
Else
    queryRows = downRowsCnt
End If

    Call Log_Line("Update Table: Starting Table Instantiation", 2, LOG_FILE)
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)

    'Write Revision Control Table into SSFS
    If queryRows > 0 Then
        tableNo = Find_Table_By_Title(doc, tableTitle, "update")
        If tableNo > 0 Then
        Call Log_Line("Update Table: TableFound", 2, LOG_FILE)
                If queryRows > 1 Then
                    doc.Tables.Item(tableNo).Rows(2).Select
                    app.Selection.InsertRowsBelow queryRows - 1
                    Call Log_Line("Update Table: " & tableTitle & " - Inserted " & queryRows - 1 & " rows.", 3, LOG_FILE)
                End If
            Set rs = rsUp
            If (upRowsCnt > 0) Then
                rs.MoveFirst
                i = 1
                Do Until rs.EOF = True
                    For j = 1 To rs.Fields.Count - 1
                        doc.Tables.Item(tableNo).Cell(i + 1, j).Range.text = rs.Fields(j - 1)
                    Next
                    rs.MoveNext
                    i = i + 1
                Loop
                rs.MoveFirst
                i = 1
                Do Until rs.EOF = True
                    For j = 2 To rs.Fields.Count
                        doc.Tables.Item(tableNo).Cell(i + 1, j).Range.text = siteID
                    Next
                    rs.MoveNext
                    i = i + 1
                Loop
            End If
            

            
            Set rs = rsDown
            If (downRowsCnt > 0) Then
                rs.MoveFirst
                i = 1
                Do Until rs.EOF = True
                    For j = 1 To rs.Fields.Count - 1
                        doc.Tables.Item(tableNo).Cell(i + 1, j + 2).Range.text = rs.Fields(j)
                    Next
                    rs.MoveNext
                    i = i + 1
                Loop
                rs.MoveFirst
                i = 1
                Do Until rs.EOF = True
                    For j = 2 To rs.Fields.Count
                        doc.Tables.Item(tableNo).Cell(i + 1, j).Range.text = siteID
                    Next
                    rs.MoveNext
                    i = i + 1
                Loop
            End If

            Call Log_Line("Update Table: " & tableTitle & " - Populated " & queryRows & " rows.", 3, LOG_FILE)
        Else
            Call Log_Line("Update Table: " & tableTitle & " - Table not found.", 3, LOG_FILE)
        End If
        Call Log_Line("Update Table: " & tableTitle & " - Completed.", 3, LOG_FILE)
        Call Log_Line_Break(LOG_FILE, 1)
    End If

    End Function
Public Function Table_Deletion(app As Word.Application, doc As Word.Document, tableTitle As String, LOG_FILE As TextStream)
        ' This Function is used to Delete Tables which are not used
        '
        'Inputs:
        '       app
        '       doc
        '       tableTitle
        '       LOG_FILE
        'Outputs: None

        Dim newRange As Range
        Call Log_Line("Delete Table: " & tableTitle, 3, LOG_FILE)
        
        tableNo = Find_Table_By_Title(doc, tableTitle, "update")
        If tableNo > 0 Then
            'Finding table number
            doc.Tables.Item(tableNo).Select
            ' Setting Range for Deleting white space
            Range = app.Selection.Next(Unit:=wdParagraph, Count:=1)
            doc.Tables.Item(tableNo).Delete
            
            'Deleting empty paragraph below table
            If Len(Range) = 1 Then
                 'app.Selection.Next(Unit:=wdParagraph, Count:=1).Select
                 app.Selection.Delete
             End If
            
            'Deleting empty paragraph above table
            Range = app.Selection.Previous(Unit:=wdParagraph, Count:=1)
            If Len(Range) = 1 Then
                 app.Selection.Previous(Unit:=wdParagraph, Count:=1).Select
                 app.Selection.Delete
             End If
            Call Log_Line("Delete Table: Delete " & tableTitle, 3, LOG_FILE)
        Else
            Call Log_Line("Delete Table: Table Not Found " & tableTitle, 3, LOG_FILE)
        End If
        
End Function

Public Function removeTagSection(app As Word.Application, doc As Word.Document, tag As String)
    'Removes tags and all between entirely, including trailing LFCR
    'Section denoted by $$tagname:Start$$ and $$tagname:End$$
    ' New in V2
    Dim removeRange As Range, rangeStart As Integer, rangeEnd As Integer
    Dim found As Boolean
    
    found = False
    doc.Range.Select
    app.Selection.Find.ClearFormatting
    app.Selection.Find.Replacement.ClearFormatting
    With app.Selection.Find
            With app.Selection.Find
                .text = Left(tag, Len(tag) - 2) & ":Start$$"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
                app.Selection.Find.Execute
            If .found = True Then
                rangeStart = app.Selection.Range.Start
            End If
    End With
        
    'app.Selection.Range.End = app.Selection.Range.Start + 20
    
    app.Selection.Find.ClearFormatting
    app.Selection.Find.Replacement.ClearFormatting
    doc.Range.Select
    With app.Selection.Find
            With app.Selection.Find
                .text = Left(tag, Len(tag) - 2) & ":End$$"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            app.Selection.Find.Execute
            If .found = True Then
                found = True
                rangeEnd = app.Selection.Range.End
            Else
                found = False
            End If
        
    End With
    'Method 1
    If (found) Then
        app.Selection.Start = rangeStart
        app.Selection.End = rangeEnd
        app.Selection.Delete
    End If
    
    'Method 2
    'Set removeRange = ActiveDocument.Range(rangeStart, rangeEnd)
    'removeRange.Select
    'app.Selection.Delete

End Function
Public Function insertAboveTag(app As Word.Application, doc As Word.Document, tag As String, method As String) As Boolean
    'Inserts Empty Paragraph above tag denoted $$tagname$$ and positions selection point there
    'New in V2
    doc.Range.Select
    app.Selection.Find.ClearFormatting
    app.Selection.Find.Replacement.ClearFormatting
    
    With app.Selection.Find
        .ClearFormatting
        .Execute FindText:=tag
        If .found = True Then
            With app.Selection.Find
                .text = tag
                .Replacement.text = text
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = True
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            app.Selection.Find.Execute
            'app.Selection.InsertBefore (Chr(11))
            
            ' Method for selecting caption
            If method = "caption" Then
                ' Moving Cursor to the start of the word
                app.Selection.Start = app.Selection.Start - 1
                app.Selection.End = app.Selection.Start
                app.Selection.TypeParagraph
                'app.Selection.InsertBefore (Chr(11))
            ' Method for updating caption
            ElseIf method = "table" Then
                ' Currently when a table is selected nothing in input, the procceding funciton inputs a (insert above)
                ' Moving Cursor to the start of the word
                app.Selection.Start = app.Selection.Start - 1
                app.Selection.End = app.Selection.Start
            ElseIf method = "heading" Then
                'Method for updating header
                ' Currently when a heading is selected nothing in input, the procceding funciton inputs a (insert above)
                ' Moving Cursor to the start of the word
                app.Selection.Start = app.Selection.Start - 1
                app.Selection.End = app.Selection.Start
            ElseIf method = "find" Then
                'When the find method is called nothing is returned, this is used to put the cursor on the word.
            
            ElseIf method = "section" Then
                ' Section Heading
                ' Moving Cursor to the start of the word
                app.Selection.Start = app.Selection.Start - 1
                app.Selection.End = app.Selection.Start
                app.Selection.TypeParagraph
            
            ElseIf method = "" Then
                'Fault/Debugging finding
                ' Moving Cursor to the start of the word
                app.Selection.Start = app.Selection.Start - 1
                app.Selection.End = app.Selection.Start
                app.Selection.TypeParagraph
            Else
            
            End If
            insertAboveTag = True
        Else
        inserAbovTag = False
        End If
    End With

End Function

Public Function Insert_Caption(app As Word.Application, doc As Word.Document, Caption As String, locationTag As String, LOG_FILE As TextStream)
    Dim tagFound As Boolean
        Call Log_Line("Insert Caption: Caption place holder" & locationTag, 3, LOG_FILE)
        
        'Setting Styling and inserting text

        With app.Selection
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .InsertCaption Label:="Table", _
            title:=Caption, _
            Position:=wdCaptionPositionAbove, _
            ExcludeLabel:=0
        End With
        'Selection.Style = doc.Styles("Caption")
        'Selection.InsertAfter "" & vbCrLf
    
End Function
Public Function Insert_Table(app As Word.Application, doc As Word.Document, tableTitle As String, tableTemplate As String, query As String, Caption As String, scope As String, LOG_FILE As TextStream)
    
    ' This function inserts a caption and a table
    Dim rs As New ADODB.Recordset
    Dim tag As String
    Dim captionIns As String
    tag = "$$STANDARD_EQUIPMENT$$"
    
    captionIns = Caption
    
    'Inserting Line break for Section
    Call insertAboveTag(app, doc, tag, "caption")
    Call Log_Line("Inserting Caption:" & Caption, 3, LOG_FILE)
    Call Insert_Caption(app, doc, captionIns, tag, LOG_FILE)
    'Inserting Caption
    
    'Finding Table template
    'Dim templateTableName As String
    templateTableName = "$$STANDARD_EQUIPMENT$$"
    
    Call Table_Copy_Paste_rename(app, doc, tableTemplate, tableTitle, LOG_FILE)
    
    
    'Find Table and update with new data
    Call Find_Table_By_Title(doc, tableTemplate, "Update")
    Set rs = DatabaseConnection.Connection_Query(query)
    Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    
    rs.Close
    Set rs = Nothing
    query = ""
    

End Function
Public Function Table_Copy_Paste_rename(app As Word.Application, doc As Word.Document, tableTemplate As String, tableTitle As String, LOG_FILE As TextStream)
' This function copies a table template and pastes the table then updates the table
        
    Call Log_Line("Inserting Table: " + tableTitle, 3, LOG_FILE)
    tableNo = Find_Table_By_Title(doc, tableTemplate, "update")
    doc.Tables.Item(tableNo).Select
    app.Selection.Copy
    Call insertAboveTag(app, doc, "$$STANDARD_EQUIPMENT$$", "table")
    app.Selection.PasteAndFormat (wdFormatOriginalFormatting)
    ' formating after inserting
    'app.Selection.InsertBefore (Chr(11))
    
    ' Update table that was just inserted
    tableNo = Find_Table_By_Title(doc, tableTemplate, "template")
    doc.Tables.Item(tableNo).title = tableTitle


End Function



Public Function Insert_Section_Heading(app As Word.Application, doc As Word.Document, heading As String, LOG_FILE As TextStream)
    ' This function insert a setion heading
    ' Inputs
    '   - heading (String)
    ' Ouputs
    '   - None
    
    'app.Selection.Previous(Unit:=wdParagraph, Count:=1).Delete
    Call Log_Line("Inserting Heading:" & heading, 3, LOG_FILE)
    'app.Selection.TypeParagraph
    app.Selection.Style = doc.Styles("Heading 3")
    app.Selection.TypeText text:=heading


End Function
Function Delete_Section(app As Word.Application, doc As Word.Document, tag As String)
    'Removes tags and all between entirely, including trailing LFCR
    'Section denoted by $$tagname:Start$$ and $$tagname:End$$
    ' New in V2
    Dim removeRange As Range, rangeStart As Long, rangeEnd As Long
    
    doc.Range.Select
    app.Selection.Find.ClearFormatting
    app.Selection.Find.Replacement.ClearFormatting
    With app.Selection.Find
        .text = Left(tag, Len(tag) - 2) & ":START$$"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    app.Selection.Find.Execute
    rangeStart = app.Selection.Range.Start
    
    'app.Selection.Range.End = app.Selection.Range.Start + 20
    
    app.Selection.Find.ClearFormatting
    app.Selection.Find.Replacement.ClearFormatting
    doc.Range.Select
    With app.Selection.Find
        .text = Left(tag, Len(tag) - 2) & ":END$$"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    app.Selection.Find.Execute
    rangeEnd = app.Selection.Range.End
    
    'Method 1
    app.Selection.Start = rangeStart
    ' +1 at the end is used to delete the empty paragraph
    app.Selection.End = rangeEnd + 1
    app.Selection.Delete
    

End Function

Function Table_Pumps_Instantiate(app As Word.Application, doc As Word.Document, rs As ADODB.Recordset, equipment As String, tag As String, LOG_FILE As TextStream)
    ' This Seciton is used on for inserting the Table Pumps
    ' It should not be used for other sections to Insert tables and pumps
    
    Dim tableTitle As String
    Dim queryRows As Integer
    Dim Caption As String
    
    Call insertAboveTag(app, doc, tag, "Heading")
    Call Insert_Section_Heading(app, doc, equipment, LOG_FILE)
    
    ' Building Captions
        Caption = ": " + equipment + " Details"
        Call Log_Line("Inserting Caption:" & Caption, 3, LOG_FILE)
        Call insertAboveTag(app, doc, tag, "caption")
        Call Insert_Caption(app, doc, Caption, tag, LOG_FILE)
        
        'inserting table 1. locate template 2. copy template 3. paste template 4. update title
        Call Log_Line("Inserting Table: " + equipment, 3, LOG_FILE)
        tableTitle = "Template Pump Details"
        tableNo = Find_Table_By_Title(doc, tableTitle, "update")
        doc.Tables.Item(tableNo).Select
        app.Selection.Copy
        Call insertAboveTag(app, doc, tag, "table")
        app.Selection.PasteAndFormat (wdFormatOriginalFormatting)
        ' formating after inserting
        'app.Selection.InsertBefore (Chr(11))
        
        'Updating pasted tables title
        tableNo = Find_Table_By_Title(doc, tableTitle, "template")
        tableTitle = equipment + " Details"
        doc.Tables.Item(tableNo).title = tableTitle
        
        queryRows = rs.RecordCount
        If queryRows > 0 Then
            tableNo = Find_Table_By_Title(doc, tableTitle, "update")
            If tableNo > 0 Then
                Call Log_Line("Update Table: TableFound", 2, LOG_FILE)
                rs.MoveFirst
                i = 2
                X = 0
                Do Until rs.EOF = True
                    For j = 1 To rs.Fields.Count
                        doc.Tables.Item(tableNo).Cell(j + 1, i).Range.text = rs.Fields(j - 1)
                    Next
                    rs.MoveNext
                    i = i + 1
                Loop
                Call Log_Line("Update Table: " & tableTitle & " - Populated " & queryRows & " rows.", 3, LOG_FILE)
            Else
                Call Log_Line("Update Table: " & tableTitle & " - Table not found.", 3, LOG_FILE)
            End If
            Call Log_Line("Update Table: " & tableTitle & " - Completed.", 3, LOG_FILE)
            Call Log_Line_Break(LOG_FILE, 1)
        End If
        




End Function

Public Function NonStandardTables(app As Word.Application, doc As Word.Document, siteID As String, nsType As String, nsDescription As String, scope As String, LOG_FILE As TextStream)
    ' This function updates all nonstandard IO tables depending on switch (nsType) that is given.
    ' The nsType will change the query and iterate over the specified table (tableTitel)
    
    'Function Variables
    Dim rs As New ADODB.Recordset
    Dim queryIO As String
    Dim queryAlarmsEvents As String
    Dim queryParams As String
    Dim querySetpoints As String
    Dim queryCalcs As String
    Dim querySCADAPoints As String
    Dim tableTitle As String
    
    ' IO QUERY
    queryIO = "SELECT " & _
                    "[IOL].[MODULE], " & _
                    "[IOL].[CHANNEL], " & _
                    "[IOL].[DATA_TYPE], " & _
                    "[IOL].[DESCRIPTION], " & _
                    "ISNULL([ELECTRICAL_DESCRIPTION],'') " & _
                "FROM " & _
                    "[SITE_SPECIFIC_TAG_DATA] AS SS " & _
                    "JOIN [IO_LIST] AS IOL ON ([SS].[ID] = [IOL].[TAG_ID]) " & _
                "WHERE " & _
                    "[SS].[SITE_ID] LIKE '" & siteID & "' " & _
                    "AND [SS].[EquipmentType] LIKE '" & nsType & "' " & _
                    "AND ( " & _
                        "[SS].[Tag_Attribute] LIKE '[_]di%' " & _
                        "OR [SS].[Tag_Attribute] LIKE '[_]do%' OR [SS].[Tag_Attribute] LIKE '[_]ai%' " & _
                        "OR [SS].[Tag_Attribute] LIKE '[_]ao%' " & _
                    ") " & _
                "ORDER BY " & _
                    "CASE " & _
                        "WHEN [IOL].[MODULE] LIKE 'Main Board' THEN 1 " & _
                        "WHEN [IOL].[MODULE] LIKE 'Expansion Module 1' THEN 2 " & _
                        "WHEN [IOL].[MODULE] LIKE 'Expansion Module 2' THEN 3 " & _
                        "ELSE 4 " & _
                    "END, " & _
                    "[IOL].[CHANNEL] DESC "
    
    ' Alarms and Events Query
    queryAlarmsEvents = "SELECT " & _
                            "CONCAT([Asset_Name], ' ', [Asset_Description]) AS [Description], " & _
                            "CONCAT([Object_Group], [Tag_Attribute]) AS [Tag], " & _
                            "ISNULL([severity],'')" & _
                        "FROM " & _
                            "SITE_SPECIFIC_TAG_DATA " & _
                        "WHERE " & _
                            "[SITE_ID] LIKE '" & siteID & "' " & _
                            "AND [EquipmentType] LIKE '" & nsType & "' " & _
                            "AND [Tag_Attribute] LIKE '[_]ds%' " & _
                        "ORDER BY [Object_Group],[ASSET_DESCRIPTION] ASC "
    
    ' Params and Setpoints Query
    queryParams = "SELECT " & _
                            "CONCAT([Asset_Name], ' ', [Asset_Description]) AS [Description], " & _
                            "CONCAT([Object_Group], [Tag_Attribute]) AS [Tag], " & _
                            "[Tag_Data_Type], " & _
                            "ISNULL([Site_Specific], ISNULL([Default_Value], '')) AS [Value], " & _
                            "[EU] " & _
                        "FROM " & _
                            "SITE_SPECIFIC_TAG_DATA " & _
                        "WHERE " & _
                            "SITE_ID LIKE '" & siteID & "' " & _
                            "AND [EquipmentType] LIKE '" & nsType & "' " & _
                            "AND [Tag_Attribute] LIKE '[_]kr%' " & _
                        "ORDER BY [Object_Group],[ASSET_DESCRIPTION] ASC "
                                
    querySetpoints = "SELECT " & _
                            "CONCAT([Asset_Name], ' ', [Asset_Description]) AS [Description], " & _
                            "CONCAT([Object_Group], [Tag_Attribute]) AS [Tag], " & _
                            "[Tag_Data_Type], " & _
                            "ISNULL([Site_Specific], ISNULL([Default_Value], '')) AS [Value], " & _
                            "[EU] " & _
                        "FROM " & _
                            "SITE_SPECIFIC_TAG_DATA " & _
                        "WHERE " & _
                            "SITE_ID LIKE '" & siteID & "' " & _
                            "AND [EquipmentType] LIKE '" & nsType & "' " & _
                            "AND [Tag_Attribute] LIKE '[_]kr%' " & _
                            "AND( " & _
                                "Tag_Attribute LIKE '[_]ac%' " & _
                                "OR Tag_Attribute LIKE '%Def%' " & _
                                ") " & _
                        "ORDER BY [Object_Group],[ASSET_DESCRIPTION] ASC "
    
    ' Calculations Query (Not yet developed)
    queryCalcs = ""
    
    ' SCADAPoints Query
    querySCADAPoints = "SELECT " & _
                            "CONCAT([Asset_Name], ' ', [Asset_Description]) AS [Description], " & _
                            "CONCAT([Object_Group], [Tag_Attribute]) AS [Tag], " & _
                            "DNP3_Point_Number " & _
                        "FROM " & _
                            "SITE_SPECIFIC_TAG_DATA " & _
                        "WHERE " & _
                            "SITE_ID LIKE '" & siteID & "' " & _
                            "AND [EquipmentType] LIKE '" & nsType & "' " & _
                            "AND [DNP3_Point_Number] NOT LIKE '' " & _
                            "AND [DNP3_Point_Number] IS NOT NULL " & _
                        "ORDER BY [Object_Group],[ASSET_DESCRIPTION] ASC "
                            
                            
    ' Building Nonstandard Equipment/Instrument IO table
    Call Log_Line("NonStandard Table update: " & scope & "Physical IO", 1, LOG_FILE)
    tableTitle = "Non-Standard " + nsDescription + " Physical IO"
    Set rs = DatabaseConnection.Connection_Query(queryIO)
    Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    Set rs = Nothing

    ' Building Nonstandard Equipment/Instrument Alarm and Events Table
    Call Log_Line("NonStandard Table update: " & scope & "Alarm and Events", 1, LOG_FILE)
    tableTitle = "Non-Standard " + nsDescription + " Alarms and Events"
    Set rs = DatabaseConnection.Connection_Query(queryAlarmsEvents)
    Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    Set rs = Nothing


    ' Building Nonstandard Equipment/Instrument Parameters Table
    Call Log_Line("NonStandard Table update: " & scope & "Parameters", 1, LOG_FILE)
    tableTitle = "Non-Standard " + nsDescription + " Parameters (Site Specific Parameters)"
    Set rs = DatabaseConnection.Connection_Query(queryParams)
    Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    Set rs = Nothing

    ' Building Nonstandard Equipment/Instrument Setpoints table
    Call Log_Line("NonStandard Table update: " & scope & "Setpoints", 1, LOG_FILE)
    tableTitle = "Non-Standard " + nsDescription + " Setpoints"
    Set rs = DatabaseConnection.Connection_Query(querySetpoints)
    Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    Set rs = Nothing
    
    ' Building Nonstandard Equipment/Instrument Calculations table
    'Call Log_Line("NonStandard Table update: " & scope & "Calculations", 1, LOG_FILE)
    'tableTitle ="Non-Standard" + nsDescription +" Equipment Calculations and Statistics"
    'Set rs = DatabaseConnection.Connection_Query(queryCalcs)
    'Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    'Set rs = Nothing

    ' Building Nonstandard Equipment/Instrument SCADA Points table
    Call Log_Line("NonStandard Table update: " & scope & "SCADAPoints", 1, LOG_FILE)
    tableTitle = "Non-Standard " + nsDescription + " SCADA Points"
    Set rs = DatabaseConnection.Connection_Query(querySCADAPoints)
    Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    Set rs = Nothing

End Function


Public Function Insert_Image(app As Word.Application, doc As Word.Document, imageAddress As String, tag As String, scope As String, LOG_FILE As TextStream)
    ' This Function inserts and image inside the document
    ' This Image also sets the image Height and width using.
    ' It first determines the aspect ration then mutliples it by hiegh value.
    ' Functional Inputs
    '   - imageAddress (host file location)
    '   - tag (location where image lives)
    
    'Desired height
    Height = 250
    
    ' Local Variables
    Dim image As Word.InlineShape
    Dim originalHeight As Long
    Dim orginalWeight As Long
    Dim heigh As Integer
    Dim width As Double, aspect As Double
    
    If (File_Exists(imageAddress)) Then
        ' Locating postion to insert Tag
        Call insertAboveTag(app, doc, tag, "table")
        
        Call Log_Line(scope + "Inserting Image", 1, LOG_FILE)
        ' Inserting Image
        Set image = app.Selection.InlineShapes.AddPicture(FileName:=imageAddress _
            , LinkToFile:=False, SaveWithDocument:=True)
            
        ' Changing image aspect
        orginalWeight = image.width
        originalHeight = image.Height
        
        'aspect ratio
        aspect = orginalWeight / originalHeight
        Call Log_Line(scope + "Image Height: " + Str(Height), 3, LOG_FILE)
        width = aspect * Height
        image.Height = Height
        image.width = width
    Else
        Call Log_Line(scope + "Image does not exist at path : ", 3, LOG_FILE)
        Call Log_Line("---->" + imageAddress + "<-----", 3, LOG_FILE)
    End If
    

End Function

Public Function File_Exists(filePath As String)
    'Checks to see if file exists
    File_Exists = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
End Function

Public Function Asset_Pump_Instantiation(app As Word.Application, doc As Word.Document, siteID As String, assetAbbreviation As String, scope As String, process As Integer, LOG_FILE As TextStream)
    'This Function Builds all of the pump information for the asset.
    '   - Site Pump Duty setpoints (equipment area)
    '   - Pump variables
    '       - number of pumps
    '       - Number of continous pumps
    '   - Builds pump information
    
    'Once all the pumps are built the pumps tags are removed, further if no pumps are detected
    'all pump tags are removed
    
    ' Determinging if pumps are on the Asset
    
    'local variables
    Dim query As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim queryRows As Integer
    Dim numPumps As String
    Dim tag As String
    Dim concurrentPumpsRunning As String
    Dim tableTemplate As String
    Dim tableTitle As String
    Dim Caption As String
    Dim sectionHeading As String
    Dim tagRemoveString As String
    Dim equipment As String
    Dim tagRemoveSection As String
    
    ' Determining the max concurrent pumps and number of pumps
    query = "SELECT * FROM "
    
    'Updating pump params
    Call Log_Line(scope & queryRows & " pumps on the station", 3, LOG_FILE)
    
    query = "SELECT " & _
                "A.* " & _
            "FROM " & _
                "Look_Up_Table_FS_Values AS A " & _
                "JOIN Look_Up_Table_FS AS B ON A.TAG_KEY = B.ID " & _
            "WHERE " & _
                "A.SITE_ID LIKE '" & siteID & "' " & _
                "AND B.FS_Table LIKE 'Site Pump Overview' " & _
            "ORDER BY A.TAG_KEY ASC "
    Set rs = DatabaseConnection.Connection_Query(query)
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        tag = "$numPumps"
        numPumps = Nz(rs.Fields("VALUE").value, "")
        Call Variable_Replacement(app, doc, tag, numPumps, process, LOG_FILE)
        rs.MoveNext
        tag = "$concurrentPumpsRunning"
        concurrentPumpsRunning = Nz(rs.Fields("VALUE").value, "")
        Call Variable_Replacement(app, doc, tag, concurrentPumpsRunning, process, LOG_FILE)
        
    End If
    ' Determining the number of pumps on the asset
    ' This might Change to an independent field
    query = "SELECT " & _
                "DISTINCT SS.Object_Group AS [EQUIPMENT]" & _
            "FROM " & _
                "SITE_SPECIFIC_TAG_DATA AS SS " & _
                "LEFT JOIN SITE_OPTIONS AS OPT ON SS.SITE_ID = OPT.SITE_ID " & _
                "AND SS.OPTIONS = OPT.[OPTION] " & _
            "WHERE " & _
                "SS.SITE_ID LIKE '" & siteID & "' " & _
                "AND OPT.ACTIVE = 1 " & _
                "AND Object_Group LIKE 'PMP%' " & _
                "AND Object_Group NOT LIKE '%000X'" & _
            "ORDER BY " & _
                "Object_Group ASC "
    
    Set rs = DatabaseConnection.Connection_Query(query)
    
    queryRows = rs.RecordCount
    
    If (queryRows > 0) Then
        'Adding pump Station Setpoints
        ' Updating Table 'Pump Station Duty Setpoints'
        ' Inserting Pumps
        If (queryRows > 0) Then
         rs.MoveFirst
         Do Until rs.EOF = True
             equipment = rs.Fields("EQUIPMENT").value
             
             
             query = "SELECT " & _
                         "[EQUIPMENT], " & _
                         "[Make], " & _
                         "[model], " & _
                         "ISNULL([DUTY_POINT],''), " & _
                         "ISNULL([Power],''), " & _
                         "ISNULL([FLC],'') " & _
                     "FROM " & _
                         "[INFRASTRUCTURE].[DBO].[EQUIPMENT_TABLE] " & _
                     "WHERE " & _
                         "[SITE_ID] LIKE '" & siteID & "' " & _
                         "AND [OBJECT_GROUP] LIKE '" & equipment & "' "
             Set rs1 = DatabaseConnection.Connection_Query(query)
             
             If (rs1.RecordCount > 0) Then
                Call Log_Line(scope + "Building " + equipment, 3, LOG_FILE)
                sectionHeading = equipment
                tag = "$$PUMP_TABLES$$"
                'Find String -> copy table -> paste table -> Update title -> insert content
                Call Table_Pumps_Instantiate(app, doc, rs1, equipment, tag, LOG_FILE)
             End If
             
             rs1.Close
             Set rs1 = Nothing
             query = ""
             rs.MoveNext
         Loop
        Else
            Call Log_Line(scope + "Failed to add pumps - NO pumps listed in the EQUIPMENT TABLE -", 3, LOG_FILE)
        End If
            rs.Close
            Set rs = Nothing
            query = ""
            
            'Removing Pump place holder tags
            tagRemoveString = "$$PUMP_TABLES:START$$"
            Call Select_Current_Paragraph(app, doc, tagRemoveString, "word", "delete")
            'Call Variable_Remove(app, doc, tagRemoveString, LOG_FILE)
            Call Log_Line("Tag Removed : " & tagRemoveString, 3, LOG_FILE)
            
            tagRemoveString = "$$PUMP_TABLES:END$$"
            Call Select_Current_Paragraph(app, doc, tagRemoveString, "word", "delete")
            'Call Variable_Remove(app, doc, tagRemoveString, LOG_FILE)
            Call Log_Line("Tag Removed : " & tagRemoveString, 3, LOG_FILE)
            
            tagRemoveString = "$$PUMP_TABLES$$"
            Call Select_Current_Paragraph(app, doc, tagRemoveString, "word", "delete")
            'Call Variable_Remove(app, doc, tagRemoveString, LOG_FILE)
            Call Log_Line("Tag Removed : " & tagRemoveString, 3, LOG_FILE)
            
            tableTitle = "Template Pump Details"
            Call Table_Deletion(app, doc, tableTitle, LOG_FILE)
            Call Log_Line("Table Removed : " & tableTitle, 3, LOG_FILE)
                    
        
    Else
        ' Removing pump sections
        Call Log_Line(scope + "removing PUMP section - NO PUMPS FOUND -", 3, LOG_FILE)
        tagRemoveSection = "$$PUMP_TABLES$$"
        Call Delete_Section(app, doc, tagRemoveSection)
        Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
        tagRemoveString = "$$PUMP_TABLES$$"
        Call Select_Current_Paragraph(app, doc, tagRemoveString, "word", "delete")
        'Call Variable_Remove(app, doc, tagRemoveSection, LOG_FILE)
        Call Log_Line("Tag Removed : " & tagRemoveString, 3, LOG_FILE)
    
    End If
    
    ' If no pumps remove all sections

End Function
Public Function Asset_Equipment_check(app As Word.Application, doc As Word.Document, equipment As String, siteID As String, scope As String, LOG_FILE As TextStream) As Boolean
    ' This function checks to see if piece of equipment is on the site.
    ' If the equipment does then the function returns true otherwise it returns false
    
    Asset_Equipment_check = False
    Dim rs As New ADODB.Recordset
    Dim query As String
    
    'Performing check to see if pump is on site, this check is different from the equipment list
    If (InStr(equipment, "PMP")) Then
        query = "SELECT " & _
                    "A.Object_Group " & _
                "FROM " & _
                    "SITE_SPECIFIC_TAG_DATA AS A " & _
                    "LEFT JOIN ( " & _
                    "SELECT SITE_ID,[OPTION],ACTIVE FROM SITE_OPTIONS " & _
                    "WHERE SITE_ID like '" & siteID & "' AND ACTIVE = 1 " & _
                    "UNION " & _
                    "SELECT SITE_ID,[COMPOSITE_OPTION],ACTIVE FROM SITE_OPTIONS_COMPOSITE " & _
                    "WHERE SITE_ID like '" & siteID & "' AND ACTIVE = 1) " & _
                    "AS B ON [A].[OPTIONS] = [B].[OPTION] " & _
                    "JOIN EQUIPMENT_TABLE AS C ON A.SITE_ID = C.SITE_ID and A.Object_Group = C.Object_Group AND C.Object_Group like '" & equipment & "' AND C.SITE_ID like '" & siteID & "' " & _
                "WHERE " & _
                    "A.SITE_ID LIKE '" & siteID & "' " & _
                    "AND A.Object_Group LIKE '" & equipment & "' AND B.ACTIVE IS NOT NULL OR (A.Object_Group like '" & equipment & "' AND EquipmentType LIKE 'NSS' AND B.ACTIVE IS NULL) "
        
    Else
    
    query = "SELECT " & _
                "Object_Group " & _
            "FROM " & _
                "SITE_SPECIFIC_TAG_DATA AS A " & _
                "LEFT JOIN ( " & _
                "SELECT SITE_ID,[OPTION],ACTIVE FROM SITE_OPTIONS " & _
                "WHERE SITE_ID like '" & siteID & "' AND ACTIVE = 1 " & _
                "UNION " & _
                "SELECT SITE_ID,[COMPOSITE_OPTION],ACTIVE FROM SITE_OPTIONS_COMPOSITE " & _
                "WHERE SITE_ID like '" & siteID & "' AND ACTIVE = 1) " & _
                "AS B ON [A].[OPTIONS] = [B].[OPTION] " & _
            "WHERE  " & _
                "A.SITE_ID LIKE '" & siteID & "' " & _
                "AND Object_Group LIKE '" & equipment & "' AND B.ACTIVE IS NOT NULL OR (Object_Group like '" & equipment & "' AND EquipmentType LIKE 'NSS' AND B.ACTIVE IS NULL)"

    End If
    Set rs = DatabaseConnection.Connection_Query(query)
    
    If (rs.RecordCount > 0) Then
        ' Equipment Type exists on the asset
        Asset_Equipment_check = True
    Else
        ' Equipment TYpe does not exist on the asset
        Asset_Equipment_check = False
    End If
    
End Function

Public Function Site_Options_Table(app As Word.Application, doc As Word.Document, siteID As String, assetAbbreviation As String, scope As String, LOG_FILE As TextStream)
'   This function builds the site Options table
'   Inputs
'       - word application
'       - Word Document
'       - siteID
'       - scope
'       - LOG_FILE

'   Local Variables
    Dim rs As New ADODB.Recordset
    Dim tableTitle As String
    Dim query As String
    Dim tableNo As Integer
    Dim tableRows As Integer
    Dim i As Integer
    
    tableTitle = "Design Options"
    If (assetAbbreviation = "WP") Then
    
    query = "SELECT " & _
        "A.Option_ID, " & _
        "CASE " & _
            "WHEN B.ACTIVE = 1 THEN 'YES' " & _
            "WHEN B.ACTIVE = 0 THEN 'NO' " & _
            "WHEN B.ACTIVE IS NULL " & _
            "AND A.Option_ID NOT LIKE 'F' AND A.Asset_Type like '" & assetAbbreviation & "' THEN ( " & _
                "SELECT " & _
                    "CASE WHEN a.cnt > 0 THEN 'YES' WHEN a.cnt < 1 THEN 'NO' END AS cntVal FROM (SELECT count(*) AS cnt FROM SITE_OPTIONS WHERE SITE_ID LIKE '" & siteID & "' AND [OPTION] LIKE CONCAT(A.Option_ID, '%') AND ACTIVE = 1 ) AS a ) " & _
            "ELSE 'NO' END Active_Option, " & _
            "CASE When A.Selection_Flag = 1 THEN(SELECT ISNULL(STRING_AGG(ISNULL([DESCRIPTION], ' '), ','),ISNULL(A.[Description],'')) FROM SITE_OPTIONS WHERE SITE_ID LIKE '" & siteID & "' AND ACTIVE = 1 AND [OPTION] LIKE CONCAT(A.Option_ID, '%') ) " & _
            "ELSE " & _
            "A.[Description] END [DESCRIPTION] " & _
            "FROM app.ALL_OPTIONS AS A LEFT JOIN SITE_OPTIONS AS B ON A.Option_ID = B.[OPTION] AND B.SITE_ID LIKE '" & siteID & "' " & _
            "WHERE A.Asset_Type LIKE '" & assetAbbreviation & "' AND A.[Type] IN ('Electrical') " & _
            "ORDER BY A.Option_ID ASC "
        
    Else
        query = "SELECT " & _
        "A.Option_ID, " & _
        "CASE " & _
            "WHEN B.ACTIVE = 1 THEN 'YES' " & _
            "WHEN B.ACTIVE = 0 THEN 'NO' " & _
            "WHEN B.ACTIVE IS NULL " & _
            "AND A.Option_ID NOT LIKE 'F' AND A.Asset_Type like '" & assetAbbreviation & "' THEN ( " & _
                "SELECT " & _
                    "CASE WHEN a.cnt > 0 THEN 'YES' WHEN a.cnt < 1 THEN 'NO' END AS cntVal FROM (SELECT count(*) AS cnt FROM SITE_OPTIONS WHERE SITE_ID LIKE '" & siteID & "' AND [OPTION] LIKE CONCAT(A.Option_ID, '%') AND ACTIVE = 1 ) AS a ) " & _
            "ELSE 'NO' END Active_Option, " & _
            "CASE When A.Selection_Flag = 1 THEN(SELECT ISNULL(STRING_AGG(ISNULL([DESCRIPTION], ' '), ','),ISNULL(A.[Description],''))FROM SITE_OPTIONS WHERE SITE_ID LIKE '" & siteID & "' AND ACTIVE = 1 AND [OPTION] LIKE CONCAT(A.Option_ID, '%') ) " & _
            "ELSE " & _
            "A.[Description] END [DESCRIPTION] " & _
            "FROM app.ALL_OPTIONS AS A LEFT JOIN SITE_OPTIONS AS B ON A.Option_ID = B.[OPTION] AND B.SITE_ID LIKE '" & siteID & "' " & _
            "WHERE A.Asset_Type LIKE '" & assetAbbreviation & "' AND A.[Type] IN ('Both', 'Electrical') " & _
            "ORDER BY A.Option_ID ASC "
    End If
    
    ' Building Recordset
    Set rs = Connection_Query(query)
    
    'Building Table
    Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    
    'Function to update table rows
    tableNo = Find_Table_By_Title(doc, tableTitle, "Update")
    
    tableRows = doc.Tables.Item(tableNo).Rows.Count
    i = 0
    For j = 1 To tableRows
        If (InStr(doc.Tables.Item(tableNo).Cell(i + 1, 2).Range.text, "No")) Then
               doc.Tables.Item(tableNo).Rows(i + 1).Shading.BackgroundPatternColor = 11842740
        End If
        i = i + 1
    Next
    
End Function
Public Function NonStandard_RTU_Communication(app As Word.Application, doc As Word.Document, siteID As String, scope As String, LOG_FILE As TextStream)
    ' This function updates all nonstandard RTU Communications tables
    ' In the event that the station has no Non-Standard RTU Communication the tables are deleted
    '   Inputs (key)
    '       SiteID
    
    Dim rs As New ADODB.Recordset
    Dim tableTitle As String
    Dim Caption As String
    Dim query As String
    Dim sectionTag As String
    Dim tableArray(2) As Variant
    Dim modbusType(2) As Variant
    Dim captions(2) As Variant
    Dim tempScope As String
    Dim sectionActive As String
    
    sectionActive = False
    tableArray(1) = "Non-Standard Communications Sent to SEQ Water"
    tableArray(2) = "Non-Standard Communications Recieved from SEQ Water"
    captions(1) = "Table 36: Non-Standard Communications Sent to SEQ Water"
    captions(2) = "Table 37: Non-Standard Communications Received from SEQ Water"
    
    modbusType(1) = "40"
    modbusType(2) = "41"
        
    'Building input tag table (Recieved from SEQ water)
    For i = 1 To 2
        tableTitle = tableArray(i)
        Caption = captions(i)
        tempScope = scope + "Inserting [" + tableTitle + "]"
        'Building Query
        query = "SELECT " & _
            "CONCAT(TAG_DATA.Asset_Name, TAG_DATA.Asset_Description) AS [DESCRIPTION], " & _
            "MODBUS_TAGS.Tag, " & _
            "CASE " & _
                "WHEN TAG_DATA.Tag_Data_Type = 'BOOL' THEN 'Boolean' " & _
                "WHEN TAG_DATA.Tag_Data_Type = 'REAL' THEN 'IEEE Real' " & _
                "WHEN TAG_DATA.Tag_Data_Type = 'INT' THEN 'Integer' " & _
                "ELSE '' " & _
            "END AS Data_Type, " & _
            "MODBUS_TAGS.Modbus_Address " & _
        "FROM " & _
            "[SITE_MODBUS_RTU_COMMUNICATION] AS MODBUS_TAGS " & _
            "JOIN SITE_SPECIFIC_TAG_DATA AS TAG_DATA ON MODBUS_TAGS.Tag_Id = TAG_DATA.ID " & _
            "AND TAG_DATA.SITE_ID LIKE '" & siteID & "' " & _
        "WHERE " & _
            "MODBUS_TAGS.Site_Id LIKE '" & siteID & "' " & _
            "AND MODBUS_TAGS.Modbus_Address LIKE '" & modbusType(i) & "%'"
            
     Set rs = DatabaseConnection.Connection_Query(query)
     
    ' If records exists insert the table
     If (rs.RecordCount > 0) Then
        sectionActive = True
        Call Log_Line(tempScope, 3, LOG_FILE)
        Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)

     Else
        'Removing Table, caption and PeerDataMap placeholder
        Call Log_Line(scope + "Deleting " + tableTitle, 3, LOG_FILE)
        Call Remove_Previous_Paragraph(app, doc, tableTitle, 1, "table")
        Call Table_Deletion(app, doc, tableTitle, LOG_FILE)
        'Call removeTagSection(app, doc, "$$NS_RTU_COMMUNICATIONS$$")
        
     End If
     
    Next
        If (sectionActive) Then
        ' Removing placeholder tags
        Call Select_Current_Paragraph(app, doc, "$$NS_RTU_COMMUNICATIONS:START$$", "word", "delete")
        Call Select_Current_Paragraph(app, doc, "$$NS_RTU_COMMUNICATIONS:END$$", "word", "delete")
        'Call Variable_Remove(app, doc, "$$NS_RTU_COMMUNICATIONS:START$$", LOG_FILE)
        'Call Variable_Remove(app, doc, "$$NS_RTU_COMMUNICATIONS:END$$", LOG_FILE)
    Else
        ' Removing section about DNP3 information
        Call Delete_Section(app, doc, "$$NS_RTU_COMMUNICATIONS$$")
        
    End If
    
    
    
End Function

Public Function Peer_To_Peer_DNP3(app As Word.Application, doc As Word.Document, siteID As String, scope As String, LOG_FILE As TextStream, process As Integer)
' This function Builds the Peer to peer DNP3 table

    Dim rs As New ADODB.Recordset
    Dim query As String
    Dim tableTitle As String
    Dim sourceSites As String
    Dim Caption As String

    ' Build first query which build list of sites that asset recieves tags from
    '--------------------------------------------------- SITE AS SENDER ---------------------------------------------------------------------------------
   Caption = "Table 15: Peer to Peer DNP3 Signals Sent by RTU"
   tableTitle = "Peer to Peer DNP3 Signals Sent by RTU"
   query = "SELECT " & _
                "DISTINCT [SOURCE] " & _
            "FROM " & _
                "PEER_DATA_MAP " & _
            "WHERE " & _
                "[Source] LIKE '" & siteID & "' "
                
    Set rs = DatabaseConnection.Connection_Query(query)
        
        
    Call UI_update(scope, process)
    If (rs.RecordCount > 0) Then
        Call UI_update(scope + "Building source List", process)
        'Building string for next query
            rs.MoveFirst
             Do Until rs.EOF = True
                    For j = 1 To rs.Fields.Count
                        If j = 1 Then
                            sourceSites = "'" + rs.Fields(j - 1) + "'"
                        Else
                            sourceSites = sourceSites + ",'" + rs.Fields(j - 1) + "'"
                        End If
                    Next
                    rs.MoveNext
                Loop
        rs.Close
        Set rs = Nothing
        query = ""
        'Building New record set which holds Data for the table
         query = "SELECT " & _
                    "CONCAT(C.Asset_Name, ' ', C.Asset_Description) AS Tag_Description,Main.Source_Tag,ISNULL(C.EU, '') AS SRC_EU,ISNULL(C.Tag_Data_Type, '') AS SRC_DNP3,ISNULL(C.DNP3_Point_Number, '') AS SRC_DNP3,ISNULL(C.Analogue_Deviation, '') AS Analogue_Deviation,Main.[Source] AS SRC,ISNULL(B.DNP3_Point_Number, '') AS DST_DNP3 " & _
                "FROM " & _
                    "PEER_DATA_MAP AS Main " & _
                    "JOIN ( " & _
                        "SELECT " & _
                            "A.ID,Source_Tag,Source_TagID,B.SITE_ID,B.DNP3_Point_Number " & _
                        "FROM " & _
                            "PEER_DATA_MAP AS A " & _
                            "JOIN SITE_SPECIFIC_TAG_DATA AS B ON B.ID = A.Source_TagID " & _
                            "AND B.SITE_ID IN ('" & siteID & "') " & _
                        "WHERE " & _
                            "A.[SOURCE] IN ('" & siteID & "') " & _
                    ") AS B ON B.ID = Main.Id " & _
                    "JOIN SITE_SPECIFIC_TAG_DATA AS C ON Main.Source_TagID = C.ID " & _
                    "AND C.SITE_ID LIKE '" & siteID & "' " & _
                "WHERE " & _
                    "MAIN.Source LIKE '" & siteID & "' " & _
                    "AND MAIN.[Group] = 1 "
                                
        Set rs = DatabaseConnection.Connection_Query(query)
                
        Call UI_update(scope + tableTitle + "Building", process)
        Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    Else
        Call UI_update(scope + "No peer to peer Deleting Table [" + tableTitle + "]", process)
        Call Remove_Previous_Paragraph(app, doc, tableTitle, 1, "table")
        Call Table_Deletion(app, doc, tableTitle, LOG_FILE)
    End If
    
    '--------------------------------------------------- SITE AS RECIVER ---------------------------------------------------------------------------------
    Caption = "Table 15: Peer to Peer DNP3 Signals Sent to RTU"
    tableTitle = "Peer to Peer DNP3 Signals Received by RTU"
    query = "SELECT " & _
                "DISTINCT [SOURCE] " & _
            "FROM " & _
                "PEER_DATA_MAP " & _
            "WHERE " & _
                "Destination LIKE '" & siteID & "' "
                
    Set rs = DatabaseConnection.Connection_Query(query)
        
        
    Call UI_update(scope, process)
    If (rs.RecordCount > 0) Then
        Call UI_update(scope + "Building source List", process)
        'Building string for next query
            rs.MoveFirst
             Do Until rs.EOF = True
                    For j = 1 To rs.Fields.Count
                        If j = 1 Then
                            sourceSites = "'" + rs.Fields(j - 1) + "'"
                        Else
                            sourceSites = sourceSites + ",'" + rs.Fields(j - 1) + "'"
                        End If
                    Next
                    rs.MoveNext
                Loop
        rs.Close
        Set rs = Nothing
        query = ""
        'Building New record set which holds Data for the table
         query = "SELECT " & _
                        "CONCAT(C.Asset_Name, C.Asset_Description),Main.Destination_Tag,ISNULL(C.EU, '')AS SRC_EU,ISNULL(C.Tag_Data_Type, '') AS SRC_DNP3,ISNULL(C.DNP3_Point_Number, '') AS SRC_DNP3, " & _
                        "ISNULL(C.Analogue_Deviation, ''),Main.[Source] AS SRC,ISNULL(B.DNP3_Point_Number, '') AS DST_DNP3 " & _
                "FROM " & _
                    "PEER_DATA_MAP AS Main " & _
                    "JOIN ( " & _
                        "SELECT " & _
                            "A.ID, Destination_Tag, Destination_TagID, B.SITE_ID, B.DNP3_Point_Number FROM PEER_DATA_MAP AS A  JOIN SITE_SPECIFIC_TAG_DATA AS B ON B.ID = A.Source_TagID  AND B.SITE_ID IN (" & sourceSites & ") " & _
                        "WHERE " & _
                            "A.[SOURCE] IN (" & sourceSites & ") " & _
                    ") AS B ON B.ID = Main.Id " & _
                    "JOIN SITE_SPECIFIC_TAG_DATA AS C ON Main.Destination_TagID = C.ID " & _
                    "AND C.SITE_ID LIKE '" & siteID & "' " & _
                "WHERE " & _
                "MAIN.Destination LIKE '" & siteID & "' " & _
                "AND MAIN.[GROUP] = 1"
                
        Set rs = DatabaseConnection.Connection_Query(query)
                
        Call UI_update(scope + tableTitle + "Building", process)
        Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    Else
        Call UI_update(scope + "No peer to peer Deleting Table [" + tableTitle + "]", process)
        Call Remove_Previous_Paragraph(app, doc, tableTitle, 1, "table")
        Call Table_Deletion(app, doc, tableTitle, LOG_FILE)
    End If
    

End Function
Public Function Remove_Previous_Paragraph(app As Word.Application, doc As Word.Document, location As String, numParagraphs As Integer, typeSwitch As String)
    ' This function looks at the paragraph of the above selection and deletes the paragraph(s)
    ' Inputs
    '   - location      String      (Tag, table)
    '   - numParagraphs Integer     (number of paragraphs above the selection
    '   - typeSwitch    String      (This switch determins what is going to be filtered, table,tag) inputs (table,"")
    '
    'Local varibales
    Dim tableNumber As Integer
    Dim i As Integer
    Dim totalCharacters As Integer
    i = 0
    
    
    If (typeSwitch = "table") Then
        ' If the input is a table then it will remove the table caption
        tableNumber = Find_Table_By_Title(doc, location, "Update")
        
        doc.Tables.Item(tableNumber).Select
        For i = 1 To numParagraphs
        ' Deleting caption and line break
            app.Selection.Previous(Unit:=wdParagraph, Count:=1).Delete
        Next i
    ElseIf (typeSwitch = "LineBreak") Then
        tableNumber = Find_Table_By_Title(doc, location, "Update")
        
        doc.Tables.Item(tableNumber).Select
        ' This is reserved for normal tags (For future development)
        totalCharacters = app.Selection.Previous(Unit:=wdParagraph, Count:=1).Characters.Count
        If (totalCharacters > 0) Then
            'Inserting line break
            'Insert function
            Debug.Print ("we got here")
        End If
    ElseIf (typeSwitch = "heading") Then
        ' This tag checks the above paragraph to see if it is empty
        Call insertAboveTag(app, doc, location, "heading")
        
        totalCharacters = app.Selection.Previous(Unit:=wdParagraph, Count:=1).Characters.Count
        If (totalCharacters < 2) Then
            'app.Selection.TypeParagraph
            'Inserting line break
            'Insert function
            Debug.Print ("we got here")
        End If
    Else
        ' Else catch statment (Error checking maybe)
    End If

End Function

Public Function Select_Current_Paragraph(app As Word.Application, doc As Word.Document, tag As String, tagType As String, method As String)
    ' Currently Not used
    ' This function finds a paragraph based on a single input (table or word)
    Dim tableNumber As Integer
    Dim i As Integer
    i = 0
    Dim removeRange As Range, rangeStart As Integer, rangeEnd As Integer
    Dim found As Boolean
    
    If (tagType = "table") Then
        ' If type is a table then the table number needs to be determined
        tableNumber = Find_Table_By_Title(doc, tag, "Update")
        doc.Tables.Item(tableNumber).Select
        
    ElseIf (tagType = "word") Then
        ' This method first uses the insertAboveTag function to locate the tag
        found = insertAboveTag(app, doc, tag, "find")
        If (found) Then
            'Selecting start and end of the paragraph
            app.Selection.StartOf Unit:=wdParagraph
            app.Selection.MoveEnd Unit:=wdParagraph
            app.Selection.Delete
        End If

    ElseIf (tagType = "section") Then
        
    Else
    
    End If
End Function
Public Function Insert_Paragraph(app As Word.Application, doc As Word.Document, paragraph As String, direction As String)
    ' This function takes a paragraph and inserts it below the selected point.
    ' Inputs
    '   - pragraph  |   This is the inserted paragraph  |String
    'Outputs
    '   - Outputs   |   None
    
    If (direction = "before") Then
        app.Selection.InsertBefore (paragraph)
    ElseIf (direction = "after") Then
        app.Selection.InsertAfter (paragraph)
    Else
    
    End If

        With app.Selection.Font
            .Bold = False
            .ColorIndex = wdBlue
            .Italic = True
            .Size = 11
            .name = "Calibri"
        End With
        app.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify



End Function

Public Function Select_Section_Headings(app As Word.Application, doc As Word.Document, method As String, formating As String)
    ' This function loops through all sectoin headings and then can deleting/updating as requried
    ' When the method is set to delete the paragraph above the section will be deleted
    
    For Head_n = 1 To 5
    Head = ("Heading " & Head_n)
    
    ' This selected the holw document for looping through each section
    app.Selection.HomeKey wdStory, wdMove
    Do
        ' Selecting a section heading
        With app.Selection
            .MoveStart Unit:=wdLine, Count:=1
            .Collapse direction:=wdCollapseEnd
        End With
        With app.Selection.Find
          .ClearFormatting:          .text = "":
          .MatchWildcards = False:   .Forward = True
          .Style = doc.Styles(Head)
          
         If (.Execute = False) Then
            ' If the section heading can not be found then the selection heading number is increased
            GoTo Level_exit
        Else
            ' Checking to see if the paragraph has contains information, if it is null it will be deleted
            If (method = "delete") Then
                paraNumber = doc.Range(0, app.Selection.Range.End).Paragraphs.Count
                If Len(doc.Paragraphs(paraNumber - 1).Range.text) = 1 Then
                    doc.Paragraphs(paraNumber - 1).Range.Delete
                End If
            End If
            
            .ClearFormatting
        End If
        End With
        If app.Selection.Style = "Heading 1" Then GoTo Level_exit
        app.Selection.Collapse direction:=wdCollapseStart
   Loop
Level_exit:
    Next Head_n
        
    
End Function

Public Function Generator_Details(app As Word.Application, doc As Word.Document, siteID As String, assetAbbreviation As String, scope As String, LOG_FILE As TextStream, process As Integer)
' This function inserts the generator values for the generator table
'   - Generator size (Currently for all sites this is "No Generator installed")
'   - maximum concurrent pumps while generator is running

' Local variables
    Dim query As String
    Dim rs As New ADODB.Recordset
    Dim generatorSize As String
    Dim generatorMaxPumpsRunning As String
    
    generatorSize = "No Generator Installed"
    If (assetAbbreviation = "SP") Then
        ' Currently SP is the only asset which has pump limiting interlocking for generator and mains.
        query = "SELECT " & _
                    "* " & _
                "FROM " & _
                    "SITE_OPTIONS " & _
                "WHERE " & _
                    "SITE_ID LIKE '" & siteID & "' " & _
                    "AND [OPTION] LIKE 'O_' " & _
                    "AND ACTIVE = 1 "
        
        Set rs = Connection_Query(query)
            
        If (rs.RecordCount > 0) Then
            generatorMaxPumpsRunning = "0"
        Else
            generatorMaxPumpsRunning = "Not Applicable"
        End If
        
        rs.Close
        Set rs = Nothing
        query = ""
        
    Else
        ' For all other assets
        generatorMaxPumpsRunning = "Not Applicable"
        
    End If
    
    'Updating the template
    Call Variable_Replacement(app, doc, "$generatorSize", generatorSize, process, LOG_FILE)
    Call Variable_Replacement(app, doc, "$generatorMaxPumpsRunning", generatorMaxPumpsRunning, process, LOG_FILE)



End Function
Function Table_Row_Background_Colour(app As Word.Application, doc As Word.Document, tableName As String, keyWord As String, columnNumber As Integer, colour As String, scope As String)
    ' This function changes the background colour of a row.
    ' Key Inputs
    '       - tableName: Table that will be affected
    '       - keyWord: Word that will be used to determine the colour
    '       - columnNumber: The column with the key word to check
    '       - colour: Colour to set the background colour
    
    'Local Variables
    Dim colourId As String
    
    Select Case colour
     Case "Grey"
        colourId = "11842740"
     Case ""
        colourId = "-603914241"
     Case Else
        colourId = "588001"
     End Select
     

    tableNo = Find_Table_By_Title(doc, tableName, "update")
    
    tableRows = doc.Tables.Item(tableNo).Rows.Count
    i = 0
    For j = 1 To tableRows
        If (InStr(doc.Tables.Item(tableNo).Cell(i + 1, 2).Range.text, keyWord)) Then
               doc.Tables.Item(tableNo).Rows(i + 1).Shading.BackgroundPatternColor = 11842740
        End If
        i = i + 1
    Next
End Function
