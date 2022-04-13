Attribute VB_Name = "AssetSpecificFunctions"
Option Compare Database

Public Function SP_Specific_Tables(app As Word.Application, doc As Word.Document, siteID As String, scope As String, LOG_FILE As TextStream)
    '-------------------------->  Sewage Pump Tables<------------------------------------------------------
    ' This function updates the two unique sewage pump tables
    '   Wet Well Lookup Table
    '   Key Wet Well Levels
    'Call
    Dim tableTitle As String
    Dim query As String
    Dim rs As New ADODB.Recordset
    
    tableTitle = "Key Wet Well Levels"
    scope = "Inserting Standard Sewage Pumping Station Tables: "
    
    scope = scope + tableTitle
    'Read References And Inputs Data
    query = "SELECT " & _
        "A.[Tag_Description], " & _
        "A.[Tag], " & _
        "COALESCE(ISNULL(ISNULL(B.SITE_SPECIFIC,B.Default_value), C.[VALUE]),'') AS [VALUE], " & _
        "COALESCE(ISNULL(B.EU, C.UNITS),'') AS [UNITS] " & _
    "FROM " & _
        "Look_Up_Table_FS AS A " & _
        "LEFT JOIN SITE_SPECIFIC_TAG_DATA AS B ON A.Tag = CONCAT(B.Object_Group, B.Tag_Attribute) " & _
        "AND B.SITE_ID LIKE '" & siteID & "' " & _
        "LEFT JOIN Look_Up_Table_FS_Values AS C ON A.ID = C.TAG_KEY " & _
        "AND C.SITE_ID LIKE '" & siteID & "' " & _
    "WHERE " & _
        "A.FS_Table LIKE '" & tableTitle & "' " & _
    "ORDER BY " & _
        " A.[ORDER] ASC "

    'Connetion to Database and returning recordset
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line(scope + " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation(app, doc, rs, tableTitle, LOG_FILE)
    
    'Closing Record set connetion and clearing recordset
    rs.Close
    Set rs = Nothing
    
    ' This table is  bit funky and could be implemented other ways
    '   call the record set 3 times with each record set being one column
    Dim tagAttribute(3) As Variant
    Dim columnOffSet As Integer
    Dim length As Integer
    
    tableTitle = "Wet Well Lookup Table"
    columnOffSet = 0
    tableTitle = "Wet Well Lookup Table"
    tagAttribute(1) = "[_]krWWLLookup"
    tagAttribute(2) = "[_]krRemStorCap"
    tagAttribute(3) = "[_]krCurrStorVol"
    
    length = UBound(tagAttribute) - LBound(tagAttribute)
    
    
    For columnOffSet = 1 To length
    
    
    query = "SELECT " & _
            "Site_Specific " & _
        "FROM " & _
            "SITE_SPECIFIC_TAG_DATA " & _
        "WHERE " & _
            "Object_Group LIKE 'LIT0001' " & _
            "AND Tag_Attribute LIKE '" & tagAttribute(columnOffSet) & "[0-9][0-9]%' " & _
            "AND SITE_ID LIKE '" & siteID & "' " & _
        "ORDER BY " & _
            "Tag_Attribute DESC "
        
        Call Custom_Table_Instantiation(app, doc, siteID, query, tableTitle, columnOffSet, scope, LOG_FILE)
        
    Next columnOffSet
    
    
    
    
    
End Function

Public Function Custom_Table_Instantiation(app As Word.Application, doc As Word.Document, siteID As String, query As String, tableTitle As String, columnNum As Integer, scope As String, LOG_FILE As TextStream)
    ' This function itterates through a table however inserts all values vertically
    ' A column number must be specified which will determine which row is developed
    '   - This function was developed for sewage pumping stations (wet Well level)
    
    Dim rs As New ADODB.Recordset
    Dim queryRows As Integer
    Dim tableNo As Integer
    Set rs = Connection_Query(query)
    
    queryRows = rs.RecordCount
    Call Log_Line(scope & " Starting Table Instantiation", 2, LOG_FILE)
    Call Log_Line(scope & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)

    If queryRows > 0 Then
        tableNo = Find_Table_By_Title(doc, tableTitle, "update")
        If tableNo > 0 Then
        Call Log_Line("Update Table: TableFound", 2, LOG_FILE)
            rs.MoveFirst
            i = 1
            Do Until rs.EOF = True
                For j = 1 To rs.Fields.Count
                    doc.Tables.Item(tableNo).Cell(i + 1, j + columnNum).Range.text = rs.Fields(j - 1)
                    'doc.Tables.Item(tableNo).Cell(i + 1, j).Range.text = Replace(rs.Fields(j - 1), "_", " ")
                Next
                rs.MoveNext
                i = i + 1
            Loop

            Call Log_Line(scope & tableTitle & " - Populated " & queryRows & " rows.", 3, LOG_FILE)
        Else
            Call Log_Line(scope & tableTitle & " - Table not found.", 3, LOG_FILE)
        End If
        Call Log_Line(scope & tableTitle & " - Completed.", 3, LOG_FILE)
        Call Log_Line_Break(LOG_FILE, 1)
    End If
    
    
    
    rs.Close
    Set rs = Nothing
    query = ""

    





End Function
