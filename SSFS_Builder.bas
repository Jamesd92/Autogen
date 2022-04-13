Attribute VB_Name = "SSFS_Builder"
'VBA Menu ==> Tools ==> References & Add:
'  - Microsoft Word 16.0 Object Library
'  - Microsoft ActiveX Data Objects x.x library
'  - Microsoft Scritpting Runtime
'  - Microsoft VBScript Regular Expressions 5.5

'Word Menu ==> File ==> Options ==> Advanced ==> Show Document Content
'  - Check "Show bookmarks" Option

'Form entries for:
'  - Site ID
'  - Revision Details for this version (8 fields with default values)
'      - Maybe, just present editable revision history table on selection of site id
'  - Path to word template
'  - Optionally: Manual Override of site Asset Type Selections (and those selections)

'Supporting Table
'  - Site ID and Bool for all Asset Type Selections to include associated data

'Paths to Set
'  - Template Path
'  - File Outpath
'  - Needs to be added
'  - DebugPath

'Variable Definition
'   - Global variables (SCREAMING_SNAKE_CASE)
'   - Modules (PascalCase)
'   - local varibles (camelCase)
'   - Functions (Pascale_Snake_Case)


'----------------------------------- Cycles through tables, return desired table position ------------------------------





'------------------------------------------------------------------------------------------------------------------------
'------------------------------------- Debugging Function: Write out to file --------------------------------------------
'Functionality written in subModule (Debugging)
'-------------------------------------------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'--------------------------------------------------- GLOBAL Varibales ----------------------------------------------------
' File Paths
Global TEMPLATE_FOLDER_PATH As String
Global SESSION_LEVEL As Integer
Global TIME_STAMP As String
Global LOG_FILE As TextStream
Global LOG_FILE_NAME As String


Public Function BuildSSFS(process As Integer, TestingSite As String)
    '******************************************************************************************************************
    'Dimensioning
    
    
    If (process = 1) Then
    On Error GoTo ErrorHandler
    End If
    'General use
    Dim i, j, k As Integer
    
    'Configuration
    Dim siteID As String
    
    'LOG_FILE
    Dim fso As New FileSystemObject
    Dim LOG_FILE As TextStream
    Dim buildTIME_STAMP As String
    Dim LogLevel As Integer
    Dim scope As String
    Dim uiString As String
    'Regex
    Dim regexString As String
    Dim regexObject As Object
    Set regexObject = New RegExp
    
    'Word Document
    Dim wordapp As New Word.Application
    Dim wordDoc As Word.Document
    Dim bookMarkRange As Word.Range
    Dim templateFilePath As String
    'Dim TEMPLATE_FOLDER_PATH As String
    Dim tableTitle As String, tableNo As Integer
    Dim tagFindString As String, tagReplaceString As String, tagRemoveString As String
    
    
    'Database
    Dim cn As New ADODB.connection
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rsUp As New ADODB.Recordset
    Dim rsDown As New ADODB.Recordset
    Dim query As String
    Dim queryRows As Integer
    Dim exists As Boolean
    
    'Dictionaries
    Dim replaceDict As Scripting.Dictionary
    Dim removeDict As Scripting.Dictionary
    Dim scectionRemoveDict As Scripting.Dictionary
    Set replaceDict = New Scripting.Dictionary
    Set removeDict = New Scripting.Dictionary
    Set scectionRemoveDict = New Scripting.Dictionary
    Dim key As Variant
    
    'Strings Words and Formating
    Dim paragraph As String
    Dim currentEquipmentType As String
    Dim previousEquipmentType As String
    Dim paragraphCounter As Integer
    Dim tableTypeParameter As Boolean
    Dim tableRange As Word.Range
    
    
    ' Table and Caption variables
    Dim Caption As String
    Dim tag As String
    Dim tagRemoveSection As String
    
    ' Table, Caption Section
    Dim tableInfo As String
    Dim sectionHeading As String
    Dim Sectioncounter As Integer
    
    ' Equipment Switches
    Dim compositeSite As String
    Dim generatorExists As Boolean
    Dim pumpString As String
    
    'Variables needed to Build the Document

    'Site_List variables
    Dim assetType As String
    Dim address As String
    Dim name As String
    Dim assetAbbreviation As String
    Dim GPS As String
    Dim boardType As String
    Dim allSites As String
    Dim compositeType As String
    
    'Document References Variables
    Dim ssfsDocumentNumber As String
    Dim stdAssetDocumentDesignNum As String
    Dim stdAssetDocumentDesignName As String
  
    'Siteoverview variables
    Dim locationDescription As String
    
    'DNP3 Master and Slave Addressing
    Dim DNP3MasterAddr As String
    Dim DNP3SlaveAddr As String
    
    'Equipment Table Variables
    Dim equipment As String
    Dim numPumps As String
    Dim concurrentPumpsRunning As String
    Dim upQuery As String
    Dim downQuery As String
    
    
    
    'SSFS variables
    Dim siteDetails As String
    Dim UpgradeTypeText As String
    Dim processNarrative As String
    Dim sitePurpose As String
    Dim networkDetails As String
    Dim peerCommunicationsText As String
    Dim nonStandardHighLevelText As String
    Dim nonStandardText As String
    Dim nonStandardEquipment As String
    Dim nonStandardInstrumentation As String
    Dim nonstandardControlSystem As String
    Dim nonstandardProgram As String
    Dim nonstandardCommunication As String
    Dim nonstandardSCADA As String
    
    ' Template Variables
    Dim tableTemplate As String
    Dim siteArray As String

    
    'Image Variables
    'Dim peerDataMap As Image
    'Dim networkImage As Image
    'Dim siteImage As Image
    'Dim SCADAImage AS Image
    Dim imageAddress As String
    Dim userLogon As String
    
    'Revision Variables
    Dim revisionDocument As String
    Dim revisionClass As String
    
    
    
    'FilePaths
    Dim workDir As String
    Dim RevisionNum As String
    Dim imagePath As String
    

    TEMPLATE_FOLDER_PATH = Application.CurrentProject.Path & "\"
    
    'ErroChecking
    Dim rowCountChk As Integer
    Dim ERROR As Boolean
    Dim errorString As String
    
    
    '******************************************************************************************************************
    'Initialisation
    
    ' Methods
    If (process = 1) Then
        LogLevel = 3
        siteID = TestingSite
        RevisionNum = "AutoGen"
        wordapp.Visible = False
    ElseIf (process = 0) Then
        LogLevel = Forms!AutoGen.Form!LogLevel.value '1=Info, 2=Detail, 3=Debug
        siteID = Forms!AutoGen.Form!siteID.value
        RevisionNum = Forms!AutoGen.Form!RevisionNum.value
        wordapp.Visible = Forms!AutoGen.Form!Visible.value
    Else
    
    End If
    
    
    'Configuration
    uiString = "Starting"
    Call UI_update(uiString, process)
    templateFilePath = TEMPLATE_FOLDER_PATH & "SSFS_Template_v01.docx"
    userLogon = Environ$("username")

        
    'Debugging Session level
    SESSION_LEVEL = 3
    buildTIME_STAMP = Format(Now(), "yymmdd-hhMMSS")
    LOG_FILE_NAME = "Buildlog_" & siteID & "_" & buildTIME_STAMP & ".log"
    Set LOG_FILE = fso.OpenTextFile(TEMPLATE_FOLDER_PATH & LOG_FILE_NAME, ForWriting, True)

    Call Log_Line("STARTING TO LOG", 1, LOG_FILE)
    Call Log_Line("File Output Path:" + templateFilePath, 3, LOG_FILE)
    Call Log_Line("-------------------------------------------------------------", 1, LOG_FILE)
    
    'Word Document
    Call Log_Line("Opening Template", 3, LOG_FILE)
    Set wordDoc = wordapp.Documents.Open(FileName:=templateFilePath, ReadOnly:=False)
    If (process = 1) Then
            wordDoc.SaveAs2 (TEMPLATE_FOLDER_PATH & "Comparison\" & "C1282B-WP01-IDD-" + siteID + "-EL19-SPC-0001_AutoGen.docx")
    ElseIf (process = 0) Then
            wordDoc.SaveAs2 (TEMPLATE_FOLDER_PATH & "C1282B-WP01-IDD-" + siteID + "-EL19-SPC-0001_Rev" & RevisionNum & "_" & buildTIME_STAMP & ".docx")
    Else
    
    End If

    wordDoc.Activate
    Call Log_Line("Opening Word", 3, LOG_FILE)
    Call Log_Line_Break(LOG_FILE, 1)
    
    
    '------------ TESTING AREA -----------------------------


    'Call Remove_Previous_Paragraph(wordapp, wordDoc, tableTitle, 1, "LineBreak")
    '-------------------------------------------------------
    
    '---------------------------------- Site Images -----------------------------------------------
    ' Currently Images live in sharepoint. This requires literal paths to be set to uploaded these
    ' images
    ' building following variables
    '   siteImage - $siteImage
    '   SCADAImage - Not used right now
    '   peerDataMap - ?? Not used?
    '   networkImage - $networkImage
    
    'Inserting Site Image
    scope = "Image Insert [siteImage] :"
    Call UI_update(scope, process)
    tag = "$siteImage"
    
    If (userLogon = "james.don") Then
        imagePath = "C:\Users\" + userLogon + "\LOGICAMMS LTD\QUU - Documents\1b  Design\02 Infrastructure\Software Programming\Design Deliverables - Site Specific\Sites\" + siteID + "\" + siteID
    Else
        imagePath = "C:\Users\" + userLogon + "\OneDrive - LOGICAMMS LTD\Design Deliverables - Site Specific\Sites\" + siteID + "\" + siteID
    End If
    imageAddress = imagePath + " - Location Map.png"
    Call Insert_Image(wordapp, wordDoc, imageAddress, tag, scope, LOG_FILE)
    Call Select_Current_Paragraph(wordapp, wordDoc, tag, "word", "delete")
    'Call Variable_Remove(wordapp, wordDoc, tag, LOG_FILE)
    
    
    scope = "Image Insert [networkImage] :"
    Call UI_update(scope, process)
    tag = "$networkImage"
    imageAddress = imagePath + " - Network Diagram.png"
    Call Insert_Image(wordapp, wordDoc, imageAddress, tag, scope, LOG_FILE)
    Call Select_Current_Paragraph(wordapp, wordDoc, tag, "word", "delete")
    'Call Variable_Remove(wordapp, wordDoc, tag, LOG_FILE)
    
    ' Inserting Water Network Image
    'scope = "Image Insert [SCADA] :"
    'Call UI_update(scope)
    'tag = "$networkImage"
    'imageAddress = "C:\Users\" + userLogon + "\LOGICAMMS LTD\QUU - Documents\1b  Design\02 Infrastructure\Software Programming\Design Deliverables - Site Specific\Sites\" + siteID + "\" + siteID + " - SCADA Mimic.png"
    'Call Insert_Image(wordApp, wordDoc, imageAddress, tag, scope, LOG_FILE)
    'Call Variable_Remove(wordApp, wordDoc, tag, LOG_FILE)
    
    
   
    '******************************************************************************************************************
    'Building variable list
    'This section queries multiple tables to return all veriables required
    'A Query string is developed and passed the function DatabaseConnection.ConnectionQuery, which returns a record set
    'Values are are then extracted from the recored set and placed in variables
    '-----------------> Each Variable is wrapped with Nz(input) which masks null values to ""<--------------------------
    '-----------------> Without this VBA will fault <-------------------------------------------------------------------
    'Tables which contain information
    '   - Site_List
    '   - SSFS_Text_Fields
    '   - Document_references
    '   - Site_Overview_details
    '   - site_Image (Need to implement)
    '   - EquipmentTable
    '---------------
    'Selecting from site_list table
    ' confirming variables are coming from correct key
    ' building following variables
    '   name
    '   assetType
    '   address
    '   assetAbbreviation
    '   GPS
    '   boardType - SPS only
    '---------------
    scope = "Building Variable List"
    Call UI_update(scope, process)
    Call Log_Line("Building Variables", 1, LOG_FILE)
    query = "SELECT * FROM [Infrastructure].[DBO].[SITE_LIST] " & _
            "WHERE SITE_ID LIKE '" & siteID & "'"
    Call Log_Line("Query: " + query, 3, LOG_FILE)
    
    'Connection to Database
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line("Site_List queried", 1, LOG_FILE)
    queryRows = rs.RecordCount

    Call Log_Line("Current SiteID: " + siteID, 3, LOG_FILE)
    Call Log_Line("Site Returned from query: " + rs.Fields("site_ID").value, 3, LOG_FILE)
    replaceDict.Add "siteID", siteID
    
    assetType = Nz(rs.Fields("name").value, "")
    replaceDict.Add "assetType", assetType
    Call Log_Line("AssetType: " + assetType, 3, LOG_FILE)
    'compositeType
    compositeType = Nz(rs.Fields("Composite_Type").value, "")
    replaceDict.Add "compositeType", compositeType
    Call Log_Line("AssetType: " + assetType, 3, LOG_FILE)
    
    name = Nz(rs.Fields("Asset_name").value, "")
    replaceDict.Add "name", name
    Call Log_Line("name: " + name, 3, LOG_FILE)
    
    address = Nz(rs.Fields("address").value, "")
    replaceDict.Add "address", address
    Call Log_Line("address: " + address, 3, LOG_FILE)
    
    assetAbbreviation = Nz(rs.Fields("ASSET_Type").value, "")
    replaceDict.Add "assetAbbreviation", assetAbbreviation
    Call Log_Line("assetAbbreviation: " + assetAbbreviation, 3, LOG_FILE)
   
    boardType = Nz(rs.Fields("Std_Design").value, "")
    replaceDict.Add "boardType", boardType + " "
    Call Log_Line("assetAbbreviation: " + assetAbbreviation, 3, LOG_FILE)
    
    allSites = Nz(rs.Fields("ALL_SITES").value, "")
    replaceDict.Add "allSites", allSites + " "
    Call Log_Line("assetAbbreviation: " + assetAbbreviation, 3, LOG_FILE)
    
    GPS = Nz(rs.Fields("GPS").value, "")
    replaceDict.Add "GPS", GPS
    Call Log_Line("GPS: " + GPS, 3, LOG_FILE)
    'Closing connetion to the Database and nulling the record set
    rs.Close
    Set rs = Nothing
    query = ""
    Call Log_Line("Site_List connection closed", 1, LOG_FILE)
    Call Log_Line("-------------------------------------------------------------", 1, LOG_FILE)
    Call Log_Line_Break(LOG_FILE, LogLevel)
    
    '-------------------------------------------------------------- Site Specific Text Fields ---------------------------------------------
    'Selecting from SSFS_Site_List
    ' confirming variables are coming from correct key
    ' building following variables
    '   siteDetails
    '   processNarrative
    '   upgradeTypeText
    '   peerCommunicationsText
    '   nonStandardHighLevelText
    '   nonStandardText
    '   nonStandardEquipment
    '   nonStandardInstrumentation
    '   nonstandardControlSystem
    '   nonstandardProgram
    '   nonstandardCommunication
    '   nonstandardSCADA
    '---------------

    Call Log_Line("Opening Connection to SSFS table", 1, LOG_FILE)
    query = "SELECT * FROM [Infrastructure].[DBO].[SSFS_Text_Fields] " & _
            "WHERE SITE_ID LIKE '" & siteID & "'"
    Call Log_Line("Query: " + query, 3, LOG_FILE)
    
    'Connection to Database
    Set rs = DatabaseConnection.Connection_Query(query)
            

    siteDetails = Nz(rs.Fields("SITE_DETAILS").value, "")
    replaceDict.Add "siteDetails", siteDetails
    Call Log_Line("siteDetails: " + siteDetails, 3, LOG_FILE)
    
    processNarrative = Nz(rs.Fields("SS_PROCESS_NARRATIVE").value, "")
    replaceDict.Add "processNarrative", processNarrative
    Call Log_Line("processNarrative: " + processNarrative, 3, LOG_FILE)
    
    sitePurpose = Nz(rs.Fields("SITE_PURPOSE").value, "")
    replaceDict.Add "sitePurpose", sitePurpose
    Call Log_Line("processNarrative: " + processNarrative, 3, LOG_FILE)
    
    networkDetails = Nz(rs.Fields("NETWORK_DETAILS").value, "")
    replaceDict.Add "networkDetails", networkDetails
    Call Log_Line("networkDetails: " + networkDetails, 3, LOG_FILE)
    
    UpgradeTypeText = Nz(rs.Fields("UPGRADE_TYPE").value, "")
    replaceDict.Add "UpgradeTypeText", UpgradeTypeText
    Call Log_Line("upgradeTypeText: " + UpgradeTypeText, 3, LOG_FILE)
    
    peerCommunicationsText = Nz(rs.Fields("PEER_COMMS_TEXT").value, "")
    replaceDict.Add "peerCommunicationsText", peerCommunicationsText
    Call Log_Line("peerCommunicationsText: " + peerCommunicationsText, 3, LOG_FILE)
    
    nonStandardHighLevelText = Nz(rs.Fields("NS_HIGH_LEVEL_DESIGN").value, "")
    replaceDict.Add "nonStandardHighLevelText", nonStandardHighLevelText
    Call Log_Line("nonStandardHighLevelText: " + nonStandardHighLevelText, 3, LOG_FILE)
    
    nonStandardText = Nz(rs.Fields("NS_CONT_FUNC_TEXT").value, "")
    replaceDict.Add "nonStandardText", nonStandardText
    Call Log_Line("nonStandardText: " + nonStandardText, 3, LOG_FILE)
    
    nonStandardEquipment = Nz(rs.Fields("NS_EQUIP_TEXT").value, "")
    replaceDict.Add "nonStandardEquipment", nonStandardEquipment
    Call Log_Line("nonStandardEquipment: " + nonStandardEquipment, 3, LOG_FILE)
    
    nonStandardInstrumentation = Nz(rs.Fields("NS_INST_TEXT").value, "")
    replaceDict.Add "nonStandardInstrumentation", nonStandardInstrumentation
    Call Log_Line("nonStandardInstrumentation: " + nonStandardInstrumentation, 3, LOG_FILE)
    
    nonstandardControlSystem = Nz(rs.Fields("NS_CONT_SYS_HW").value, "")
    replaceDict.Add "nonstandardControlSystem", nonstandardControlSystem
    Call Log_Line("nonstandardControlSystem: " + nonstandardControlSystem, 3, LOG_FILE)
    
    nonstandardProgram = Nz(rs.Fields("NS_RTU_PROG").value, "")
    replaceDict.Add "nonstandardProgram", nonstandardProgram
    Call Log_Line("nonstandardProgram: " + nonstandardProgram, 3, LOG_FILE)
    
    nonstandardCommunication = Nz(rs.Fields("NS_RTU_COMMS").value, "")
    replaceDict.Add "nonstandardCommunication", nonstandardCommunication
    Call Log_Line("nonstandardCommunication: " + nonstandardCommunication, 3, LOG_FILE)
    
    nonstandardSCADA = Nz(rs.Fields("NS_SCADA").value, "")
    replaceDict.Add "nonstandardSCADA", nonstandardSCADA
    Call Log_Line("nonstandardSCADA: " + nonstandardSCADA, 3, LOG_FILE)
    'Closing connetion to the Database and nulling the record set
    rs.Close
    Set rs = Nothing
    query = ""
    Call Log_Line("SSFS table connection closed", 1, LOG_FILE)
    Call Log_Line("-------------------------------------------------------------", 1, LOG_FILE)
    Call Log_Line_Break(LOG_FILE, LogLevel)
    
    '------------------------------------------------- Document References -----------------------------------------------------
    'Selecting from Document_references table
    ' confirming variables are coming from correct key
    ' building following variables
    '   stdAssetDocumentDesigNum
    '   stdAssetDocumentDesignName
    
    Call Log_Line("Opening Connection to Document References", 1, LOG_FILE)
    
    query = "SELECT * FROM [Infrastructure].[DBO].[Document_references] " & _
            "WHERE SITE_ID LIKE '" & siteID & "' AND TITLE LIKE '%Standard functional specification%'"
            
    Call Log_Line("Query: " + query, 3, LOG_FILE)
    
    'Connection to Database
    Set rs = DatabaseConnection.Connection_Query(query)
    
    'Confirming only one row has returned, if more then one row has returned return error funtion
    rowCountChk = rs.RecordCount
    If rowCountChk = 1 Then
        Call Log_Line("PASS - One Record returned from query.", 1, LOG_FILE)
        ERROR = False
    ElseIf rowCountChk > 1 Then
       Call Log_Line("FAIL - Multiple Records Returned.", 1, LOG_FILE)
        ERROR = True
    Else
        Call Log_Line("FAIL - No Records Returned.", 1, LOG_FILE)
        ERROR = True
    End If
    
    'HAZZAH WE PASSED
    If ERROR = False Then
        Call Log_Line("Setting varibables from Doucment_refernces.", 3, LOG_FILE)
    
        stdAssetDocumentDesignNum = Nz(rs.Fields("DOCUMENT_NUMBER").value, "")
        replaceDict.Add "stdAssetDocumentDesignNum", stdAssetDocumentDesignNum
        Call Log_Line("stdAssetDocumentDesignNum: " + stdAssetDocumentDesignNum, 3, LOG_FILE)

        stdAssetDocumentDesignName = Nz(rs.Fields("TITLE").value, "")
        replaceDict.Add "stdAssetDocumentDesignName", stdAssetDocumentDesignName
        Call Log_Line("stdAssetDocumentDesignName: " + stdAssetDocumentDesignName, 3, LOG_FILE)
    
    Else
        Call Log_Line("Did not set the following values.", 3, LOG_FILE)
        Call Log_Line("stdAssetDocumentDesigNum.", 3, LOG_FILE)
        Call Log_Line("stdAssetDocumentDesigName.", 3, LOG_FILE)
        ' Reseting Error for next test
        ERROR = False
    End If
    
    'Closing Connection
    rs.Close
    Set rs = Nothing
    query = ""
    Call Log_Line("Doucment References table connection closed", 1, LOG_FILE)
    Call Log_Line("-------------------------------------------------------------", 1, LOG_FILE)
    Call Log_Line_Break(LOG_FILE, LogLevel)
    
    '-------------------------------------------------------- Site Overview Details--------------------------------------------------
    'Selecting from Site_Overview_Details table
    ' confirming variables are coming from correct key
    ' building following variables
    '   locationDescription
    
    Call Log_Line("Opening Connection to Site_Overview_Details", 1, LOG_FILE)
    
    query = "SELECT * FROM [Infrastructure].[DBO].[Site_Overview_Details] " & _
            "WHERE SITE_ID like '" & siteID & "' AND [DESCRIPTION] like 'Location_Description'"
            
    Call Log_Line("Query: " + query, 3, LOG_FILE)
    
    'Connection to Database
    Set rs = DatabaseConnection.Connection_Query(query)
    
    'Confirming only one row has returned, if more then one row has returned return error funtion
    rowCountChk = rs.RecordCount
    If rowCountChk = 1 Then
        Call Log_Line("PASS - One Record returned from query.", 1, LOG_FILE)
        ERROR = False
    ElseIf rowCountChk > 1 Then
       Call Log_Line("FAIL - Multiple Records Returned.", 1, LOG_FILE)
        ERROR = True
    Else
        Call Log_Line("FAIL - No Records Returned.", 1, LOG_FILE)
        ERROR = True
    End If
    
    'HAZZAH WE PASSED
    If ERROR = False Then
        Call Log_Line("Setting varibables from Site_Overveiw_Details.", 3, LOG_FILE)
    
        locationDescription = Nz(rs.Fields("value").value, "")
        replaceDict.Add "locationDescription", locationDescription
        Call Log_Line("locationDescription: " + locationDescription, 3, LOG_FILE)

    
    Else
        Call Log_Line("Did not set the following values.", 3, LOG_FILE)
        Call Log_Line("locationDescription.", 3, LOG_FILE)
        ' Reseting Error for next test
        ERROR = False
    End If
        'Closing Connection
    rs.Close
    Set rs = Nothing
    query = ""
    Call Log_Line("Site_Overveiw_Details table connection closed", 1, LOG_FILE)
    Call Log_Line("-------------------------------------------------------------", 1, LOG_FILE)
    Call Log_Line_Break(LOG_FILE, LogLevel)
    
    '----------------------------------------------------------------- Revision History Table ----------------------------------------
    'Selecting from Revision_History table
    ' confirming variables are coming from correct key
    ' building following variables
    '   revisionDocument
    '   revisionClass
    
    Call Log_Line("Opening Connection to Revsion_History table", 1, LOG_FILE)
    
    query = "SELECT TOP (1) [SITE_ID] " & _
                ",[REVISION] " & _
                ",[PROJECT] " & _
                ",CONVERT(date,[DATE],103) As DateOrder " & _
                ",[date] " & _
                ",[PURPOSE] " & _
                ",[COMPANY] " & _
                ",[PREPARED] " & _
                ",[REVIEWED] " & _
                ",[APPROVED] " & _
                ",[DOCUMENT_NAME] " & _
                ",[ID] " & _
            "FROM [Infrastructure].[dbo].[REVISION_HISTORY] " & _
            "WHERE SITE_ID like '" & siteID & "' " & _
            "ORDER BY DateOrder DESC "
        
    Call Log_Line("Query: " + query, 3, LOG_FILE)
    
    'Connection to Database
    Set rs = DatabaseConnection.Connection_Query(query)
    
    'Confirming only one row has returned, if more then one row has returned return error funtion
    rowCountChk = rs.RecordCount
    If rowCountChk = 1 Then
        Call Log_Line("PASS - One Record returned from query.", 1, LOG_FILE)
        ERROR = False
    ElseIf rowCountChk > 1 Then
       Call Log_Line("FAIL - Multiple Records Returned.", 1, LOG_FILE)
        ERROR = True
    Else
        Call Log_Line("FAIL - No Records Returned.", 1, LOG_FILE)
        ERROR = True
    End If
    
    'HAZZAH WE PASSED
    If ERROR = False Then
        Call Log_Line("Setting varibables from Revision_History Table.", 3, LOG_FILE)
    
        revisionDocument = Nz(rs.Fields("document_name").value, "")
        replaceDict.Add "revisionDocument", revisionDocument
        Call Log_Line("revisionDocument: " + revisionDocument, 3, LOG_FILE)

        revisionClass = Nz(rs.Fields("revision").value, "")
        replaceDict.Add "revisionClass", revisionClass
        Call Log_Line("revisionClass: " + revisionClass, 3, LOG_FILE)
    
    Else
        Call Log_Line("Did not set the following values.", 3, LOG_FILE)
        Call Log_Line("revisionDocument.", 3, LOG_FILE)
        Call Log_Line("revisionClass.", 3, LOG_FILE)
        ' Reseting Error for next test
        ERROR = False
    End If
    
    'Closing Connection
    rs.Close
    Set rs = Nothing
    query = ""
    Call Log_Line("Revision_History table connection closed", 1, LOG_FILE)
    Call Log_Line("-------------------------------------------------------------", 1, LOG_FILE)
    Call Log_Line_Break(LOG_FILE, LogLevel)
    
    '--------------------------------------------------------REVISION HISTROY--------------------------------------------------------
    'Selecting from Communications configuration
    ' confirming variables are coming from correct key
    ' building following variables
    '   DNP3_Master
    '   DNP3_Slave
    
    Call Log_Line("Opening Connection to Communications Configuration table", 1, LOG_FILE)
    
    query = "SELECT " & _
                "[DNP3_MASTER], " & _
                "[DNP3_SLAVE] " & _
            "FROM " & _
                "[Infrastructure].[dbo].[COMMUNICATIONS_CONFIGURATION] " & _
            "WHERE " & _
                "[SITE_ID] LIKE '" & siteID & "' "

    Call Log_Line("Query: " + query, 3, LOG_FILE)
    
    'Connection to Database
    Set rs = DatabaseConnection.Connection_Query(query)
    
        Call Log_Line("Setting varibables from Communications Configuration Table.", 3, LOG_FILE)
    
        DNP3MasterAddr = Nz(rs.Fields("DNP3_MASTER").value, "")
        replaceDict.Add "DNP3MasterAddr", DNP3MasterAddr
        Call Log_Line("DNP3MasterAddr: " + DNP3MasterAddr, 3, LOG_FILE)

        DNP3SlaveAddr = Nz(rs.Fields("DNP3_SLAVE").value, "")
        replaceDict.Add "DNP3SlaveAddr", DNP3SlaveAddr
        Call Log_Line("DNP3SlaveAddr: " + DNP3SlaveAddr, 3, LOG_FILE)
    
    'Closing Connection
    rs.Close
    Set rs = Nothing
    query = ""
    Call Log_Line("Communications Configuration table connection closed", 1, LOG_FILE)
    Call Log_Line("-------------------------------------------------------------", 1, LOG_FILE)
    Call Log_Line_Break(LOG_FILE, LogLevel)
    
    Call Log_Line("All variables Gathered and Set", 1, LOG_FILE)
    Call Log_Line("Moving onto writing all variables", 1, LOG_FILE)
    Call Log_Line("******************************************************************************************************************", 1, LOG_FILE)

    '*******************************************************************************************************************************************
    ' Populating varibales in FS
    ' Printing all Keys + values to Debug
    scope = "Updating Document placeholders"
    Call UI_update(scope, process)
    Call Log_Line("These Keys will be Replaced", 3, LOG_FILE)
    'Cycling through keys and values
    For Each key In replaceDict.Keys
        DoEvents
        Call Log_Line("Key:" + key + "Value:" + replaceDict(key), 3, LOG_FILE)
    Next key
    
    'Looping through replaceDict to update all values
    
    For Each key In replaceDict.Keys
        DoEvents
        tagFindString = "$" + key
        tagReplaceString = replaceDict(key)
        If (Len(tagReplaceString) < 1) Then
            Call Select_Current_Paragraph(wordapp, wordDoc, tagFindString, "word", "delete")
        Else
            Call Log_Line("finding: " + tagFindString + " Updating: " + tagReplaceString, 1, LOG_FILE)
            Call DocumentTools.Variable_Replacement(wordapp, wordDoc, tagFindString, tagReplaceString, process, LOG_FILE)
            Call Header_Footer(wordapp, wordDoc, tagFindString, tagReplaceString, LOG_FILE)
        End If
    Next key
          
    '************************************************************************************************************************************************
    ' Building Options Table
    
    scope = "BUILDING Site Options Table: "
    Call UI_update(scope, process)
    Call Site_Options_Table(wordapp, wordDoc, siteID, assetAbbreviation, scope, LOG_FILE)
    '-------------------------->  Asset Specific Equipment Generation   <-------------------------------
    scope = "Building Pump Information |Tables & sections |"
    Call UI_update(scope, process)
    Call Asset_Pump_Instantiation(wordapp, wordDoc, siteID, assetAbbreviation, scope, process, LOG_FILE)

    If (assetAbbreviation = "SP") Then
    
        scope = "Inserting SP Specific Tables: "
        Call UI_update(scope, process)
        Call SP_Specific_Tables(wordapp, wordDoc, siteID, scope, LOG_FILE)
    
    'ElseIf (assetAbbreviation = "WR") Then
    
    Else
        Call UI_update("Removing SP tables", process)
        Call Log_Line("Section Removed as not SP - Deleting Section", 3, LOG_FILE)
        tagRemoveSection = "$$WET_WELL_LEVEL_VOLUME_LOOKUP_TABLE$$"
        Call Delete_Section(wordapp, wordDoc, tagRemoveSection)
        Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
        tagRemoveSection = "$$WET_WELL_LEVELS_TABLE$$"
        Call Delete_Section(wordapp, wordDoc, tagRemoveSection)
        Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
    End If
    
    '-------------------------------------------------------------- Assumptions Table ---------------------------------------------------------------------
    ' This builds the assumptions table
    ' Tablename: Assumptions
    scope = "Building Assumptions Table"
    Call UI_update(scope, process)
    
    tableTitle = "Assumptions"
    scope = scope + tableTitle
  
    query = "SELECT " & _
                "[Order_No], " & _
                "[Definition] " & _
            "FROM " & _
                "Assumptions " & _
            "WHERE " & _
                "[SITE_ID] LIKE '" & siteID & "' " & _
            "ORDER BY " & _
                "[Order_No] ASC"
        
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line(scope + " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
  
    '--------------------------------------------------------------- Standard Tables ----------------------------------------------------------------------
    '---- This section adds the standard tables for the equipment
    '   4.6.1   General Site Control values
    '   4.6.2   Station Duty Setpoints
    '   2.2     Overall Site Information
    scope = "Inserting Standard Equipment Tables: "

    tableTitle = "General Site Control Values"
    scope = scope + tableTitle
    Call UI_update(scope, process)
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
        "A.FS_Table LIKE '" & tableTitle & "' AND ASSET_TYPE like '" & compositeType & "' " & _
    "ORDER BY " & _
        "A.[ORDER] "

    'Connetion to Database and returning recordset
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line(scope + " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    If (rs.RecordCount > 0) Then
        ' Building Table
        Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
    Else
        'No records
        tag = "$$GENERAL_SITE_CONTROL$$"
        Call UI_update(scope + " No records", process)
        Call Log_Line(scope + " No records", 3, LOG_FILE)
        Call removeTagSection(wordapp, wordDoc, tag)
    End If
    
    'Closing Record set connetion and clearing recordset
    rs.Close
    Set rs = Nothing
    
    scope = "Inserting Standard Equipment Tables: "
    tableTitle = "Station Duty Setpoints"
    scope = scope + tableTitle
    Call UI_update(scope, process)
    query = "SELECT " & _
                "A.[Tag_Description], " & _
                "A.[Tag], " & _
                "COALESCE(ISNULL(ISNULL(B.SITE_SPECIFIC,B.Default_value), C.[VALUE]),'') AS [VALUE], " & _
                "ISNULL(MIN_SETPOINT,'') AS Min_Setpoint, " & _
                "ISNULL(Max_SETPOINT,'') AS Max_Setpoint, " & _
                "COALESCE(ISNULL(B.EU, C.UNITS),'') AS [UNITS] " & _
            "FROM " & _
                "Look_Up_Table_FS AS A " & _
                "LEFT JOIN SITE_SPECIFIC_TAG_DATA AS B ON A.Tag = CONCAT(B.Object_Group, B.Tag_Attribute) " & _
                "AND B.SITE_ID LIKE '" & siteID & "' " & _
                "LEFT JOIN Look_Up_Table_FS_Values AS C ON A.ID = C.TAG_KEY " & _
                "AND C.SITE_ID LIKE '" & siteID & "' " & _
            "WHERE " & _
                "A.FS_Table LIKE '" & tableTitle & "' AND Asset_Type like '" & compositeType & "' "
    
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line(scope + " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    If (rs.RecordCount > 0) Then
        ' Building Table
        Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
    Else
        'No records
        tag = "$$STATION_DUTY_SETPOINTS$$"
        Call UI_update(scope + " No records", process)
        Call Log_Line(scope + " No records", 3, LOG_FILE)
        Call removeTagSection(wordapp, wordDoc, tag)
    End If
    
    
    
    'Inserting Overall Site Information Table
    scope = "Inserting Standard Equipment Tables: "
    
    tableTitle = "Overall Site Information"
    scope = scope + tableTitle
    Call UI_update(scope, process)

    query = "SELECT " & _
            "Tag_Description, " & _
            "ISNULL([VALUE],'') AS [VALUE], " & _
            "ISNULL([UNITS],'') AS [UNITS] " & _
        "FROM " & _
            "Look_Up_Table_FS AS TABLE_NAME " & _
            "JOIN Look_Up_Table_FS_Values AS B ON TABLE_NAme.ID = B.TAG_KEY " & _
            "AND SITE_ID = '" & siteID & "' " & _
        "WHERE " & _
            "TABLE_NAme.Asset_Type LIKE '" & compositeType & "' " & _
            "AND TABLE_NAme.FS_Table LIKE 'Overall Site Information' " & _
            "AND B.[VALUE] IS NOT NULL " & _
        "ORDER " & _
            "BY [ORDER] ASC "
                
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line(scope + " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    If (rs.RecordCount > 0) Then
        ' Building Table
        Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
    Else
        'No records
        Call UI_update(scope + " No records", process)
        Call Log_Line(scope + " No records", 3, LOG_FILE)
        Call Remove_Previous_Paragraph(wordapp, wordDoc, tableTitle, 2, "table")
        Call Table_Deletion(wordapp, wordDoc, tableTitle, LOG_FILE)
    End If
    
    '--------------------------- INSERTING STANDARD EQUIPMNET TABLES------------------------------
    ' This section queries the database In the following order to determine if equipment exists.
    ' In the Event that the equipment exists the table section is inserted with the parameters
    ' and sections table. If no equipment is present then no section is inserted.
    ' The Equipment are queried in the following order
    '
    '--------------------- Composite switch -------------------
    scope = "Inserting Equipment: "
    ' Conditioning for parameters
    previousEquipmentType = ""
    '- Form list of Equipment for standard site
    Dim counter As Integer
    counter = 1
    
    query = "SELECT " & _
                "B.*, " & _
                "A.Section_Total_Table " & _
            "FROM " & _
                "Look_Up_Table_STANDARD_EQUIPMENT AS B JOIN ( SELECT ASSET_TYPE, EQUIPMENT, COUNT(*) Section_Total_Table, Section_heading " & _
            "FROM " & _
                        "Look_Up_Table_STANDARD_EQUIPMENT " & _
                    "WHERE " & _
                        "ASSET_TYPE = '" & compositeType & "' " & _
                    "GROUP BY " & _
                        "ASSET_TYPE, " & _
                        "EQUIPMENT, " & _
                        "Section_heading " & _
                ") AS A ON A.ASSET_TYPE = B.ASSET_TYPE " & _
                "AND A.EQUIPMENT = B.EQUIPMENT " & _
                "AND A.Section_heading = B.Section_heading " & _
            "WHERE " & _
                "B.ASSET_TYPE = '" & compositeType & "' " & _
            "ORDER BY " & _
                    "B.[ORDER] ASC "
            
     Set rs = DatabaseConnection.Connection_Query(query)
     
     queryRows = rs.RecordCount
     If (queryRows > 0) Then
        rs.MoveFirst
        Do Until rs.EOF = True
            ' Buidling Variables
            sectionHeading = rs.Fields("Section_heading").value
            Sectioncounter = rs.Fields("Section_Total_Table").value
            equipment = rs.Fields("EQUIPMENT").value
            scope = "Inserting Equipment: " + equipment
            
            'Used to work out if text needs to be inserted
            currentEquipmentType = Left(equipment, 3)
            
            Call Log_Line(scope + " Checking if ", 1, LOG_FILE)
            
            'Checking to see if the equipment is present
            If (InStr(equipment, "LIT")) Then
            Debug.Print ("Here")
            End If
            
            
            exists = Asset_Equipment_check(wordapp, wordDoc, equipment, siteID, scope, LOG_FILE)
            If (exists) Then
                                
                Call Log_Line(scope + " Adding ", 2, LOG_FILE)
                ' The equipment exists
                tableTitle = rs.Fields("FS_TABLE_NAME").value
                Call UI_update(scope + " " + tableTitle, process)
                'All of this code inserts Equipment in section 4
                
                If (counter = 1) Then
                    'Inserting Heading
                    sectionHeading = rs.Fields("Section_heading").value
                    tag = "$$STANDARD_EQUIPMENT$$"
                    Call insertAboveTag(wordapp, wordDoc, tag, "section")
                    Call Insert_Section_Heading(wordapp, wordDoc, sectionHeading, LOG_FILE)
                    counter = counter + 1
                ElseIf (counter < Sectioncounter) Then
                    counter = counter + 1
                ElseIf (counter = Sectioncounter) Then
                    counter = 1
                Else
                    ' Do nothing I could make the top one an else but then I lose 1`ity on conditioning
                End If
                
                'Inserting table and caption
                
                'Insert caption, table and instantiate table.
                ' Maybe add a field for table caption (??)
                Caption = ": " + rs.Fields("FS_Table_Name")
                tableTitle = rs.Fields("FS_Table_Name")
                
                ' Setting the copy template
                If (InStr(tableTitle, "Parameters")) Then
                    tableTemplate = "Template Table Equipment Parameters"
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
                                        "AND A.ASSET_TYPE LIKE '" & compositeType & "' " & _
                                    "ORDER BY " & _
                                        "A.[ORDER] ASC "
                    tableTypeParameter = True

                ElseIf (InStr(tableTitle, "Setpoints")) Then
                    tableTemplate = "Template Table Equipment Setpoints"
                                query = "SELECT " & _
                                        "A.[Tag_Description], " & _
                                        "A.[Tag], " & _
                                        "COALESCE(ISNULL(ISNULL(B.SITE_SPECIFIC,B.Default_value), C.[VALUE]),'') AS [VALUE], " & _
                                        "ISNULL(MIN_SETPOINT,'') AS Min_Setpoint, " & _
                                        "ISNULL(Max_SETPOINT,'') AS Max_Setpoint, " & _
                                        "COALESCE(ISNULL(B.EU, C.UNITS),'') AS [UNITS] " & _
                                    "FROM " & _
                                        "Look_Up_Table_FS AS A " & _
                                        "LEFT JOIN SITE_SPECIFIC_TAG_DATA AS B ON A.Tag = CONCAT(B.Object_Group, B.Tag_Attribute) " & _
                                        "AND B.SITE_ID LIKE '" & siteID & "' " & _
                                        "LEFT JOIN Look_Up_Table_FS_Values AS C ON A.ID = C.TAG_KEY " & _
                                        "AND C.SITE_ID LIKE '" & siteID & "' " & _
                                    "WHERE " & _
                                        "A.FS_Table LIKE '" & tableTitle & "' " & _
                                        " AND A.ASSET_TYPE LIKE '" & compositeType & "' " & _
                                    "ORDER BY " & _
                                        "A.[ORDER] ASC "
                    tableTypeParameter = False
                Else
                    tableTemplate = ""
                End If

                
                Call Insert_Table(wordapp, wordDoc, tableTitle, tableTemplate, query, Caption, scope, LOG_FILE)
                
                
                ' Inserting the paragraph under the table for the requested equipment
                If (tableTypeParameter) Then
                    ' Parameter Table inserted
                    If (InStr(equipment, "PMP")) Then
                        paragraph = "Note: All Current and Power limits are based on electrical drawing and to be set onsite during commissioning. " & _
                                   "Alarms to be +- 30% of normal range. Testing at full speed should be conducted to determine maximum power/current. " & _
                                   "Site specific Functional Specification shall be updated by commissioning engineer. Please refer to Section 4.3." & counter & "."
                        
                        Call insertAboveTag(wordapp, wordDoc, "$$STANDARD_EQUIPMENT$$", "")
                        Call Insert_Paragraph(wordapp, wordDoc, paragraph, "before")
                        
                        wordapp.Selection.MoveRight Unit:=wdCharacter, Count:=1
                        wordapp.Selection.MoveUp Unit:=wdParagraph, Count:=1
                        wordapp.Selection.TypeBackspace
                        
                        paragraphCounter = paragraphCounter + 1
                        
                    ElseIf (InStr(equipment, "FIT")) Then
                        ' Room for FIT paragraph
                        paragraph = "Note: All flow alarm limits to be set onsite during commissioning. Alarms to be -20% of normal range. " & _
                                   "Site Specification Functional shall be updated by commissioning engineer."
                        Call insertAboveTag(wordapp, wordDoc, "$$STANDARD_EQUIPMENT$$", "")
                        Call Insert_Paragraph(wordapp, wordDoc, paragraph, "before")
                        
                        wordapp.Selection.MoveRight Unit:=wdCharacter, Count:=1
                        wordapp.Selection.MoveUp Unit:=wdParagraph, Count:=1
                        wordapp.Selection.TypeBackspace
                    ElseIf (InStr(equipment, "PIT")) Then
                        ' Room for PIT paragraph
                        paragraph = "Note: All pressure alarm limits to be set onsite during commissioning. Alarms to be -20% of normal range. " & _
                                   "Site Specification Functional shall be updated by commissioning engineer."
                                   
                        Call insertAboveTag(wordapp, wordDoc, "$$STANDARD_EQUIPMENT$$", "")
                        Call Insert_Paragraph(wordapp, wordDoc, paragraph, "before")
                        
                        wordapp.Selection.MoveRight Unit:=wdCharacter, Count:=1
                        wordapp.Selection.MoveUp Unit:=wdParagraph, Count:=1
                        wordapp.Selection.TypeBackspace
                    
                    Else
                        ' Setting paragraph to null
                        paragraph = ""
                    End If

                End If
                
                ' Reseting Counter this need to be done for when the equipment switches from PMP to FIT
                If (previousEquipmentType = currentEquipmentType) Then
                    ' Do nothing
                Else
                    ' Reset counter
                    paragraphCounter = 1
                End If
                
                
                
                
                'Call the function to build the equipment based on table
                ' Define
                '   - Section Heading
                '   - caption
                '   - Table Heading
                '   - Insert all
                '   - Update Table
                
            Else
                Call Log_Line(scope + " Not present - Not adding ", 2, LOG_FILE)
            End If
                
            
            rs.MoveNext
        Loop
     Else
     
     End If
    '---------------------------------------------------- NON-Standard Tags ---------------------------------------------------------------
    ' This section Queires the Database to determine if their is any Non-standard tags
    ' In the event that the asset has are non-standard Equip/Instrumentation then the following tables are updated.
    ' The three section which are filled out include NS-Control Functions, NS Equipment, NS-Instrumentation
    '   IO                              -
    '   Parameters and setpoints        -
    '   Calculations and Statistics     -
    '   SCADA Points                    -
    
    
    '---------- NS-Control Functions-------------
    ' Not Yet Defined
    Call UI_update("Inserting Nonstandard Equipment Tables", process)
    Call NonStandardTables(wordapp, wordDoc, siteID, "NSC", "Control", "Control", LOG_FILE)
    
    '---------- NS-Equipment --------------------
    Call UI_update("Inserting Nonstandard Equipment Tables", process)
    Call NonStandardTables(wordapp, wordDoc, siteID, "NSE", "Equipment", "Equipment", LOG_FILE)
    
    
    '--------- NS-Instruments ------------------
    Call UI_update("Inserting Nonstandard Instrument Tables", process)
    Call NonStandardTables(wordapp, wordDoc, siteID, "NSI", "Instrumentation", "Instrumentation", LOG_FILE)
    
    '-------- NS-RTU Communication -------------
    scope = "Builiding NS-RTU Communication: "
    Call UI_update("Inserting Nonstandard Instrument Tables", process)
    Call NonStandard_RTU_Communication(wordapp, wordDoc, siteID, scope, LOG_FILE)

    

    '------------------------------------------ GENERATOR ---------------------------------------------
    'Need to write logic for this
        generatorExists = True
    scope = "Building Generator table values"
    Call UI_update(scope, process)
    Call Generator_Details(wordapp, wordDoc, siteID, assetAbbreviation, scope, LOG_FILE, process)

    If (generatorExists) Then
    Call UI_update("Removing Generator placeholder tags", process)
        Call Log_Line("Tag Removed: Generator exists - Deleting Template Tags", 3, LOG_FILE)
        
        tagRemoveString = "$$GENERATOR_TABLE:START$$"
        Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
        'Call Variable_Remove(wordapp, wordDoc, tagRemoveString, LOG_FILE)
        Call Log_Line("Tag Removed : " & tagRemoveString, 3, LOG_FILE)
    
        tagRemoveString = "$$GENERATOR_TABLE:END$$"
        Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
        'Call Variable_Remove(wordapp, wordDoc, tagRemoveString, LOG_FILE)
        Call Log_Line("Tag Removed : " & tagRemoveString, 3, LOG_FILE)
        
    Else
        Call Log_Line("Tag Removed: Does not Generator exists - Deleting Section", 3, LOG_FILE)
        tagRemoveSection = "$$GENERATOR_TABLE$$"
        Call Delete_Section(wordapp, wordDoc, tagRemoveSection)
        Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- Building Flat Tables ------------------------------------------
    '   - Revision Control
    '   - References and Inputs
    '   - Options - Todo
    '   - Network Details
    '   - Site Pump Overview (Not Added, resolved through param insertion)
    '   - Site Network
    '   - Pump Instantiation (resolved in equipment section)
    '   - Networking Address (IP)
    '   - Modbus devices
    '   - Site Equipment (Resolved In equipment section)
    '   - Drawing List
    '   - Physcial IO
    
    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- REVISION CONTROL ----------------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   Revision Control Table
    '   FS TableName: Revision control
    '   DB Table: Revision Histroy
    '------------------------------------------------------------------------------------------------------------------
    
    tableTitle = "Revision Control"
    Call UI_update("Adding " & tableTitle & " Table", process)
    'Read Revision Control Data
    query = "select ISNULL([REVISION],'') AS [REVISION], ISNULL([PROJECT],'') AS [PROJECT], CONVERT(date,[DATE],103) as BuildDate," & _
            "       ISNULL([PURPOSE],'') AS [PURPOSE], ISNULL([COMPANY],'') AS [COMPANY], ISNULL([PREPARED],'') AS [PREPARED], ISNULL([REVIEWED],'') AS [REVIEWED], ISNULL([APPROVED],'') AS [APPROVED]" & _
            "from [Infrastructure].[dbo].[REVISION_HISTORY]" & _
            "where site_id = '" & siteID & "'" & _
            "order by BuildDate asc"

    'Connetion to Database and returning recordset
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
    
    
    'Closing Record set connetion and clearing recordset
    rs.Close
    Set rs = Nothing
    query = ""
    
    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- REFERENCES AND INPUTS -----------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   References And Inputs Table
    '   FS TableName: References And Inputs
    '   DB Table: Document_References
    '------------------------------------------------------------------------------------------------------------------
     tableTitle = "References And Inputs"
    Call UI_update("Adding " & tableTitle & " Table", process)
    'Read References And Inputs Data
    query = "SELECT " & _
                "ISNULL([INPUT],'') AS [INPUT], " & _
                "ISNULL([TITLE],'') AS [TITLE], " & _
                "ISNULL([DOCUMENT_NUMBER], '') AS [DOCUMENT_NUMBER] " & _
            "FROM " & _
                "[Document_references] " & _
            "WHERE " & _
                "SITE_ID LIKE '" & siteID & "' " & _
            "ORDER BY " & _
                "[INPUT] ASC "

    'Connetion to Database and returning recordset
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
    
    'Closing Record set connetion and clearing recordset
    rs.Close
    Set rs = Nothing
    query = ""
    
    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- OPTIONS -------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   Site Options Table
    '   FS TableName: Site Options
    '   DB Table:
    '------------------------------------------------------------------------------------------------------------------

    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- NETWORK DETAILS -----------------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   Network Details Table
    '   FS TableName: Network Details
    '   DB Table: Document_References
    '------------------------------------------------------------------------------------------------------------------
    'Query Select, if SP or other site
    tableTitle = "Network Details"
    Call UI_update("Adding " & tableTitle & " Table", process)
    If (assetAbbreviation = "SP") Then
    Call Log_Line("Update Table: Asset is a SP. Setting Query to SP", 3, LOG_FILE)
        query = "SELECT " & _
                    "Replace([DESCRIPTION], '_', ' ') AS [DESCRIPTION], " & _
                    "ISNULL([VALUE],'') AS [VALUE], " & _
                    "ISNULL([UNITS],'') AS [UNITS] " & _
                "FROM " & _
                    "[Infrastructure].[dbo].[SITE_OVERVIEW_DETAILS] " & _
                "WHERE " & _
                    "site_id = '" & siteID & "' " & _
                    "AND DESCRIPTION IN " & _
                    "('Sewer_Scheme', " & _
                    "'Macro_Catchment', " & _
                    "'ADWF', " & _
                    "'Well Fill Time', " & _
                    "'Well Empty Time', " & _
                    "'Pump Cycle Time', " & _
                    "'Dry Weather Hours')" & _
                    "AND [VALUE] IS NOT NULL " & _
                "ORDER BY [ORDER] ASC"
    ElseIf (assetAbbreviation = "WR") Then
                query = "SELECT " & _
                    "Replace([DESCRIPTION], '_', ' ') AS [DESCRIPTION], " & _
                    "ISNULL([VALUE],'') AS [VALUE], " & _
                    "ISNULL([UNITS],'') AS [UNITS] " & _
                "FROM " & _
                    "[Infrastructure].[dbo].[SITE_OVERVIEW_DETAILS] " & _
                "WHERE " & _
                    "SITE_ID like '" & siteID & "' " & _
                    "AND [VALUE] IS NOT NULL "
'                    "AND DESCRIPTION IN " & _
'                        "('Capacity', " & _
'                        "'Flowmeter', " & _
'                        "'Valve', " & _
'                        "'Reservoir_1_Description', " & _
'                        "'Reservoir_2_Description', " & _
'                        "'Reservoir_1_Options', " & _
'                        "'Reservoir_2_Options', " & _
'                        "'Supply_Zone', " & _
'                        "'Location_Description', " & _
'                        "'Peer_Site1', " & _
'                        "'Peer_Site2', " & _
'                        "'Radio', " & _
'                        "'Suction_Zone', " & _
'                        "'Discharge_Zone', " & _
'                        "'Signals') "
            Call Log_Line("Update Table: Asset NOT SP. Setting Query to Normal", 3, LOG_FILE)
    Else
            query = "SELECT " & _
                    "Replace([DESCRIPTION], '_', ' ') AS [DESCRIPTION], " & _
                    "ISNULL([VALUE],'') AS [VALUE], " & _
                    "ISNULL([UNITS],'') AS [UNITS] " & _
                "FROM " & _
                    "[Infrastructure].[dbo].[SITE_OVERVIEW_DETAILS] " & _
                "WHERE " & _
                    "SITE_ID like '" & siteID & "' " & _
                    "AND [VALUE] IS NOT NULL "
            Call Log_Line("Update Table: Asset NOT SP. Setting Query to Normal", 3, LOG_FILE)
    End If
    

    'Connetion to Database and returning recordset
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
    
    'Closing Record set connetion and clearing recordset
    rs.Close
    Set rs = Nothing
    query = ""
    
    
    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- Site Pump Overview --------------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   Site Pump Overview Table
    '   FS TableName: Site Pump Overview
    '   DB Table: Document_References
    '------------------------------------------------------------------------------------------------------------------
    
    
    
    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- Site Network -----------------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   Site Network Table
    '   FS TableName: Site Network
    '   DB Table: Network_Table
    '   If the Asset Type is SP then the table is not populated and deleted instead
    '------------------------------------------------------------------------------------------------------------------
    tableTitle = "Site Network"
   '-------------------------- NOT TESTED -----------------------------------------------------------------------------
    If (assetAbbreviation <> "SP") Then
    Call UI_update("Adding " & tableTitle & " Table", process)
    upQuery = "SELECT " & _
                    "ISNULL([UPSTREAM_ID],'') AS [UPSTREAM_ID] " & _
                    ", ISNULL([DOWNSTREAM_ID],'') AS [DOWNSTREAM_ID] " & _
                "FROM " & _
                    "[Infrastructure].[dbo].[NETWORK_TABLE] " & _
                "WHERE " & _
                    "[DOWNSTREAM_ID] LIKE '" & siteID & "'"
                    
    downQuery = "SELECT " & _
                    "ISNULL([UPSTREAM_ID],'') AS [UPSTREAM_ID] " & _
                    ", ISNULL([DOWNSTREAM_ID],'') AS [DOWNSTREAM_ID] " & _
                "FROM " & _
                    "[Infrastructure].[dbo].[NETWORK_TABLE] " & _
                "WHERE " & _
                    "[UPSTREAM_ID] LIKE '" & siteID & "'"
    
    
    Set rsUp = DatabaseConnection.Connection_Query(upQuery)
    Set rsDown = DatabaseConnection.Connection_Query(downQuery)
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation_UpDown(wordapp, wordDoc, siteID, rsUp, rsDown, tableTitle, LOG_FILE)
    
    'Closing Record set connetion and clearing recordset
    rsUp.Close
    rsDown.Close
    Set rsUp = Nothing
    Set rsDown = Nothing
    upQuery = ""
    downQuery = ""
    Else
        'Table Deletion
        Call Remove_Previous_Paragraph(wordapp, wordDoc, tableTitle, 2, "table")
        Call Table_Deletion(wordapp, wordDoc, tableTitle, LOG_FILE)

    
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- Network IP Addressing (easy) ----------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   Network IP Addressing Table
    '   FS TableName: Network IP Addressing
    '   DB Table: NETWORK_PARAMETERS
    '------------------------------------------------------------------------------------------------------------------
    tableTitle = "IP Network Information"
    Call UI_update("Adding " & tableTitle & " Table", process)
    query = "SELECT " & _
                "ISNULL([Device],'') AS [Device], " & _
                "ISNULL([Equip_ID],'') AS [Equip_ID], " & _
                "ISNULL([Port],'') AS [Port], " & _
                "ISNULL([Description],'') AS [Description], " & _
                "ISNULL([Address],'') AS [Address] " & _
            "FROM " & _
                "[INFRASTRUCTURE].[DBO].[NETWORK_PARAMETERS] " & _
            "WHERE " & _
                "SITE_ID LIKE '" & siteID & "' " & _
            "ORDER BY " & _
                "[Site_Index] "
    
    'Connetion to Database and returning recordset
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
    
    'Closing Record set connetion and clearing recordset
    rs.Close
    Set rs = Nothing
    query = ""
    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- Modbus devices (easy) -----------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   Modbus Devices Table
    '   FS TableName: Modbus devices
    '   DB Table: MODBUS_DEVICES
    '------------------------------------------------------------------------------------------------------------------
    tableTitle = "Modbus Devices"
    Call UI_update("Adding " & tableTitle & " Table", process)
    query = "SELECT " & _
                "ISNULL([INTERFACE],'') AS [INTERFACE], " & _
                "ISNULL([DEVICE],'') AS [INTERFACE] " & _
            "FROM " & _
                "MODBUS_DEVICES " & _
            "WHERE " & _
                "SITE_ID LIKE '" & siteID & "' " & _
            "ORDER BY " & _
                "DEVICE ASC "
    
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)

    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- Drawing List --------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   Drawing List Table
    '   FS TableName: Drawing List
    '   DB Table: DRAWING_TABLE
    '------------------------------------------------------------------------------------------------------------------
    tableTitle = "Drawing List"
    Call UI_update("Adding " & tableTitle & " Table", process)
    'Read References And Inputs Data
    query = "SELECT " & _
                "ISNULL([SHEET_NUMBER],'') AS [SHEET_NUMBER], " & _
                "ISNULL([DRAWING],'') AS [DRAWING], " & _
                "ISNULL([TITLE],'') AS [TITLE] " & _
            "FROM " & _
                "DRAWING_TABLE " & _
            "WHERE " & _
                "SITE_ID LIKE '" & siteID & "' " & _
            "ORDER BY " & _
                "SHEET_NUMBER ASC "

    'Connetion to Database and returning recordset
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
    
    'Closing Record set connetion and clearing recordset
    rs.Close
    Set rs = Nothing
    query = ""
    
    '------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------- Physcial IO ---------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------
    '   Physical IO Table
    '   FS TableName: Physical IO
    '   DB Table: IO_LIST
    '------------------------------------------------------------------------------------------------------------------

    tableTitle = "Physical IO"
    scope = "Inserting Physical IO"
    Call UI_update("Adding " & tableTitle & " Table", process)
    'Read References And Inputs Data
    query = "SELECT " & _
                "ISNULL([MODULE],'') AS [MODULE], " & _
                "ISNULL([CHANNEL],'') AS [CHANNEL], " & _
                "ISNULL([DATA_TYPE],'') AS [DATA_TYPE], " & _
                "ISNULL([DESCRIPTION],'') AS [DESCRIPTION], " & _
                "ISNULL([ELECTRICAL_DESCRIPTION],'') " & _
            "FROM " & _
                "[Infrastructure].[dbo].[IO_LIST] " & _
            "WHERE " & _
                "[SITE_ID] LIKE '" & siteID & "'" & _
            "ORDER BY " & _
            "CASE " & _
                "WHEN [MODULE] LIKE 'Main Board' THEN 1 " & _
                "WHEN [MODULE] LIKE 'Expansion Module 1' THEN 2 " & _
                "WHEN [MODULE] LIKE 'Expansion Module 2' THEN 3 " & _
                "ELSE 4 " & _
            "END, " & _
            "MODULE DESC ,[CHANNEL] ASC "

    'Connetion to Database and returning recordset
    Set rs = DatabaseConnection.Connection_Query(query)
    
    Call Log_Line("Update Table: " & tableTitle & " - Query returned " & queryRows & " results.", 2, LOG_FILE)
    
    Call Table_Instantiation(wordapp, wordDoc, rs, tableTitle, LOG_FILE)
    scope = "Inserting Physical IO: BackGround colour"
    Call Table_Row_Background_Colour(wordapp, wordDoc, tableTitle, "Not Used", 4, "Grey", scope)
    'Closing Record set connetion and clearing recordset
    rs.Close
    Set rs = Nothing
    query = ""
    
    '--------------------------------------------- Peer to Peer Table insert ---------------------------------
    scope = "Building DNP3 Peer to Peer Table."
    Call Peer_To_Peer_DNP3(wordapp, wordDoc, siteID, scope, LOG_FILE, process)
    '------------------------------------- Removing Template Tags --------------------------------------
    'Removing Standard Equipment Tags
    Call UI_update("Removing Place Holder Tags", process)
        Call Log_Line("Tag Removed: Equipment Tables Template - Deleting", 3, LOG_FILE)
        tagRemoveSection = "$$EQUIPMENT_TABLES$$"
        Call Delete_Section(wordapp, wordDoc, tagRemoveSection)
        Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
        
        tagRemoveString = "$$STANDARD_EQUIPMENT$$"
        Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
        'Call Variable_Remove(wordapp, wordDoc, tagRemoveString, LOG_FILE)
        Call Log_Line("Tag Removed : " & tagRemoveString, 3, LOG_FILE)
        
        Call Log_Line("Tag Removed: Equipment Tables Template - Deleting", 3, LOG_FILE)
        tagRemoveString = "$$STATION_DUTY_SETPOINTS:START$$"
        Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
        'Call Variable_Remove(wordapp, wordDoc, tagRemoveSection, LOG_FILE)
        Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
        
        Call Log_Line("Tag Removed: Equipment Tables Template - Deleting", 3, LOG_FILE)
        tagRemoveString = "$$STATION_DUTY_SETPOINTS:END$$"
        Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
        'Call Variable_Remove(wordapp, wordDoc, tagRemoveSection, LOG_FILE)
        Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
        
        Call Log_Line("Tag Removed: Equipment Tables Template - Deleting", 3, LOG_FILE)
        tagRemoveString = "$$GENERAL_SITE_CONTROL:START$$"
        Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
        'Call Variable_Remove(wordapp, wordDoc, tagRemoveSection, LOG_FILE)
        Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
        
        Call Log_Line("Tag Removed: Equipment Tables Template - Deleting", 3, LOG_FILE)
        tagRemoveString = "$$GENERAL_SITE_CONTROL:END$$"
        Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
        'Call Variable_Remove(wordapp, wordDoc, tagRemoveSection, LOG_FILE)
        Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
        
        If (assetAbbreviation = "SP") Then
            Call Log_Line("Tag Removed: Equipment Tables Template - Deleting", 3, LOG_FILE)
            tagRemoveString = "$$WET_WELL_LEVEL_VOLUME_LOOKUP_TABLE:START$$"
            Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
            'Call Variable_Remove(wordapp, wordDoc, tagRemoveSection, LOG_FILE)
            Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
            
            Call Log_Line("Tag Removed: Equipment Tables Template - Deleting", 3, LOG_FILE)
            tagRemoveString = "$$WET_WELL_LEVEL_VOLUME_LOOKUP_TABLE:END$$"
            Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
            'Call Variable_Remove(wordapp, wordDoc, tagRemoveSection, LOG_FILE)
            Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)

            Call Log_Line("Tag Removed: Equipment Tables Template - Deleting", 3, LOG_FILE)
            tagRemoveString = "$$WET_WELL_LEVELS_TABLE:START$$"
            Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
            'Call Variable_Remove(wordapp, wordDoc, tagRemoveSection, LOG_FILE)
            Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
            
            Call Log_Line("Tag Removed: Equipment Tables Template - Deleting", 3, LOG_FILE)
            tagRemoveString = "$$WET_WELL_LEVELS_TABLE:END$$"
            Call Select_Current_Paragraph(wordapp, wordDoc, tagRemoveString, "word", "delete")
            'Call Variable_Remove(wordapp, wordDoc, tagRemoveSection, LOG_FILE)
            Call Log_Line("Section Removed : " & tagRemoveSection, 3, LOG_FILE)
        Else
        
        End If
        
        Call UI_update("Removing Template Tags", process)
        Call Select_Current_Paragraph(wordapp, wordDoc, "$peerDataMap", "word", "delete")
        
        ' This function deletes trailing paragraphs
        Call UI_update("Removing Blank Paragraphs", process)
        Call Select_Section_Headings(wordapp, wordDoc, "delete", "")
        
        
        '---------------------- Legacy code --------------------------------'
'        Dim lparas As Long
'        Dim oRng As Range
'        lparas = wordDoc.Paragraphs.Count
'        Set oRng = wordDoc.Range
'
'        For i = lparas To 1 Step -1
'            oRng.Select
'            lend = lend + oRng.Paragraphs.Count
'            If Len(wordDoc.Paragraphs(i).Range.text) = 1 Then
'                wordDoc.Paragraphs(i).Range.Delete
'            End If
'        Next i
        
    
   'Finalise
        Call UI_update("Saving Document", process)
    'Update TOC/TOF/TOT
    wordDoc.SelectAllEditableRanges
    wordDoc.Fields.Update
    With wordDoc
        .TablesOfFigures(1).Update
        .TablesOfFigures(2).Update
    End With
    'Save/Close/Quit
    wordDoc.Close _
        SaveChanges:=wdPromptToSaveChanges, _
        OriginalFormat:=wdWordDocument
    wordapp.Quit
    LOG_FILE.Close
    
Done:
    Exit Function

ErrorHandler:
    On Error Resume Next

    Call Log_Line("Function Failed Exiting", 3, LOG_FILE)
    Call UI_update("Function Failed Exiting", process)
    Call UI_update("ERROR: " & Err.Description & Err.HelpContext, process)
    wordDoc.Close wdDoNotSaveChanges, wdOriginalDocumentFormat, False
    wordapp.Quit
    LOG_FILE.Close
    If (process = 0) Then
        MsgBox ("Failed to Run, restart Access [Body]")
    End If

End Function

