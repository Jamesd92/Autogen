Attribute VB_Name = "Logging"
Option Compare Database

Sub Log_Line(msg As String, LogLevel As Integer, LOGFILE As TextStream)
    'Writes out to a file for debugging
    'Inputs: msg (string), logLevel(int)
    'Output: logLine(String)
    
    'Debugging log level '1 = functionality; 2= Detail; 3 = verbose;
    
    'Logfile variables
    Dim logLine As String
    If level <= sessionLevel Then
        logLine = Format(Now(), "yyyy/mm/dd HH:MM:SS") & " : " & level & " : " & msg
        LOGFILE.WriteLine logLine
    End If
End Sub
Sub Log_Line_Break(Log As TextStream, sessionLevel As Integer)
    Log.WriteLine
End Sub


Public Function UI_update(newString As String, enabled As Integer)
    If (enabled = 0) Then
    Forms!AutoGen.Form!Status.value = Forms!AutoGen.Form!Status.value & vbCrLf & newString
    Else
    
    End If
    
End Function
