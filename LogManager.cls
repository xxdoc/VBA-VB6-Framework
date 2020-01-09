VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Framework.Common")

Option Explicit

Private Type TLogManager
    FormatedErrorMsg As String
End Type

Private this As TLogManager

Public Sub LogError(ByVal procedureName As String, ByVal errorNumber As Long, _
                    ByVal errorDescription As String, ByVal emailToAddress As String, _
                    ByVal projectName As String, ByVal networkHostName As String, _
                    ByVal displayError As Boolean, Optional ByVal adoDbConnectionErrors As Object)
        
        On Error GoTo CleanFail
        'check if ADO error first
        If IsADODBError(adoDbConnectionErrors) Then
            ManageADOErrors procedureName, errorNumber, projectName, displayError, adoDbConnectionErrors
            
        'VBA Runtime error occured
        ElseIf errorNumber > 0 Then

            ManageCommonRunTimeError procedureName, errorNumber, errorDescription, projectName, displayError
        Else
            ManageAmbigousError procedureName, errorNumber, errorDescription, projectName, displayError
        End If
        
        CreateLogfile this.FormatedErrorMsg, projectName
        
        TrySendEmail networkHostName, projectName, emailToAddress
        
        If Not displayError Then Exit Sub
        
        MsgBox "The log file for this error can be found in " & _
               "the following path: " & "C:\Users\" & _
               VBA.Environ$("USERNAME") & "\Desktop", vbInformation, "Alert: Process Complete"

        
CleanExit:
    Exit Sub
    
CleanFail:
    Resume CleanExit

End Sub

Private Sub ManageADOErrors(ByVal procedureName As String, ByVal errorNumber As Long, _
                            ByVal projectName As String, ByVal displayError As Boolean, _
                            ByVal adoDbConnectionErrors As Object)

    Dim DisplayMsg As String, ADODBErrorString As String
    
        Err.Clear
    
        ADODBErrorString = ADODBErrorsToString(adoDbConnectionErrors)
        
            If displayError Then
                DisplayMsg = "The following error(s) occured while attempting "
                DisplayMsg = DisplayMsg & "to execute the procedure named: " & procedureName & vbNewLine & vbNewLine
                DisplayMsg = DisplayMsg & ADODBErrorString & "A text file will be created, and an attempt will be "
                DisplayMsg = DisplayMsg & "made to send an email regarding the error(s) to the developer "
                DisplayMsg = DisplayMsg & "of this project. You will be notified upon completion of this process."
            
                MsgBox DisplayMsg, vbCritical, "Error: Data Provider Error"
            End If
        
    FormatErrorMsg projectName, procedureName, errorNumber, ADODBErrorString
        
        
End Sub

Private Sub ManageCommonRunTimeError(ByVal procedureName As String, ByVal errorNumber As Long, _
                                     ByVal errorDescription As String, ByVal projectName As String, _
                                     ByVal displayError As Boolean)

    Dim DisplayMsg As String

        If displayError Then
            DisplayMsg = "The following Visual Basic Runtime Error Occured while attempting to "
            DisplayMsg = DisplayMsg & "execute the procedure named: " & procedureName & vbNewLine & vbNewLine
            DisplayMsg = DisplayMsg & errorDescription & vbNewLine & vbNewLine
            DisplayMsg = DisplayMsg & "A text file will created with the details regarding this error, and "
            DisplayMsg = DisplayMsg & "an attempt will be made to send an email to the developer of this "
            DisplayMsg = DisplayMsg & "project. You will be notified upon completion of this process."
        
            MsgBox DisplayMsg, vbCritical, "Error Number: " & errorNumber
        End If
    
    FormatErrorMsg projectName, procedureName, errorNumber, errorDescription
    
End Sub


Private Sub ManageAmbigousError(ByVal procedureName As String, ByVal errorNumber As Long, _
                                ByVal errorDescription As String, ByVal projectName As String, _
                                ByVal displayError As Boolean)

    Dim DisplayMsg As String

        If displayError Then
            DisplayMsg = DisplayMsg & "An error occured while attempting to execute the procedure "
            DisplayMsg = DisplayMsg & "named: " & procedureName
            DisplayMsg = DisplayMsg & "A text file will be created with the details regarding this error, and "
            DisplayMsg = DisplayMsg & "an attempt will be made to send an email to the developer of this "
            DisplayMsg = DisplayMsg & "project. You will be notified upon completion of this process."
                
            MsgBox DisplayMsg, vbCritical, "Error: Ambiguous Error"
        End If
            
    FormatErrorMsg projectName, procedureName, errorNumber, errorDescription

End Sub
        
Private Function IsADODBError(ByRef adoDbConnectionErrors As Object) As Boolean
        If adoDbConnectionErrors Is Nothing Then Exit Function
        IsADODBError = (adoDbConnectionErrors.Count > 0)
End Function

Private Function ADODBErrorsToString(ByRef adoDbConnectionErrors As Object) As String

    Dim ADOError As ADODB.Error
    Dim i As Long
    Dim errorMsg As String
        
        For Each ADOError In adoDbConnectionErrors
            i = i + 1
            With ADOError
                errorMsg = errorMsg & "Count: " & vbTab & i & vbNewLine
                errorMsg = errorMsg & "ADO Error Number: " & vbTab & CStr(.Number) & vbNewLine
                errorMsg = errorMsg & "Description: " & vbTab & .Description & vbNewLine
                errorMsg = errorMsg & "Source: " & vbTab & .Source & vbNewLine
                errorMsg = errorMsg & "NativeError: " & vbTab & CStr(.NativeError) & vbNewLine
                errorMsg = errorMsg & "HelpFile: " & vbTab & .HelpFile & vbNewLine
                errorMsg = errorMsg & "HelpContext: " & vbTab & CStr(.HelpContext) & vbNewLine
                errorMsg = errorMsg & "SQLState: " & vbTab & .SqlState & vbNewLine & vbNewLine
            End With
        Next

    ADODBErrorsToString = vbNewLine & errorMsg

End Function

Private Sub FormatErrorMsg(ByVal projectName As String, ByVal procedureName As String, _
                           ByVal errorNumber As Long, ByVal errorDescription As String)
    
    Dim ActiveWinsession As ActiveWindowsSession
    Set ActiveWinsession = New ActiveWindowsSession

    this.FormatedErrorMsg = "Project Name: " & projectName & vbNewLine & vbNewLine & _
                            "The following Error Occured while " & ActiveWinsession.UserCommonName & _
                            " was attempting to run the Procedure Named: " & procedureName & "." & _
                            vbNewLine & vbNewLine & _
                            "Visual Basic Error: " & errorNumber & _
                            vbNewLine & vbNewLine & _
                            "Error Description: " & errorDescription & _
                            vbNewLine & vbNewLine

    With ActiveWinsession
        this.FormatedErrorMsg = this.FormatedErrorMsg & "Time of Error: " & vbTab & Now() & vbNewLine
        this.FormatedErrorMsg = this.FormatedErrorMsg & "Domain DNS Name: " & vbTab & .DomainName & vbNewLine
        this.FormatedErrorMsg = this.FormatedErrorMsg & "Site Name: " & vbTab & .CompanySiteName & vbNewLine
        this.FormatedErrorMsg = this.FormatedErrorMsg & "Computer Name: " & vbTab & .MachineName & vbNewLine
        this.FormatedErrorMsg = this.FormatedErrorMsg & "Windows Verion: " & vbTab & .WindowsVerion & vbNewLine
        this.FormatedErrorMsg = this.FormatedErrorMsg & "App Version: " & vbTab & .AppVersion & vbNewLine
        this.FormatedErrorMsg = this.FormatedErrorMsg & "User Name: " & vbTab & .userName & vbNewLine
        this.FormatedErrorMsg = this.FormatedErrorMsg & "Common Name: " & vbTab & .UserCommonName & vbNewLine
        this.FormatedErrorMsg = this.FormatedErrorMsg & "Email Address: " & vbTab & .UserEmailAddress & vbNewLine
    End With

End Sub

Private Sub CreateLogfile(ByVal FormatedErrorString As String, ByVal projectName As String)
    
    Dim FileSysObj As Object, LogFile As Object
    Dim DesktopFolderPath As String
    
        DesktopFolderPath = GetDesktopFolderPath()
        
        Set FileSysObj = CreateObject("Scripting.FileSystemObject")
        
        projectName = ReplaceIllegalCharacters(projectName)
    
        Set LogFile = FileSysObj.CreateTextFile(DesktopFolderPath & "\" & projectName & "_" & Format(CStr(Now), "yyyymmddhhmmss") & ".txt", True)
            LogFile.WriteLine this.FormatedErrorMsg
            LogFile.Close
            
End Sub

Private Function GetDesktopFolderPath() As String

    Const DEFAULT_FOLDER_NAME As String = "Desktop"
    
        With CreateObject("WScript.Shell")
            GetDesktopFolderPath = .specialfolders(DEFAULT_FOLDER_NAME)
        End With
        
End Function

Private Function ReplaceIllegalCharacters(ByVal FileName As String) As String

    With CreateObject("VBScript.RegExp")
        .pattern = "(\s|\\|/|<|>|\|\|\?|:)"
        .Global = True
        .IgnoreCase = True
        ReplaceIllegalCharacters = .Replace(FileName, "_")
    End With

End Function


Private Sub TrySendEmail(ByVal networkHostName As String, ByVal projectName As String, ByVal emailToAddress As String)

    Dim NetworkStatus As NetworkConnectionMonitor
    
    Set NetworkStatus = New NetworkConnectionMonitor
    
        On Error GoTo CleanFail
                
        If Not NetworkStatus.ConnectedToInternet Then
            MsgBox "You Are not Connected to the internet", _
                    vbExclamation, "Cannot Send Email"
            Exit Sub
        End If
                                                        
        If Not NetworkStatus.IsConnectedToHost(networkHostName, 2) Then
                MsgBox "You Are not Connected to the company VPN", _
                        vbExclamation, "Cannot Send Email"
                Exit Sub
        End If
        
        'If Not EmailManager.IsOutlookAvailiable() Then
    
        EmailManager.SendOutlookEmail emailToAddress, "Error: " & projectName, this.FormatedErrorMsg
        
CleanExit:
    Exit Sub
    
CleanFail:
    MsgBox "For some reason the email could not be sent. " & _
           "Please send the log file for this error to: " & emailToAddress, _
           vbExclamation, "Error: Cannot Send Email"
    Resume CleanExit
        
End Sub
