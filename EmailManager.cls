VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("MyLibrary.ExtensibleLogging")

Option Explicit

Public Function IsOutlookAvailiable() As Boolean
    
      On Error Resume Next
	  Dim outLookApp As Object
      Set outLookApp = CreateObject("Outlook.Application")
      IsOutlookAvailiable = Not outLookApp Is Nothing
      On Error GoTo 0
End Function

Public Sub SendOutlookEmail(ByVal EmailToAddress As String, ByVal EmailSubject As String, _
                            ByVal EmailBody As String)
        
    On Error GoTo CleanFail
		Dim outLookApp As Object
    Set outLookApp = CreateObject("Outlook.Application")
        
		Dim outMail As Object
        Set outMail = outLookApp.CreateItem(0)
		On Error GoTo CleanFail
		With outMail
			.To = EmailToAddress
			.subject = EmailSubject
			.body = EmailBody
			.Send
		End With

CleanExit:
    Set outLookApp = Nothing
    Set outMail = Nothing
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub

Public Sub SendCDOEmail(ByVal smtpServerAddress As String, ByVal fromAddress As String, _
                        ByVal toAddress As String, ByVal subject As String, ByVal body As String, _
                        Optional smtpServerPort As Integer = 25)

    On Error GoTo CleanFail
	Dim cdoMsg As Object
    Set cdoMsg = CreateObject("CDO.Message")
	
	Dim cdoConfig As Object
    Set cdoConfig = CreateObject("CDO.Configuration")

    cdoConfig.Load -1     ' CDO Source Defaults
	
	Dim configFields As Object
    Set configFields = cdoConfig.Fields
        With configFields
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServerAddress
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpServerPort") = smtpServerPort
            .Update
        End With

        With cdoMsg
            Set .Configuration = cdoConfig
            .To = toAddress
            .from = Chr(34) & VBA.Environ$("USERNAME") & Chr(34) & " <" & _
                              fromAddress & ">"
            .subject = subject

            'Set importance or Priority to high
            .Fields("urn:schemas:httpmail:importance") = 2
            .Fields("urn:schemas:mailheader:X-Priority") = 1
            .Fields.Update

            .TextBody = body

            .Send
        End With

CleanExit:
    Set cdoMsg = Nothing
    Set cdoConfig = Nothing
    Set configFields = Nothing
    Exit Sub

CleanFail:
    Resume CleanExit
    
End Sub

