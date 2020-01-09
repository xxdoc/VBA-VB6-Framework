VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppSessionVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const APP_SESSION_VARIABLES_NAME As String = "APP_SESSION_VARIABLES"

Public Enum AppSessionVariablesError
    AlreadyExists = vbObjectError + 1042
    DoesNotExist
    NonDefaultInstance
End Enum

Private Type TAppSessionVariables
    dictSessionVars As Object
    Host As New mscoree.CorRuntimeHost
End Type

Private this As TAppSessionVariables

Private Sub Class_Initialize()

    ThrowIfNonDefaultInstance
    If this.dictSessionVars Is Nothing Then Set this.dictSessionVars = GetPersistentDictionary()
	
End Sub

Private Sub Class_Terminate()

    Set this.Host = Nothing
	
End Sub

Private Function GetPersistentDictionary() As Object

    Dim domain As mscorlib.AppDomain
    Dim dict As Object
    
        this.Host.Start
        this.Host.GetDefaultDomain domain

        If IsObject(domain.GetData(APP_SESSION_VARIABLES_NAME)) Then
            Set dict = domain.GetData(APP_SESSION_VARIABLES_NAME)
			
        Else
            Set dict = CreateObject("Scripting.Dictionary")
            domain.SetData APP_SESSION_VARIABLES_NAME, dict
			
        End If

    Set GetPersistentDictionary = dict

End Function

Public Sub Add(ByVal name As String, ByVal value As Variant)

    ThrowIfAlreadyExists name
    
    If IsObject(value) Then
        Set this.dictSessionVars(name) = value
    Else
        this.dictSessionVars(name) = value
    End If
    
End Sub

Public Function Exists(ByVal name As String) As Boolean
    Exists = this.dictSessionVars.Exists(name)
End Function

Public Sub Remove(ByVal name As String)
    On Error Resume Next
    this.dictSessionVars.Remove name
End Sub

Public Sub RemoveAll()
    On Error Resume Next
    this.dictSessionVars.RemoveAll
End Sub

Public Function Item(ByVal name As String) As Variant

    ThrowIfDoesNotExist name
    
    If IsObject(this.dictSessionVars(name)) Then
        Set Item = this.dictSessionVars(name)
		
    Else
        Item = this.dictSessionVars(name)
		
    End If

End Function

Public Function Count() As Long
    Count = this.dictSessionVars.Count
End Function


Private Sub ThrowIfAlreadyExists(ByVal name As String)

    If Exists(name) Then
        Err.Raise Number:=AppSessionVariablesError.AlreadyExists, _ 
				          Source:=TypeName(Me), _
                  Description:="This App session Variable already exists."
				  
    End If
	
End Sub

Private Sub ThrowIfDoesNotExist(ByVal name As String)

    If Not Exists(name) Then
        Err.Raise Number:=AppSessionVariablesError.DoesNotExist, _ 
				          Source:=TypeName(Me), _
                  Description:="This App session does not exist."
				  
    End If
	
End Sub

Private Sub ThrowIfNonDefaultInstance()

    If Not Me Is AppSessionVariables Then
        Err.Raise Number:=AppSessionVariablesError.NonDefaultInstance, _ 
				          Source:=Typename(Me), 
				          Description:="You cannot use the non-default instance of a static class."
				  
    End If
	
End Sub



