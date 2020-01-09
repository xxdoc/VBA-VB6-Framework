VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Framwork.Common")
Option Explicit

Public Function CompareTo(currentInstance As Variant, pivot As Variant) As Integer
'Compares this instance with another; returns one of the following values:
'   -1 if [pivot] is smaller than this instance.
'    1 if [pivot] is greater than this instance.
'    0 otherwise.
End Function


Private Sub Class_Initialize()
    Err.Raise Number:=vbObjectError + 1024, _ 
			        Source:=TypeName(Me), _ 
			        Description:="An Interface class must not be instantiated."
End Sub

