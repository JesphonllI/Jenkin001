Attribute VB_Name = "VBT_CoreD"
Option Explicit


Private ExecCount As Long


'===============================================================================
' Core D Test
'===============================================================================
Public Function Core_D_Test(TestPins As String) As Long

    Dim Result As New SiteDouble
    Dim vSite As Variant

    On Error GoTo errHandler

    ExecCount = ExecCount + 1
    For Each vSite In TheExec.Sites.Selected
        Result = 11.123 * (ExecCount)
    Next vSite

    TheExec.Flow.TestLimit resultVal:=Result
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

