Attribute VB_Name = "VBT_CoreE"
Option Explicit

Private ExecCount As Long


'===============================================================================
' Core E Test
'===============================================================================
Public Function Core_E_Test(TestPins As String) As Long

    Dim Result As New SiteDouble
    Dim vSite As Variant

    On Error GoTo errHandler

    ExecCount = ExecCount + 1
    For Each vSite In TheExec.Sites.Selected
        Result = 5.678 * (ExecCount)
    Next vSite

    TheExec.Flow.TestLimit resultVal:=Result
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function


