Attribute VB_Name = "VBT_Functional"
Option Explicit

'==================================================================================
' Run a simple functional test with the specified result mode
'==================================================================================
Public Function RunFunctionalTest(PatName As Pattern, _
                                  TestResultMode As tlResultMode) As Long

    Dim errMsg As String
    
    '___ Init _____________________________________________________________________
    On Error GoTo errHandler
    
    '___ ALT ______________________________________________________________________
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
    
    '___ Run the test _____________________________________________________________
    Call TheHdw.Patterns(PatName.Value).Test(pfAlways, 0, TestResultMode)
    
    '___ Cleanup __________________________________________________________________
    RunFunctionalTest = TL_SUCCESS


    Exit Function
    
errHandler:
    If (Len(errMsg) = 0) Then
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & Err.Description)
    Else
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & errMsg)
    End If
    
    Call TheExec.ErrorReport
    RunFunctionalTest = TL_ERROR

End Function

'==================================================================================
' Test the overall result and then per the highest level elements in the set
'==================================================================================
Public Function RunOneAndElementFTL(PatSetName As PatternSet) As Long

    Dim errMsg As String
    
    '___ Init _____________________________________________________________________
    On Error GoTo errHandler
    
    '___ ALT ______________________________________________________________________
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
    
    '___ Run the pattern set ______________________________________________________
    Call TheHdw.Patterns(PatSetName.Value).Start
    TheHdw.Patterns(PatSetName.Value).HaltWait
    
    '___ Queue the PatSet Object commands _________________________________________
    Dim elementCnt As Long
    Dim elementIdx As Long
    Dim elementNames() As String
    Dim elementPF As New SiteArray
    
    With TheHdw.Digital.Patgen.ReadPatternSetResults.Elements
        elementCnt = .Count
        elementNames = .Names
        elementPF = .PFResults
    End With
        
    '___ Test the high level result _______________________________________________
    Dim overallPF As New SiteBoolean
    overallPF = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    
    TheExec.Flow.FunctionalTestLimit resultVal:=overallPF, _
                                     PatName:=PatSetName.Value, _
                                     ForceResults:=tlForceFlow
                                     
    '___ Test the high level elements _____________________________________________
    For elementIdx = 0 To (elementCnt - 1)
        TheExec.Flow.FunctionalTestLimit resultVal:=elementPF.Element(elementIdx), _
                                         PatName:=elementNames(elementIdx), _
                                         ForceResults:=tlForceFlow
    Next elementIdx
    
    '___ Cleanup __________________________________________________________________
    RunOneAndElementFTL = TL_SUCCESS


    Exit Function
    
errHandler:
    If (Len(errMsg) = 0) Then
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & Err.Description)
    Else
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & errMsg)
    End If
    
    Call TheExec.ErrorReport
    RunOneAndElementFTL = TL_ERROR

End Function

'==================================================================================
' Test the overall result and then all of the underlying elements
'==================================================================================
Public Function RunOneAndAllElementsFTL(PatSetName As PatternSet, _
                                        Optional TD_A As String = "", _
                                        Optional TD_B As String = "") As Long

    Dim errMsg As String
    
    '___ Init _____________________________________________________________________
    On Error GoTo errHandler
    
    '___ ALT ______________________________________________________________________
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
    
    '___ Run the pattern set ______________________________________________________
    Call TheHdw.Patterns(PatSetName.Value).Start
    TheHdw.Patterns(PatSetName.Value).HaltWait
    
    '___ Test the high level result and all the elements _____________________________
    Dim overallPF_A As New SiteBoolean
    Dim overallPF_B As New SiteBoolean
    
    If (Len(TD_A) = 0) And (Len(TD_B) = 0) Then
    
        '___ Single time domain ______________________________________________________
        overallPF_A = TheHdw.Digital.TimeDomains(TD_A).Patgen.PatternBurstPassedPerSite
        
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_A, _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
       
        TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.Patgen.ReadPatternSetResults.Elements, _
                                         ForceResults:=tlForceFlow
    
    Else
        
        '___ Multi time domain ______________________________________________________
        overallPF_A = TheHdw.Digital.TimeDomains(TD_A).Patgen.PatternBurstPassedPerSite
        overallPF_B = TheHdw.Digital.TimeDomains(TD_B).Patgen.PatternBurstPassedPerSite
        
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_A.LogicalAnd(overallPF_B), _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
        
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_A, _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
        
        TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.TimeDomains(TD_A).Patgen.ReadPatternSetResults.Elements, _
                                         ForceResults:=tlForceFlow
        
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_B, _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
        
        TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.TimeDomains(TD_B).Patgen.ReadPatternSetResults.Elements, _
                                         ForceResults:=tlForceFlow
    End If

        '___ Cleanup __________________________________________________________________
    RunOneAndAllElementsFTL = TL_SUCCESS


    Exit Function
    
errHandler:
    If (Len(errMsg) = 0) Then
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & Err.Description)
    Else
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & errMsg)
    End If
    
    Call TheExec.ErrorReport
    RunOneAndAllElementsFTL = TL_ERROR

End Function


'==================================================================================
' Test the overall result and then all of the underlying modules
'==================================================================================
Public Function RunOneAndAllModsFTL(PatSetName As PatternSet, _
                                        Optional TD_A As String = "", _
                                        Optional TD_B As String = "") As Long

    Dim errMsg As String
    
    '___ Init _____________________________________________________________________
    On Error GoTo errHandler
    
    '___ ALT ______________________________________________________________________
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
    
    '___ Run the pattern set ______________________________________________________
    Call TheHdw.Patterns(PatSetName.Value).Start
    TheHdw.Patterns(PatSetName.Value).HaltWait
    
    '___ Test the high level result and all the elements _____________________________
    Dim overallPF_A As New SiteBoolean
    Dim overallPF_B As New SiteBoolean
    
    If (Len(TD_A) = 0) And (Len(TD_B) = 0) Then
    
        '___ Single time domain ______________________________________________________
        overallPF_A = TheHdw.Digital.TimeDomains(TD_A).Patgen.PatternBurstPassedPerSite
        
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_A, _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
       
        TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.Patgen.ReadPatternSetResults.Modules, _
                                         ForceResults:=tlForceFlow
    
    
    Else
        
        '___ Multi time domain ______________________________________________________
        overallPF_A = TheHdw.Digital.TimeDomains(TD_A).Patgen.PatternBurstPassedPerSite
        overallPF_B = TheHdw.Digital.TimeDomains(TD_B).Patgen.PatternBurstPassedPerSite
        
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_A.LogicalAnd(overallPF_B), _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
        
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_A, _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
        
        TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.TimeDomains(TD_A).Patgen.ReadPatternSetResults.Modules, _
                                         ForceResults:=tlForceFlow
        
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_B, _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
        
        TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.TimeDomains(TD_B).Patgen.ReadPatternSetResults.Modules, _
                                         ForceResults:=tlForceFlow
    End If


    '___ Cleanup __________________________________________________________________
    RunOneAndAllModsFTL = TL_SUCCESS


    Exit Function
    
errHandler:
    If (Len(errMsg) = 0) Then
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & Err.Description)
    Else
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & errMsg)
    End If
    
    Call TheExec.ErrorReport
    RunOneAndAllModsFTL = TL_ERROR

End Function

'==================================================================================
' Test the overall result and then log modules by set to allow binning by set
'==================================================================================
Public Function RunOneAndModsByElementFTL(PatSetName As PatternSet, _
                                          Optional TD_A As String = "", _
                                          Optional TD_B As String = "") As Long

    Dim errMsg As String
    
    '___ Init _____________________________________________________________________
    On Error GoTo errHandler
    
    '___ ALT ______________________________________________________________________
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
    
    '___ Run the pattern set ______________________________________________________
    Call TheHdw.Patterns(PatSetName.Value).Start
    TheHdw.Patterns(PatSetName.Value).HaltWait
    
    '___ Queue the PatSet Object commands _________________________________________
    Dim elementCnt As Long
    Dim elementIdx As Long
    Dim elementNames() As String
    Dim elementPF As New SiteArray
    
    With TheHdw.Digital.Patgen.ReadPatternSetResults.Elements
        elementCnt = .Count
        elementNames = .Names
        elementPF = .PFResults
    End With
    
    '___ Test the high level result and the modules by element ____________________
    Dim overallPF_A As New SiteBoolean
    Dim overallPF_B As New SiteBoolean
    
    If (Len(TD_A) = 0) And (Len(TD_B) = 0) Then
    
        '___ Single time domain ______________________________________________________
        overallPF_A = TheHdw.Digital.TimeDomains(TD_A).Patgen.PatternBurstPassedPerSite
        
        '___ Overall result __________________________________________________________
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_A, _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
                                         
        '___ Modules per element _____________________________________________________
        For elementIdx = 0 To (elementCnt - 1)
            TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.Patgen.ReadPatternSetResults.Elements(elementIdx).Modules, _
                                             ForceResults:=tlForceFlow
        Next elementIdx
    
    Else
        
        '___ Multi time domain ______________________________________________________
        overallPF_A = TheHdw.Digital.TimeDomains(TD_A).Patgen.PatternBurstPassedPerSite
        overallPF_B = TheHdw.Digital.TimeDomains(TD_B).Patgen.PatternBurstPassedPerSite
        
        '___ Overall TD_A && TD_B ________________________________________________
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_A.LogicalAnd(overallPF_B), _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
        
        '___ Overall TD_A ________________________________________________________
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_A, _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
        
        '___ Modules per element TD_A _____________________________________________
        elementCnt = TheHdw.Digital.TimeDomains(TD_A).Patgen.ReadPatternSetResults.Elements.Count
        For elementIdx = 0 To (elementCnt - 1)
            TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.TimeDomains(TD_A).Patgen.ReadPatternSetResults.Elements(elementIdx).Modules, _
                                             ForceResults:=tlForceFlow
        Next elementIdx
        
        '___ Overall TD_B ________________________________________________________
        TheExec.Flow.FunctionalTestLimit resultVal:=overallPF_B, _
                                         PatName:=PatSetName.Value, _
                                         ForceResults:=tlForceFlow
        
        elementCnt = TheHdw.Digital.TimeDomains(TD_B).Patgen.ReadPatternSetResults.Elements.Count
        For elementIdx = 0 To (elementCnt - 1)
            TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.TimeDomains(TD_B).Patgen.ReadPatternSetResults.Elements(elementIdx).Modules, _
                                             ForceResults:=tlForceFlow
        Next elementIdx
        
    End If


    '___ Cleanup __________________________________________________________________
    RunOneAndModsByElementFTL = TL_SUCCESS


    Exit Function
    
errHandler:
    If (Len(errMsg) = 0) Then
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & Err.Description)
    Else
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & errMsg)
    End If
    
    Call TheExec.ErrorReport
    RunOneAndModsByElementFTL = TL_ERROR

End Function


'==================================================================================
' Test all of the underlying elements
'==================================================================================
Public Function RunAllElementsFTL(PatSetName As PatternSet, _
                                  Optional ForceAlarm As Boolean = False) As Long

    Dim errMsg As String
    
    '___ Init _____________________________________________________________________
    On Error GoTo errHandler
    
    '___ ALT ______________________________________________________________________
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
    
    '___ Optionally force an alarm ________________________________________________
    If (ForceAlarm = True) Then
        With TheHdw.DCVS.Pins("VCC")
            .Voltage.Main.Value = 0.5
            .Connect
            .Gate = True
        End With
    End If
        
    '___ Run the pattern set ______________________________________________________
    Call TheHdw.Patterns(PatSetName.Value).Start
    TheHdw.Patterns(PatSetName.Value).HaltWait
    
    '___ Test the high level result _______________________________________________
    Dim overallPF As New SiteBoolean
    overallPF = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    
    TheExec.Flow.FunctionalTestLimit resultVal:=overallPF, _
                                     PatName:=PatSetName.Value, _
                                     ForceResults:=tlForceFlow
    
    '___ Test all the elements _____________________________________________
    If (TheHdw.Digital.Patgen.ReadPatternSetResults.Elements.Count > 0) Then
        TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.Patgen.ReadPatternSetResults.Elements, _
                                         ForceResults:=tlForceFlow
    End If

    '___ Cleanup __________________________________________________________________
    If (ForceAlarm = True) Then
        With TheHdw.DCVS.Pins("VCC")
            .Gate = True
            .Disconnect
        End With
    End If

    RunAllElementsFTL = TL_SUCCESS


    Exit Function
    
errHandler:
    If (Len(errMsg) = 0) Then
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & Err.Description)
    Else
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & errMsg)
    End If
    
    Call TheExec.ErrorReport
    RunAllElementsFTL = TL_ERROR

End Function

'==================================================================================
' Test all of the underlying elements after an alarm occurs
'==================================================================================
Public Function RunAllElementsAlarmFTL(PatSetName As PatternSet) As Long

    Dim errMsg As String
    
    '___ Init _____________________________________________________________________
    On Error GoTo errHandler
    
    '___ ALT ______________________________________________________________________
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
    
    '___ Optionally force an alarm on a site ______________________________________
    Dim Site As Variant
    
    For Each Site In TheExec.Sites.Active
        If (Site = 1) Then
            With TheHdw.DCVS.Pins("VCC")
                .Voltage.Main.Value = 0.5
                .Connect
                .Gate = True
            End With
        End If
    Next Site
        
    '___ Run the pattern set ______________________________________________________
    Call TheHdw.Patterns(PatSetName.Value).Start
    TheHdw.Patterns(PatSetName.Value).HaltWait
    
    '___ Test the high level result _______________________________________________
    Dim overallPF As New SiteBoolean
    overallPF = TheHdw.Digital.Patgen.PatternBurstPassedPerSite

    TheExec.Flow.FunctionalTestLimit resultVal:=overallPF, _
                                     PatName:=PatSetName.Value, _
                                     ForceResults:=tlForceFlow
    
    '___ Test all the elements _____________________________________________
    If (TheHdw.Digital.Patgen.ReadPatternSetResults.Elements.Count > 0) Then
        TheExec.Flow.FunctionalTestLimit resultVal:=TheHdw.Digital.Patgen.ReadPatternSetResults.Elements, _
                                         ForceResults:=tlForceFlow
    End If

    '___ Cleanup __________________________________________________________________
    With TheHdw.DCVS.Pins("VCC")
        .Gate = True
        .Disconnect
    End With

    RunAllElementsAlarmFTL = TL_SUCCESS


    Exit Function
    
errHandler:
    If (Len(errMsg) = 0) Then
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & Err.Description)
    Else
        Call TheExec.ErrorLogMessage("*E* Test " & TL_C_ERRORSTR & _
                                     ", Instance: " & TheExec.DataManager.InstanceName & _
                                     ", Description: " & errMsg)
    End If
    
    Call TheExec.ErrorReport
    RunAllElementsAlarmFTL = TL_ERROR

End Function



