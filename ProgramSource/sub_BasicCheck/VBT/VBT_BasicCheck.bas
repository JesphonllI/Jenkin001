Attribute VB_Name = "VBT_BasicCheck"
Option Explicit

'===============================================================================
' Digital Continuity
'===============================================================================
Public Function Digital_Continuity(TestPins As String, _
                                   ISource As Double, _
                                   ISync As Double) As Long

    '___ ALT ___________________________________________________________________
    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered
    
    '___ Disconnect Digital ____________________________________________________
    TheHdw.Digital.Pins(TestPins).Disconnect
    
    '___ Apply the loads _______________________________________________________
    Dim SrcResults As New PinListData
    Dim SinkResults As New PinListData
    
    With TheHdw.PPMU.Pins(TestPins)
        .ClampVHi = 0.6
        .ClampVLo = -0.5
        .ForceI ISource, Abs(2 * ISource)
        .Connect
        .Gate = tlOn
        
        TheHdw.Wait 5 * ms
        SrcResults = .Read(tlPPMUReadMeasurements, 10, tlPPMUReadingFormatAverage)
        TheExec.Flow.TestLimit resultVal:=SrcResults, ForceResults:=tlForceFlow
        
        .ForceI ISync, (2 * ISource)
        
        TheHdw.Wait 5 * ms
        SinkResults = .Read(tlPPMUReadMeasurements, 10, tlPPMUReadingFormatAverage)
        TheExec.Flow.TestLimit resultVal:=SinkResults, ForceResults:=tlForceFlow
        
        .Gate = tlOff
        .Disconnect
        .ClampVHi = 6.5
        .ClampVLo = -1.5

        
    End With
       
    '___ Cleanup _______________________________________________________________
    TheHdw.Digital.Pins(TestPins).Connect
    
End Function

'===============================================================================
' Digital Continuity
'===============================================================================
Public Function Supply_Shorts(TestPins As String) As Long

    '___ ALT ___________________________________________________________________
    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered
    
    '___ Make the measurements _________________________________________________
    Dim LeakResults As New PinListData
    
    TheHdw.Wait 5 * ms
    
    With TheHdw.DCVS.Pins(TestPins)
        .CurrentRange = 1 * ma
        .Mode = tlDCVSModeVoltage
        LeakResults = .Meter.Read(tlStrobe, 10, 10000, tlDCVSMeterReadingFormatAverage)
        TheExec.Flow.TestLimit resultVal:=LeakResults, ForceResults:=tlForceFlow
        
    End With
       
End Function



