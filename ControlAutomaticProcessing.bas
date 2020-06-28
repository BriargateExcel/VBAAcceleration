Attribute VB_Name = "AutomaticProcessing"
Option Explicit

Private Const Module_Name As String = "AutomaticProcessing."

Public Sub TurnOffAutomaticProcessing()

    ' This routine turns off all the automatic processing that slows things down
    
    Const RoutineName As String = Module_Name & "TurnOffAutomaticProcessing"
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = False
    Application.Calculation = xlManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub      ' TurnOffAutomaticProcessing

Public Sub TurnOnAutomaticProcessing()

    ' This routine turns on all the automatic processing that slows things down
    ' Reverses the things that were turned off in TurnOffAutomaticProcessing
    
    Const RoutineName As String = Module_Name & "TurnOnAutomaticProcessing"
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Application.StatusBar = False

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub      ' TurnOnAutomaticProcessing