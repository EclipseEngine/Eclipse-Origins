Attribute VB_Name = "modCustomScripts"
Public Sub CustomScript(Index As Long, caseID As Long)

   On Error GoTo errorhandler

    Select Case caseID
        Case Else
            PlayerMsg Index, "You just activated custom script " & caseID & ". This script is not yet programmed.", BrightRed
    End Select


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "CustomScript", "modCustomScripts", Err.Number, Err.Description, Erl
    Err.Clear
End Sub
