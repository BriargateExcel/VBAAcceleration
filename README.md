#### Speeding up Your VBA Code

There are several automatic Excel behaviors that you can generally turn off while executing VBA code. Turning these behaviors off will make your VBA code run a bit faster.

- `DisplayAlerts` = False turns off various messages; the user does not have to acknowledge them
- `Calculation = xlManual` turns off recalculation every time a cell value changes
- `EnableEvents = False` turns off various events
- `ScreenUpdating = False` turns off the visual updating on the user’s monitor

I keep these and other routines I commonly use in a module called `CommonRoutines` that I include in every project requiring VBA code.

There is information on my error handling practices in https://github.com/BriargateExcel/Error_Handling

#### How to use

- Put `TurnOffAutomaticProcessing` in the top-level routine in your application so that it runs first
- Put `TurnOnAutomaticProcessing` in the top-level routing in your application so that it runs last

#### Caution

position `TurnOnAutomaticProcessing` so that it runs regardless of any exceptions raised during processing. If `TurnOnAutomaticProcessing` does not run, Excel's automatic processing will not execute. You could consider adding a button to your `Personal.xlsb` workbook to execute `TurnOnAutomaticProcessing` to restore normal processing. This is described in https://github.com/BriargateExcel/AddingVBARibbonElements.