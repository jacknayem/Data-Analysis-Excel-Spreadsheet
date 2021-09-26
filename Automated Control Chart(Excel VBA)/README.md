# Automated Dynamic Chart
## The view of chart
![](https://github.com/jacknayem/Data-Analysis-Excel-Spreadsheet/blob/main/Automated%20Control%20Chart(Excel%20VBA)/Automated%20Control%20Chart.png)

### To create random value
````
Sub Simulate()
    With ActiveSheet
        mean = .Range("Simulation_Mean").Value
        std = .Range("Simulation_Std").Value
        With .Range("Actual_Data_Header")
            If .Offset(1, 0).Value = “” Then

                .Offset(1, 0).Value = WorksheetFunction.Norm_Inv(Rnd(), mean, std)

            Else

                .End(xlDown).Offset(1, 0).Value = WorksheetFunction.Norm_Inv(Rnd(), mean, std)

            End If
        End With
    End With
End Sub
````
### Erase all value
````
Sub Restart()
    With ActiveSheet
        mean = .Range("Simulation_Mean").Value
        std = .Range("Simulation_Std").Value
        first_cell_ref = .Range("Actual_Data_Header").Offset(1, 0).Address
        
        .Range(first_cell_ref & ":B998").ClearContents
        .Range(first_cell_ref).Value = WorksheetFunction.Norm_Inv(Rnd(), mean, std)
    End With
End Sub
````
