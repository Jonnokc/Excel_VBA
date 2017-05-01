
If DocumentType = PowerForm OR IView AND _
   ControlType = UltraGrid OR PowerGrid AND _
   View Level = "Detail" _
   Then
    Grid Notes = "This code is part of a grid that is set to result at the detail level.  This will technically work for HealtheIntent but the results from the grid will not be easy to view in PowerChart." And _
         Team = "PCST"


    If DocumentType = PowerForm OR IView AND _
       ControlType = UltraGrid OR PowerGrid AND _
       View Level = "Grid" _
       Then
        Grid Notes = "This code is part of a grid that is set to result at the grid level.  Individual results have a view_level of 0 and do not make it to HealtheIntent.  The grid would have to be set at the detail level in order for these results to be used to meet the measure." And _
             Team = "Consulting"
