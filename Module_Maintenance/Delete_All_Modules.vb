Public Sub RemoveAllModules()

'
' Deletes all modules except for the Module_Maintenance module.
'
'
Dim project As VBProject


    Set project = Application.VBE.ActiveVBProject

    Dim comp As VBComponent
    For Each comp In project.VBComponents
        If Not comp.name = "Module_Maintenance" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
            project.VBComponents.Remove comp
        End If
    Next
End Sub
