Public Sub ImportSourceFiles(sourcePath As String)
Dim file As String
    file = Dir(sourcePath)
    While (file &lt;&gt; vbNullString)
        Application.VBE.ActiveVBProject.VBComponents.Import sourcePath &amp; file
        file = Dir
    Wend
End Sub
