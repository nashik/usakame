Attribute VB_Name = "Ç§Ç≥Ç¨"
Sub Ç§Ç≥Ç¨é¿çs()
    On Error Resume Next
    
    For i = 1 To ActivePresentation.Slides.Count
        For Each sh In ActivePresentation.Slides(i).Shapes
            If sh.name = Ç§Ç≥Ç¨ñº Then
                sh.Left = (i - 1) * (ActivePresentation.PageSetup.SlideWidth / ActivePresentation.Slides.Count)
            End If
        Next
    Next
End Sub

Function Ç§Ç≥Ç¨ñº() As String
    Ç§Ç≥Ç¨ñº = "Ç§Ç≥Ç¨"
End Function
