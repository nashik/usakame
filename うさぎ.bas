Attribute VB_Name = "うさぎ"
Sub うさぎ実行()
    On Error Resume Next
    
    For i = 1 To ActivePresentation.Slides.Count
        For Each sh In ActivePresentation.Slides(i).Shapes
            If sh.name = うさぎ名 Then
                sh.Left = (i - 1) * (ActivePresentation.PageSetup.SlideWidth / ActivePresentation.Slides.Count)
            End If
        Next
    Next
End Sub

Function うさぎ名() As String
    うさぎ名 = "うさぎ"
End Function
