Attribute VB_Name = "������"
Sub ���������s()
    On Error Resume Next
    
    For i = 1 To ActivePresentation.Slides.Count
        For Each sh In ActivePresentation.Slides(i).Shapes
            If sh.name = �������� Then
                sh.Left = (i - 1) * (ActivePresentation.PageSetup.SlideWidth / ActivePresentation.Slides.Count)
            End If
        Next
    Next
End Sub

Function ��������() As String
    �������� = "������"
End Function
