Attribute VB_Name = "�������ߍ폜"
'���߁A�������I�u�W�F�N�g�폜�p
Sub �������߃I�u�W�F�N�g�����ׂč폜()
    �I�u�W�F�N�g�폜 (������.��������)
    �I�u�W�F�N�g�폜 (����.���ߖ�)
End Sub

Sub �I�u�W�F�N�g�폜(name As String)
    For Each sld In ActivePresentation.Slides
        For Each sh In sld.Shapes
            If sh.name = name Then
                sh.Delete
            End If
        Next
    Next
End Sub
