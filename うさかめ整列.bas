Attribute VB_Name = "�������ߐ���"
Sub �������ߐ���()
    For Each sld In ActivePresentation.Slides
        For Each sh In sld.Shapes
            If sh.name = ����.���ߖ� Then
                sh.Left = 0
            End If
            If sh.name = ������.�������� Then
                sh.Left = 0
            End If
        Next
    Next
End Sub

Sub �������߃R�s�[()
    �I�u�W�F�N�g�R�s�[ (������.��������)
    �I�u�W�F�N�g�R�s�[ (����.���ߖ�)
End Sub


Sub �I�u�W�F�N�g�R�s�[(objName As String)
    Dim doCopy As Boolean
    doCopy = False
    
    For Each sld In ActivePresentation.Slides
        If sld.SlideIndex = 1 Then
            For Each sh In sld.Shapes
                If sh.name = objName Then
                    sh.Copy
                    doCopy = True
                End If
            Next
        Else
            If Not doCopy Then
                MsgBox "�I�u�W�F�N�g���R�s�[�ł��܂���ł���:" + objName
                Exit For
            End If
            
            sld.Shapes.Paste
        End If
    Next
End Sub

