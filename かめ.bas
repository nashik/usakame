Attribute VB_Name = "����"
Sub ���ߎ��s(totalSec As Integer, startTime)
    On Error Resume Next
    Dim secCount As Integer
    Dim tmpTime As Integer
    secCount = (Minute(Now - startTime) * 60) + Second(Now - startTime)
    tmpTime = 0
    
    '�J�E���g�����B���߂𓮂���
    Do While totalSec >= secCount
        DoEvents
        If (tmpTime <> secCount) Then
            tmpTime = secCount
            For Each sld In ActivePresentation.Slides
                For Each sh In sld.Shapes
                    If sh.name = ���ߖ� Then
                        sh.IncrementLeft (ActivePresentation.PageSetup.SlideWidth / totalSec)
                    End If
                Next
            Next
            
            'For Each sh In ActivePresentation.Slides(SlideShowWindows(1).View.CurrentShowPosition).Shapes
            '    If sh.name = ���ߖ� Then
            '        sh.LeIncrementLeft (ActivePresentation.PageSetup.SlideWidth / totalSec)
            '    End If
            'Next
            
            If (ActivePresentation.SlideShowWindow.Active = msoTriStateMixed) Then
                Exit Do
            End If
        End If

        secCount = (Minute(Now - startTime) * 60) + Second(Now - startTime)
    Loop
End Sub

Function ���ߖ�() As String
    ���ߖ� = "����"
End Function
