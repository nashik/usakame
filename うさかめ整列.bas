Attribute VB_Name = "うさかめ整列"
Sub うさかめ整列()
    For Each sld In ActivePresentation.Slides
        For Each sh In sld.Shapes
            If sh.name = かめ.かめ名 Then
                sh.Left = 0
            End If
            If sh.name = うさぎ.うさぎ名 Then
                sh.Left = 0
            End If
        Next
    Next
End Sub

Sub うさかめコピー()
    オブジェクトコピー (うさぎ.うさぎ名)
    オブジェクトコピー (かめ.かめ名)
End Sub


Sub オブジェクトコピー(objName As String)
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
                MsgBox "オブジェクトをコピーできませんでした:" + objName
                Exit For
            End If
            
            sld.Shapes.Paste
        End If
    Next
End Sub

