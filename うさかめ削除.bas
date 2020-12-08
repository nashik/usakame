Attribute VB_Name = "うさかめ削除"
'かめ、うさぎオブジェクト削除用
Sub うさかめオブジェクトをすべて削除()
    オブジェクト削除 (うさぎ.うさぎ名)
    オブジェクト削除 (かめ.かめ名)
End Sub

Sub オブジェクト削除(name As String)
    For Each sld In ActivePresentation.Slides
        For Each sh In sld.Shapes
            If sh.name = name Then
                sh.Delete
            End If
        Next
    Next
End Sub
