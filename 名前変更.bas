Attribute VB_Name = "名前変更"
'オブジェクトを指定して実行することでオブジェクトに内部的な名前をつけます。
'うさぎオブジェクトはうさぎ.うさぎ名、かめオブジェクトはかめ.かめ名を設定することで
'どんなオブジェクトでもうさかめとして動作します。
Sub 選択中のオブジェクトをうさぎに設定()
  On Error Resume Next
  ActiveWindow.Selection.ShapeRange.name = うさぎ.うさぎ名
End Sub

Sub 選択中のオブジェクトをかめに設定()
  On Error Resume Next
  ActiveWindow.Selection.ShapeRange.name = かめ.かめ名
End Sub

Sub 選択中のオブジェクトをうさかめから解除()
  On Error Resume Next
  ActiveWindow.Selection.ShapeRange.name = "NoName"
End Sub
