Attribute VB_Name = "���O�ύX"
'�I�u�W�F�N�g���w�肵�Ď��s���邱�ƂŃI�u�W�F�N�g�ɓ����I�Ȗ��O�����܂��B
'�������I�u�W�F�N�g�͂�����.���������A���߃I�u�W�F�N�g�͂���.���ߖ���ݒ肷�邱�Ƃ�
'�ǂ�ȃI�u�W�F�N�g�ł��������߂Ƃ��ē��삵�܂��B
Sub �I�𒆂̃I�u�W�F�N�g���������ɐݒ�()
  On Error Resume Next
  ActiveWindow.Selection.ShapeRange.name = ������.��������
End Sub

Sub �I�𒆂̃I�u�W�F�N�g�����߂ɐݒ�()
  On Error Resume Next
  ActiveWindow.Selection.ShapeRange.name = ����.���ߖ�
End Sub

Sub �I�𒆂̃I�u�W�F�N�g���������߂������()
  On Error Resume Next
  ActiveWindow.Selection.ShapeRange.name = "NoName"
End Sub
