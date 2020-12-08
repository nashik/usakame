VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RabbitAndTurtle 
   Caption         =   "RabbitAndTurtle"
   ClientHeight    =   2100
   ClientLeft      =   50
   ClientTop       =   370
   ClientWidth     =   3890
   OleObjectBlob   =   "RabbitAndTurtle.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RabbitAndTurtle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub StartButton_Click()
    Dim totalSec As Integer
    Dim startTime

    totalSec = MinutesSpinButton.Value * 60 + SecondSpinButton.Value
    startTime = Time
    
    RabbitAndTurtle.Hide
    
    Call うさかめ整列.うさかめ整列
    Call うさぎ.うさぎ実行
    Call かめ.かめ実行(totalSec, startTime)
    
End Sub

Private Sub MinutesSpinButton_Change()
    With MinutesSpinButton
        .SmallChange = 1
    End With
    MinutesBox.Value = MinutesSpinButton.Value
End Sub

Private Sub SecondSpinButton_Change()
    With SecondSpinButton
        .SmallChange = 10
        .Max = 50
    End With
    SecondBox.Value = SecondSpinButton.Value
End Sub

