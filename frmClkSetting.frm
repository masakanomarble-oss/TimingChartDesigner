VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClkSetting 
   Caption         =   "frmClkSetting"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4155
   OleObjectBlob   =   "frmClkSetting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmClkSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 設定値をプロパティとして公開
Public clkCount As Integer
Public unitW_mm As Double
Public sigH_mm  As Double

Private Sub UserForm_Initialize()
    Me.Caption = "クロック設定"
    Me.BackColor = RGB(245, 245, 250)

    ' デフォルト値
    txtClkCount.Text = "8"
    txtUnitW.Text = "20"
    txtSigH.Text = "10"

    ' ラベル
    lblClk.Caption = "クロック数:"
    lblUnitW.Caption = "1クロック幅 (mm):"
    lblSigH.Caption = "信号高さ (mm):"

    ' ボタン
    btnOK.Caption = "OK"
    btnOK.BackColor = RGB(80, 120, 200)
    btnOK.ForeColor = RGB(255, 255, 255)
    btnOK.Font.Bold = True

    btnCancel.Caption = "キャンセル"
    btnCancel.BackColor = RGB(130, 130, 140)
    btnCancel.ForeColor = RGB(255, 255, 255)
    btnCancel.Font.Bold = True

    ' 既存値があれば反映
    If clkCount > 0 Then
        txtClkCount.Text = CStr(clkCount)
        txtUnitW.Text = CStr(unitW_mm)
        txtSigH.Text = CStr(sigH_mm)
    End If
    
    Me.txtClkCount.SetFocus
    
End Sub

Private Sub btnOK_Click()
    On Error GoTo InvalidInput
    Dim c As Integer: c = CInt(txtClkCount.Text)
    Dim u As Double: u = CDbl(txtUnitW.Text)
    Dim h As Double: h = CDbl(txtSigH.Text)
    On Error GoTo 0

    If c < 1 Or c > 64 Then
        MsgBox "クロック数は1?64で指定してください。", vbExclamation: Exit Sub
    End If
    If u < 5 Or u > 100 Then
        MsgBox "幅は5?100mmで指定してください。", vbExclamation: Exit Sub
    End If
    If h < 3 Or h > 50 Then
        MsgBox "高さは3?50mmで指定してください。", vbExclamation: Exit Sub
    End If

    clkCount = c
    unitW_mm = u
    sigH_mm = h
    Me.Hide
    Exit Sub

InvalidInput:
    MsgBox "数値を入力してください。", vbExclamation
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

