VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditSig 
   Caption         =   "UserForm1"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmEditSig.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmEditSig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -------------------------------------------------------
' frmEditSig：信号編集フォーム
' 操作モード（排他・3択）:
'   ● 挿入     : 指定位置から後ろをN CLK右シフト
'   ○ 削除     : 指定位置から後ろをN CLK左シフト
'   ○ 以降を削除: 指定位置から後ろを全削除
' -------------------------------------------------------

' 外部から参照するプロパティ
Public EditMode    As String   ' "Insert" / "Delete" / "Truncate"
Public ShiftAmount As Integer  ' シフト量（挿入・削除時のみ使用）

Private Sub UserForm_Initialize()
    Me.Caption = "波形編集"
    Me.BackColor = RGB(245, 245, 250)

    optInsert.Caption = "挿入"
    optDelete.Caption = "削除"
    optTruncate.Caption = "以降を削除"
    lblAmount.Caption = "CLK数:"
    btnDone.Caption = "実行"
    btnCancel.Caption = "キャンセル"

    Call StyleButton(btnDone, RGB(50, 80, 140))
    Call StyleButton(btnCancel, RGB(130, 130, 140))

    ' デフォルト
    optInsert.Value = True
    txtAmount.Text = "1"

    Call UpdateState
End Sub

' -------------------------------------------------------
' モード切替で入力欄を有効/無効
' -------------------------------------------------------
Private Sub optInsert_Click():   Call UpdateState: End Sub
Private Sub optDelete_Click():   Call UpdateState: End Sub
Private Sub optTruncate_Click(): Call UpdateState: End Sub

Private Sub UpdateState()
    Dim needAmount As Boolean
    needAmount = (optInsert.Value Or optDelete.Value)
    txtAmount.enabled = needAmount
    lblAmount.ForeColor = IIf(needAmount, RGB(80, 80, 100), RGB(180, 180, 180))
End Sub

' -------------------------------------------------------
' 実行ボタン
' -------------------------------------------------------
Private Sub btnDone_Click()
    If optInsert.Value Or optDelete.Value Then
        If Not IsNumeric(txtAmount.Text) Or CInt(txtAmount.Text) <= 0 Then
            MsgBox "CLK数は1以上の整数を入力してください。", vbExclamation
            Exit Sub
        End If
        ShiftAmount = CInt(txtAmount.Text)
        EditMode = IIf(optInsert.Value, "Insert", "Delete")
    Else
        EditMode = "Truncate"
        ShiftAmount = 0
    End If
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    EditMode = ""
    ShiftAmount = 0
    Me.Hide
End Sub

' -------------------------------------------------------
' ボタンスタイル
' -------------------------------------------------------
Private Sub StyleButton(ByVal btn As MSForms.CommandButton, _
                         ByVal bgColor As Long)
    btn.BackColor = bgColor
    btn.ForeColor = RGB(255, 255, 255)
    btn.Font.Bold = True
    btn.Font.Size = 9
End Sub

