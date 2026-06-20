VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddSignal 
   Caption         =   "frmAddSignal"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   OleObjectBlob   =   "frmAddSignal.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmAddSignal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' =======================================================
' 画面：frmAddSignal の宣言部とデータ入力窓口
' =======================================================
'
' バッファ列の定義（定数で管理）
'   COL_NAME  (0) : 信号名
'   COL_TYPE  (1) : sigType  ("Digital" or "Bus")
'   COL_INIT  (2) : initLevel ("0" or "1")
'   COL_TRANS (3) : 遷移データ（AppendDigital/Busで更新）
'   COL_BVAL  (4) : Bus値データ（AppendBusで更新）
'
Private Const COL_NAME  As Integer = 0
Private Const COL_TYPE  As Integer = 1
Private Const COL_INIT  As Integer = 2
Private Const COL_TRANS As Integer = 3
Private Const COL_BVAL  As Integer = 4
Private Const COL_END   As Integer = 5
Private Const COL_MAX   As Integer = 5  ' 最大列インデックス

Private mAddBuffer()   As String
Private mSignalCount   As Integer


' -------------------------------------------------------
' フォーム初期化
' -------------------------------------------------------
Private Sub UserForm_Initialize()
    Me.Caption = "信号追加"
    Me.BackColor = RGB(245, 245, 250)

    lblName.Caption = "信号名:"
    lblType.Caption = "種別:"
    lblLevel.Caption = "初期レベル:"

    txtSigName.Text = ""
    optDigital.Value = True
    optLo.Value = True

    optDigital.Caption = "Digital"
    optBus.Caption = "Bus"
    optHi.Caption = "Hi"
    optLo.Caption = "Lo"

    btnOK.Caption = "OK"
    btnOK.BackColor = RGB(60, 160, 100)
    btnOK.ForeColor = RGB(255, 255, 255)
    btnOK.Font.Bold = True

    btnCancel.Caption = "キャンセル"
    btnCancel.BackColor = RGB(130, 130, 140)
    btnCancel.ForeColor = RGB(255, 255, 255)
    btnCancel.Font.Bold = True
    
    Call UpdateLevelState
    Me.txtSigName.SetFocus

End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub optLo_Click()
End Sub

Private Sub optDigital_Click()
    Call UpdateLevelState
End Sub

Private Sub optBus_Click()
    Call UpdateLevelState
End Sub

Private Sub UpdateLevelState()
    Dim enabled As Boolean: enabled = optDigital.Value
    optHi.enabled = enabled
    optLo.enabled = enabled
    lblLevel.ForeColor = IIf(enabled, RGB(80, 80, 100), RGB(180, 180, 180))
End Sub

' -------------------------------------------------------
' OKボタン：バッファに追加してリスト更新
' -------------------------------------------------------
Private Sub btnOK_Click()
    If Trim(txtSigName.Text) = "" Then
        MsgBox "信号名を入力してください。", vbExclamation
        Exit Sub
    End If

    Dim sName     As String: sName = Trim(txtSigName.Text)
    Dim sigType   As String
    Dim initLevel As String

    If optDigital.Value Then
        sigType = "Digital"
        initLevel = IIf(optHi.Value, "1", "0")
    Else
        sigType = "Bus"
        initLevel = "0"
    End If

    Call Me.AppendToBuffer(sName, sigType, initLevel)
    Call frmMain.RefreshList

    txtSigName.Text = ""
    optDigital.Value = True
    optLo.Value = True
    Me.Hide
End Sub


' =======================================================
' 公開インターフェース
' =======================================================

' バッファ全体を返す（Variant で2次元配列ごと渡す）
Public Function GetAddBufferArray() As Variant
    GetAddBufferArray = mAddBuffer
End Function

' 件数を返す
Public Function GetSignalCount() As Integer
    GetSignalCount = mSignalCount
End Function

' 信号を1件追加する
Public Sub AppendToBuffer(ByVal sigName As String, _
                           ByVal sigType As String, _
                           ByVal initLevel As String)
    sigName = Trim(Replace(Replace(sigName, vbCr, ""), vbLf, ""))
    sigType = Trim(Replace(Replace(sigType, vbCr, ""), vbLf, ""))
    initLevel = Trim(Replace(Replace(initLevel, vbCr, ""), vbLf, ""))

    If sigName = "" Then Exit Sub

    mSignalCount = mSignalCount + 1
    ReDim Preserve mAddBuffer(COL_MAX, mSignalCount - 1)

    mAddBuffer(COL_NAME, mSignalCount - 1) = sigName
    mAddBuffer(COL_TYPE, mSignalCount - 1) = sigType
    mAddBuffer(COL_INIT, mSignalCount - 1) = initLevel
    mAddBuffer(COL_TRANS, mSignalCount - 1) = ""   ' 遷移データは初期空
    mAddBuffer(COL_BVAL, mSignalCount - 1) = ""    ' Bus値も初期空
    mAddBuffer(COL_END, mSignalCount - 1) = ""    '終了位置
End Sub


' 遷移データを更新する（AppendDigital / AppendBus から呼ぶ）
Public Sub UpdateTransitions(ByVal sigName As String, _
                               ByVal newTrans As String, _
                               ByVal newBvals As String, _
                               ByVal endClk As String)
    Dim i As Integer
    For i = 0 To mSignalCount - 1
        If mAddBuffer(COL_NAME, i) = sigName Then
            mAddBuffer(COL_TRANS, i) = newTrans
            mAddBuffer(COL_BVAL, i) = newBvals
            mAddBuffer(COL_END, i) = endClk
            Debug.Print "UpdateTransitions COL_END=" & mAddBuffer(COL_END, i)
            Exit For
        End If
    Next i
End Sub

' バッファ内の2要素を入れ替える（▲▼ボタンから呼ぶ）
Public Sub SwapBuffer(ByVal idxA As Integer, ByVal idxB As Integer)
    If idxA < 0 Or idxA >= mSignalCount Then Exit Sub
    If idxB < 0 Or idxB >= mSignalCount Then Exit Sub

    Dim c As Integer
    Dim tmp As String
    For c = 0 To COL_MAX
        tmp = mAddBuffer(c, idxA)
        mAddBuffer(c, idxA) = mAddBuffer(c, idxB)
        mAddBuffer(c, idxB) = tmp
    Next c
End Sub

' バッファを全クリア
Public Sub ClearBuffer()
    ReDim mAddBuffer(COL_MAX, 0)
    Erase mAddBuffer
    mSignalCount = 0
End Sub

' 指定インデックスの信号を削除（前詰め）
Public Sub RemoveAt(ByVal idx As Integer)
    If idx < 0 Or idx >= mSignalCount Then Exit Sub

    Dim i As Integer, c As Integer
    For i = idx To mSignalCount - 2
        For c = 0 To COL_MAX
            mAddBuffer(c, i) = mAddBuffer(c, i + 1)
        Next c
    Next i

    mSignalCount = mSignalCount - 1
    If mSignalCount = 0 Then
        Call ClearBuffer
    Else
        ReDim Preserve mAddBuffer(COL_MAX, mSignalCount - 1)
    End If
End Sub

' -------------------------------------------------------
' 編集機能
' -------------------------------------------------------

' 指定信号の遷移データをシフトする
' sigName  : 対象信号名
' fromClk  : このCLK位置以降を操作
' amount   : シフト量（正=右シフト/挿入、負=左シフト/削除）
Public Sub ShiftTransitions(ByVal sigName As String, _
                             ByVal fromClk As Integer, _
                             ByVal amount As Integer)
    Dim i As Integer
    For i = 0 To mSignalCount - 1
        If mAddBuffer(COL_NAME, i) = sigName Then

            Dim trans As String: trans = mAddBuffer(COL_TRANS, i)
            Dim bvals As String: bvals = mAddBuffer(COL_BVAL, i)

            If trans = "" Then Exit Sub   ' 遷移なしなら何もしない

            Dim toks()  As String: toks = Split(trans, ",")
            Dim bToks() As String
            Dim hasBus  As Boolean: hasBus = (bvals <> "")
            If hasBus Then bToks = Split(bvals, ",")

            Dim newTrans As String: newTrans = ""
            Dim newBvals As String: newBvals = ""
            Dim j        As Integer
            Dim bIdx     As Integer: bIdx = 0   ' Bus値インデックス（遷移2個で1Bus値）

            For j = 0 To UBound(toks)
                Dim clk As Integer: clk = CInt(Trim(toks(j)))

                If clk >= fromClk Then
                    clk = clk + amount

                    ' 削除シフト時にfromClk未満になったら除去
                    If clk < fromClk Then
                        ' Bus値も除去（遷移2個で1Bus値なのでペアで捨てる）
                        If hasBus And (j Mod 2 = 0) Then bIdx = bIdx + 1
                        GoTo NextTok
                    End If
                End If

                ' 結果に追加
                newTrans = newTrans & IIf(newTrans = "", "", ",") & CStr(clk)

                ' Bus値は遷移2個ごとに1つ対応
                If hasBus And j Mod 2 = 1 Then
                    If bIdx <= UBound(bToks) Then
                        newBvals = newBvals & IIf(newBvals = "", "", ",") & bToks(bIdx)
                        bIdx = bIdx + 1
                    End If
                End If

NextTok:
            Next j

            mAddBuffer(COL_TRANS, i) = newTrans
            mAddBuffer(COL_BVAL, i) = newBvals
            Exit For
        End If
    Next i
End Sub

' 指定信号の指定CLK以降の遷移データを全削除する
' sigName : 対象信号名
' fromClk : このCLK位置以降を削除（fromClk自身も含む）
Public Sub TruncateTransitions(ByVal sigName As String, _
                                ByVal fromClk As Integer)
    Dim i As Integer
    For i = 0 To mSignalCount - 1
        If mAddBuffer(COL_NAME, i) = sigName Then

            Dim trans As String: trans = mAddBuffer(COL_TRANS, i)
            Dim bvals As String: bvals = mAddBuffer(COL_BVAL, i)

            If trans = "" Then
                mAddBuffer(COL_BVAL, i) = ""
                mAddBuffer(COL_END, i) = CStr(fromClk)
                Exit Sub
            End If

            Dim toks()  As String: toks = Split(trans, ",")
            Dim hasBus  As Boolean: hasBus = (bvals <> "")
            Dim bToks() As String
            If hasBus Then bToks = Split(bvals, ",")

            Dim newTrans As String: newTrans = ""
            Dim newBvals As String: newBvals = ""
            Dim j   As Integer
            Dim bIdx As Integer: bIdx = 0

            For j = 0 To UBound(toks)
                Dim clk As Integer: clk = CInt(Trim(toks(j)))
                If clk >= fromClk Then GoTo TruncNext  ' 以降は全部捨てる

                newTrans = newTrans & IIf(newTrans = "", "", ",") & CStr(clk)

                If hasBus Then
                    If j <= UBound(bToks) Then
                        newBvals = newBvals & IIf(newBvals = "", "", ",") & bToks(j)
                    End If
                End If
TruncNext:
            Next j

            mAddBuffer(COL_TRANS, i) = newTrans
            mAddBuffer(COL_BVAL, i) = newBvals
            mAddBuffer(COL_END, i) = CStr(fromClk)
            Exit For
        End If
    Next i
End Sub

