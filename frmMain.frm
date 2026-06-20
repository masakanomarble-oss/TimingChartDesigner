VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5715
   OleObjectBlob   =   "frmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mClkCount       As Integer
Private mUnitW_mm       As Double
Private mSigH_mm        As Double
Private mLastLoadedFile As String

Private Const LIST_NAME_WIDTH As String = "150pt"
Private Const LIST_HIDE_WIDTH As String = "0pt"


Private Sub btnClose_Click()
    Dim ans As Integer
    ans = MsgBox("設定を保存しますか？", vbYesNo + vbQuestion, "終了確認")
    
    If ans = vbYes Then
        If mClkCount = 0 Then
            MsgBox "クロック設定がありません。保存をスキップします。", vbExclamation
        Else
            Dim defName As String
            defName = IIf(mLastLoadedFile <> "", mLastLoadedFile, Visio.ActivePage.Name)
            
            Dim fName As String
            fName = InputBox("保存ファイル名を入力してください(.txtは自動付加):", "設定保存", defName)
            If fName = "" Then GoTo CloseApp
            
            Dim fPath As String: fPath = GetSettingsDir() & fName & ".txt"
            Call SaveSettingsToFile(fPath, mClkCount, mUnitW_mm, mSigH_mm)
            MsgBox "保存しました。" & vbCrLf & fPath, vbInformation
        End If
    End If
    
CloseApp:
    ThisDocument.TerminateEvents
    Unload Me
End Sub

' -------------------------------------------------------
' 初期化
' -------------------------------------------------------
Private Sub UserForm_Initialize()
    mClkCount = 0
    mUnitW_mm = 20
    mSigH_mm = 10
    mLastLoadedFile = ""

    Dim pgName   As String: pgName = Visio.ActivePage.Name
    Dim autoPath As String: autoPath = GetSettingsDir() & pgName & ".txt"

    If Dir(autoPath) <> "" Then
        Call LoadSettingsFromFile(autoPath, mClkCount, mUnitW_mm, mSigH_mm)
        mLastLoadedFile = pgName
    End If

    Me.Caption = "Timing Chart Designer"
    Me.BackColor = RGB(245, 245, 250)

    lblTitle.Caption = "Timing Chart Designer"
    lblTitle.Font.Size = 11
    lblTitle.Font.Bold = True
    lblTitle.ForeColor = RGB(50, 80, 140)

    lblSignals.Caption = "信号一覧"
    lblSignals.Font.Bold = True
    lblSignals.ForeColor = RGB(80, 80, 100)

    lstSignals.Font.Size = 10

    Call StyleButton(btnClkSetting, "クロック設定", RGB(80, 120, 200))
    Call StyleButton(btnAddSig, "+ 信号追加", RGB(60, 160, 100))
    Call StyleButton(btnDelSig, "- 削除", RGB(200, 80, 80))
    Call StyleButton(btnUp, "▲", RGB(150, 150, 170))
    Call StyleButton(btnDown, "▼", RGB(150, 150, 170))
    Call StyleButton(btnDraw, "描　画", RGB(50, 80, 140))
    Call StyleButton(btnSave, "設定保存", RGB(100, 160, 100))
    Call StyleButton(btnLoad, "設定読込", RGB(100, 160, 160))
    Call StyleButton(btnClose, "閉じる", RGB(130, 130, 140))
    Call StyleButton(btnEdit, "波形編集", RGB(160, 100, 180))

    Call RefreshClkLabel

    lstSignals.Clear
    lstSignals.ColumnCount = 2
    lstSignals.ColumnWidths = LIST_NAME_WIDTH & ";" & LIST_HIDE_WIDTH

    Call RefreshList
End Sub

' -------------------------------------------------------
' ＋ 信号追加
' -------------------------------------------------------
Private Sub btnAddSig_Click()
    Call ShowChildFormBelowMain(frmAddSignal)
End Sub

' -------------------------------------------------------
' 描画ボタン
' -------------------------------------------------------
Private Sub btnDraw_Click()
    Dim vsoPage As Visio.Page: Set vsoPage = Visio.ActivePage

    If mClkCount = 0 Then
        MsgBox "クロック設定がありません。", vbExclamation
        Exit Sub
    End If

    Dim sigCount As Integer: sigCount = frmAddSignal.GetSignalCount()
    If sigCount = 0 Then
        MsgBox "描画する信号データがありません。", vbExclamation
        Exit Sub
    End If

    ' ページ初期化（CLK描画まで済む）
    Call InitTimingChart(mClkCount, mUnitW_mm, mSigH_mm)

    Dim buf As Variant: buf = frmAddSignal.GetAddBufferArray()

    ' リストボックスの表示順（row 0 = CLKなのでrow 1から）で描画
    Dim drawIdx As Integer: drawIdx = 1
    Dim r       As Integer
    For r = 1 To lstSignals.ListCount - 1

        Dim bufIdx As Integer: bufIdx = r - 1   ' リスト順 = バッファ順
        If bufIdx >= sigCount Then GoTo NextRow

        Dim sigName   As String: sigName = buf(0, bufIdx)
        Dim sigType   As String: sigType = buf(1, bufIdx)
        Dim initLevel As String: initLevel = buf(2, bufIdx)
        Dim trans     As String: trans = buf(3, bufIdx)
        Dim bvals     As String: bvals = buf(4, bufIdx)
        Dim endClk    As String: endClk = buf(5, bufIdx)

        Dim unitW As Double: unitW = mUnitW_mm / 25.4
        Dim sigH  As Double: sigH = mSigH_mm / 25.4
        Dim drawClk As Integer
        drawClk = IIf(trans = "", 1, mClkCount)

        Select Case sigType
            Case "Digital"
                Call DrawSignalDigital(vsoPage, drawIdx, sigName, trans, _
                                       initLevel, unitW, sigH, drawClk)
            Case "Bus"
                Debug.Print "DrawSignalBus endClk=" & endClk
                Call DrawSignalBus(vsoPage, drawIdx, sigName, trans, _
                   bvals, unitW, sigH, drawClk, endClk)
        End Select

        drawIdx = drawIdx + 1
NextRow:
    Next r

    'Me.Hide
End Sub

' -------------------------------------------------------
' － 削除ボタン
' -------------------------------------------------------
Private Sub btnDelSig_Click()
    Dim idx As Integer: idx = lstSignals.ListIndex
    If idx <= 0 Then Exit Sub          ' CLK行（0）は削除不可

    Call frmAddSignal.RemoveAt(idx - 1)  ' リスト行-1 = バッファインデックス
    Call RefreshList
End Sub

' -------------------------------------------------------
' ▲ 上へ
' -------------------------------------------------------
Private Sub btnUp_Click()
    Dim idx As Integer: idx = lstSignals.ListIndex
    If idx <= 1 Then Exit Sub          ' CLK行の直下（idx=1）は上に動かせない

    ' リストボックスの表示を入れ替え
    Dim tmpName As String: tmpName = lstSignals.List(idx, 0)
    lstSignals.List(idx, 0) = lstSignals.List(idx - 1, 0)
    lstSignals.List(idx - 1, 0) = tmpName

    ' バッファ自体も入れ替え（リスト行-1 = バッファインデックス）
    Call frmAddSignal.SwapBuffer(idx - 1, idx - 2)

    lstSignals.ListIndex = idx - 1
End Sub

' -------------------------------------------------------
' ▼ 下へ
' -------------------------------------------------------
Private Sub btnDown_Click()
    Dim idx As Integer: idx = lstSignals.ListIndex
    If idx < 1 Or idx = lstSignals.ListCount - 1 Then Exit Sub

    ' リストボックスの表示を入れ替え
    Dim tmpName As String: tmpName = lstSignals.List(idx, 0)
    lstSignals.List(idx, 0) = lstSignals.List(idx + 1, 0)
    lstSignals.List(idx + 1, 0) = tmpName

    ' バッファ自体も入れ替え
    Call frmAddSignal.SwapBuffer(idx - 1, idx)

    lstSignals.ListIndex = idx + 1
End Sub

' -------------------------------------------------------
' リストボックスをバッファから再構築
' -------------------------------------------------------
Public Sub RefreshList()
    lstSignals.Clear

    Dim sigCount As Integer: sigCount = frmAddSignal.GetSignalCount()
    Dim totalRows As Integer: totalRows = 1 + sigCount

    Dim displayArray() As Variant
    ReDim displayArray(0 To totalRows - 1, 0 To 1)

    displayArray(0, 0) = "CLK"
    displayArray(0, 1) = -1

    If sigCount > 0 Then
        Dim buf As Variant: buf = frmAddSignal.GetAddBufferArray()
        Dim i As Integer
        For i = 0 To sigCount - 1
            displayArray(i + 1, 0) = buf(0, i)  ' 信号名
            displayArray(i + 1, 1) = i           ' バッファインデックス（参考用）
        Next i
    End If

    lstSignals.List = displayArray
End Sub

' -------------------------------------------------------
' クロック設定ボタン・ラベル更新
' -------------------------------------------------------
Private Sub btnClkSetting_Click()
    frmClkSetting.clkCount = mClkCount
    frmClkSetting.unitW_mm = mUnitW_mm
    frmClkSetting.sigH_mm = mSigH_mm
    Call ShowChildFormBelowMain(frmClkSetting)
    If frmClkSetting.clkCount > 0 Then
        mClkCount = frmClkSetting.clkCount
        mUnitW_mm = frmClkSetting.unitW_mm
        mSigH_mm = frmClkSetting.sigH_mm
    End If
    Call RefreshClkLabel
    Call RefreshList
End Sub

Private Sub RefreshClkLabel()
    If mClkCount > 0 Then
        lblClkInfo.Caption = "CLK: " & mClkCount & "クロック  " & mUnitW_mm & "mm  H=" & mSigH_mm & "mm"
        lblClkInfo.ForeColor = RGB(50, 80, 140)
    Else
        lblClkInfo.Caption = "CLK: 未設定"
        lblClkInfo.ForeColor = RGB(200, 80, 80)
    End If
End Sub

' -------------------------------------------------------
' 描画の土台初期化（CLKとラベルだけ描く）
' -------------------------------------------------------
Sub InitTimingChart(ByVal clkCount As Integer, _
                    ByVal unitW_mm As Double, _
                    ByVal sigH_mm As Double)
    Dim unitW   As Double:       unitW = unitW_mm / 25.4
    Dim sigH    As Double:       sigH = sigH_mm / 25.4
    Dim vsoPage As Visio.Page:   Set vsoPage = Visio.ActivePage

    Call ClearPage(vsoPage)

    Call SetPageProp(vsoPage, "ClkCount", CStr(clkCount), "クロック数")
    Call SetPageProp(vsoPage, "UnitW_mm", CStr(unitW_mm), "1クロック幅(mm)")
    Call SetPageProp(vsoPage, "SigH_mm", CStr(sigH_mm), "信号高さ(mm)")

    Call DrawClockLabels(vsoPage, clkCount, unitW)
    Call DrawSignalCLK(vsoPage, 0, clkCount, unitW, sigH)

    ThisDocument.StartEvents
End Sub

' -------------------------------------------------------
' 設定保存
' -------------------------------------------------------
Private Sub btnSave_Click()
    If mClkCount = 0 Then
        MsgBox "クロック設定がありません。", vbExclamation
        Exit Sub
    End If

    Dim defName As String
    defName = IIf(mLastLoadedFile <> "", mLastLoadedFile, Visio.ActivePage.Name)

    Dim fName As String
    fName = InputBox("保存ファイル名を入力してください(.txtは自動付加):", "設定保存", defName)
    If fName = "" Then Exit Sub

    Dim fPath As String: fPath = GetSettingsDir() & fName & ".txt"
    Call SaveSettingsToFile(fPath, mClkCount, mUnitW_mm, mSigH_mm)
    MsgBox "保存しました。" & vbCrLf & fPath, vbInformation
End Sub

' -------------------------------------------------------
' 設定読込
' -------------------------------------------------------
Private Sub btnLoad_Click()
    Dim fName As String
    fName = InputBox("読み込むファイル名を入力してください(.txtは自動付加):", "設定読込", Visio.ActivePage.Name)
    If fName = "" Then Exit Sub

    Dim fPath As String: fPath = GetSettingsDir() & fName & ".txt"

    If Dir(fPath) = "" Then
        MsgBox "ファイルが見つかりません。" & vbCrLf & fPath, vbExclamation
        Exit Sub
    End If

    Call LoadSettingsFromFile(fPath, mClkCount, mUnitW_mm, mSigH_mm)  ' fPath に修正済み
    mLastLoadedFile = fName

    Call RefreshClkLabel
    Call RefreshList
    MsgBox "読み込みました。", vbInformation
End Sub

' -------------------------------------------------------
' 編集機能
' -------------------------------------------------------
Private Sub btnEdit_Click()
    ' 編集フォームを先に開いてモードを設定してもらう
    Call ShowChildFormBelowMain(frmEditSig)
 
    ' キャンセルされた場合は何もしない
    If frmEditSig.EditMode = "" Then Exit Sub
 
    ' Visioイベントを編集モードで起動
    ' （既にStartEventsしてある前提。まだならここで呼ぶ）
    ThisDocument.StartEditMode
 
    ' フォームは閉じておく（Visio画面で操作してもらうため）
    'Me.Hide
End Sub


Public Sub TerminateEvents()   ' frmMainの閉じるボタンから呼ぶ
    Call StopEvents
    Set mWin = Nothing
End Sub

' -------------------------------------------------------
' ユーティリティ
' -------------------------------------------------------
Private Sub ShowChildFormBelowMain(ByVal child As Object)
    child.StartUpPosition = 0
    child.Left = Me.Left
    child.Top = Me.Top + Me.Height + 6
    child.Show
End Sub

Private Sub StyleButton(ByVal btn As MSForms.CommandButton, _
                         ByVal cap As String, _
                         ByVal bgColor As Long)
    btn.Caption = cap
    btn.BackColor = bgColor
    btn.ForeColor = RGB(255, 255, 255)
    btn.Font.Bold = True
    btn.Font.Size = 9
End Sub

Private Function GetSettingsDir() As String
    Dim docPath As String: docPath = Visio.ActiveDocument.FullName
    GetSettingsDir = Left(docPath, InStrRev(docPath, "\"))
End Function


