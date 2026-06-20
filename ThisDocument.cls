VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private mSelectedSig As Visio.Shape
Private mGuideShape  As Visio.Shape
Private mLastSnapClk As Double

Private Const HL_FILL  As String = "RGB(180,210,255)"
Private Const HL_TEXT  As String = "RGB(0,0,0)"
Private Const NM_FILL  As String = "RGB(255,255,255)"
Private Const NM_TEXT  As String = "RGB(0,0,0)"
Private Const SNAP_TOL As Double = 0.4

Private mUnitW    As Double
Private mSigH     As Double
Private mClkCount As Integer
Private mSigCount As Integer

Private mBusPrevClk As Double
Private mBusPrevVal As String
Private mBusDrawing As Boolean

Private WithEvents mWin As Visio.Window
Attribute mWin.VB_VarHelpID = -1
Private mEditMode As Boolean   ' 編集モードフラグ


' -------------------------------------------------------
' 縦横座標計算（ローカル版）
' -------------------------------------------------------
Private Sub GetLocalBounds(ByVal pg As Visio.Page, _
                            ByVal sigIdx As Integer, _
                            ByVal sigH As Double, _
                            ByRef outYHi As Double, _
                            ByRef outYLo As Double, _
                            ByRef outLW As Double)
    Dim pageH    As Double: pageH = pg.PageSheet.Cells("PageHeight").Result("Inches")
    Dim sigPitch As Double: sigPitch = sigH * 1.5

    outYHi = (pageH - TOP_MARGIN) - sigIdx * sigPitch
    outYLo = outYHi - sigH
    outLW = LEFT_MARGIN
End Sub

' -------------------------------------------------------
' イベント登録
' -------------------------------------------------------
Public Sub StartEvents()
    Set mWin = Visio.ActiveWindow
    Dim pg As Visio.Page: Set pg = Visio.ActivePage
 
    mUnitW = CDbl(GetPageProp(pg, "UnitW_mm")) / 25.4
    mSigH = CDbl(GetPageProp(pg, "SigH_mm")) / 25.4
    mClkCount = CInt(GetPageProp(pg, "ClkCount"))
    mSigCount = frmAddSignal.GetSignalCount()
    
    mLastSnapClk = -1
    mBusDrawing = False
    mEditMode = False
End Sub

Public Sub StartEditMode()
    If mWin Is Nothing Then Call StartEvents
    mEditMode = True
    mBusDrawing = False
End Sub


' -------------------------------------------------------
' イベント終了
' -------------------------------------------------------
Public Sub StopEvents()
    Debug.Print "StopEvents called"
    Call RemoveGuide
    If Not mSelectedSig Is Nothing Then
        Call ResetHighlight(mSelectedSig)
        Set mSelectedSig = Nothing
    End If
    mBusDrawing = False
    mEditMode = False
    ' Set mWin = Nothing  ← 削除！mWinは生かしたまま
End Sub
 

Public Sub TerminateEvents()   ' frmMainの閉じるボタンから呼ぶ
    Call StopEvents
    Set mWin = Nothing
End Sub

Public Sub StopEditMode()
Debug.Print "StopEditMode called"
    mEditMode = False
End Sub


' -------------------------------------------------------
' 選択変更イベント
' -------------------------------------------------------
Private Sub mWin_SelectionChanged(ByVal Window As Visio.Window)

    If Window.Selection.Count = 0 Then Exit Sub

    Dim shp As Visio.Shape: Set shp = Window.Selection.Item(1)
Debug.Print "selected shape: " & shp.NameU & " hasSigType=" & shp.CellExistsU("Prop.SigType", visExistsAnywhere)
    If Not shp.CellExistsU("Prop.SigType", visExistsAnywhere) Then Exit Sub

    Dim sigType As String: sigType = shp.CellsU("Prop.SigType").ResultStr("")
    If sigType = "CLK" Then Exit Sub

    If Not mSelectedSig Is Nothing Then
        If mSelectedSig.ID = shp.ID Then Exit Sub
        Call ResetHighlight(mSelectedSig)
        Set mSelectedSig = Nothing
        mBusDrawing = False
    End If

    Set mSelectedSig = shp
    Call ApplyHighlight(shp)
End Sub

' -------------------------------------------------------
' マウス移動：ガイド縦線
' -------------------------------------------------------
Private Sub mWin_MouseMove(ByVal Button As Long, _
                            ByVal KeyButtonState As Long, _
                            ByVal x As Double, ByVal y As Double, _
                            CancelDefault As Boolean)
    Dim pg As Visio.Page: Set pg = Visio.ActivePage
    If mUnitW = 0 Or mClkCount = 0 Then Exit Sub

    Dim yTop As Double, dummyYLo As Double, LW As Double
    Call GetLocalBounds(pg, 0, 0, yTop, dummyYLo, LW)

    Dim snapClk As Double: snapClk = -1
    Dim snapX   As Double
    Dim ci As Integer
    
    For ci = 0 To mClkCount - 1
        Dim candX As Double: candX = LW + ci * mUnitW + mUnitW / 2  ' 中央で判定
        If Abs(x - candX) < mUnitW / 2 Then
            snapClk = ci
            snapX = LW + ci * mUnitW      ' ← ガイド線は左端境界に表示
            Exit For
        End If
    Next ci

    If snapClk = mLastSnapClk Then Exit Sub

    Call RemoveGuide
    mLastSnapClk = snapClk
    If snapClk < 0 Then Exit Sub

    yTop = yTop + 0.15
    Dim yBot As Double: yBot = yTop - (mSigCount + 1) * mSigH * 1.5
    
    Set mGuideShape = pg.DrawLine(snapX, yTop, snapX, yBot)
    mGuideShape.CellsU("LineColor").FormulaU = "RGB(0,120,220)"
    mGuideShape.CellsU("LinePattern").FormulaU = "2"
    mGuideShape.CellsU("LineWeight").FormulaU = "1pt"
    mGuideShape.CellsU("LockSelect").FormulaU = "TRUE"
End Sub

' -------------------------------------------------------
' マウスクリック：信号追記
' -------------------------------------------------------
Private Sub mWin_MouseDown(ByVal Button As Long, _
                            ByVal KeyButtonState As Long, _
                            ByVal x As Double, ByVal y As Double, _
                            CancelDefault As Boolean)
Debug.Print "右クリック mLastSnapClk=" & mLastSnapClk & " mBusDrawing=" & mBusDrawing

    If Button = 2 Then
        If mEditMode Then
            Call StopEditMode
        ElseIf Not mSelectedSig Is Nothing Then
            Dim pg2 As Visio.Page: Set pg2 = Visio.ActivePage
            Dim sigType2 As String: sigType2 = mSelectedSig.CellsU("Prop.SigType").ResultStr("")
            Select Case sigType2
                Case "Bus"
                    If mBusDrawing Then
                        Call ExtendBusToEnd(pg2, mSelectedSig, CInt(mLastSnapClk), mUnitW, mSigH)
                    End If
                Case "Digital"
                    Call ExtendDigitalToEnd(pg2, mSelectedSig, CInt(mLastSnapClk), mUnitW, mSigH)
            End Select
        End If
        Exit Sub
    End If
    
    If Button <> 1 Then Exit Sub
    If mSelectedSig Is Nothing Then Exit Sub
    If mLastSnapClk < 0 Then Exit Sub
 
    Dim pg As Visio.Page: Set pg = Visio.ActivePage
     
    If mEditMode Then
        Call EditSignal(pg, mSelectedSig, CInt(mLastSnapClk))
    Else
        Dim sigType As String: sigType = mSelectedSig.CellsU("Prop.SigType").ResultStr("")
        Select Case sigType
            Case "Digital": Call AppendDigital(pg, mSelectedSig, mLastSnapClk, mUnitW, mSigH)
            Case "Bus":     Call AppendBus(pg, mSelectedSig, mLastSnapClk, mUnitW, mSigH)
        End Select
    End If
End Sub


 
' -------------------------------------------------------
' Digital信号の追記
' -------------------------------------------------------
Private Sub AppendDigital(ByVal pg As Visio.Page, _
                           ByVal sigShp As Visio.Shape, _
                           ByVal clkPos As Double, _
                           ByVal unitW As Double, _
                           ByVal sigH As Double)
                           
    Dim sigIdx As Integer: sigIdx = CInt(sigShp.CellsU("Prop.ClkIndex").ResultStr(""))

    Dim yHi As Double, yLo As Double, LW As Double
    Call GetLocalBounds(pg, sigIdx, sigH, yHi, yLo, LW)

    Dim trans     As String: trans = sigShp.CellsU("Prop.Transitions").ResultStr("")
    Dim initLevel As String: initLevel = sigShp.CellsU("Prop.InitLevel").ResultStr("")

    Dim endClk   As Double
    Dim endLevel As Double
    Call GetDigitalEnd(trans, initLevel, yHi, yLo, endClk, endLevel)

    If clkPos < endClk Then Exit Sub

    Debug.Print "clkPos=" & clkPos & " endClk=" & endClk
    Dim newLevel As Double
    newLevel = IIf(endLevel = yHi, yLo, yHi)

    Dim pts(7) As Double
    pts(0) = LW + endClk * unitW: pts(1) = endLevel
    pts(2) = LW + clkPos * unitW: pts(3) = endLevel
    pts(4) = LW + clkPos * unitW: pts(5) = newLevel
    pts(6) = LW + (clkPos + 1) * unitW: pts(7) = newLevel

    Dim wshp As Visio.Shape: Set wshp = pg.DrawPolyline(pts, 0)
    Call SetShapeProp(wshp, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")

    ' Shape の Prop.Transitions を更新
    Dim newTrans As String
    newTrans = IIf(trans = "", CStr(CLng(clkPos)), trans & "," & CStr(CLng(clkPos)))
    sigShp.CellsU("Prop.Transitions").FormulaU = """" & newTrans & """"

    ' ★ バッファも同時更新（並び替え後の再描画に備える）
    Dim sigName As String: sigName = sigShp.CellsU("Prop.SigName").ResultStrU("")
    Call frmAddSignal.UpdateTransitions(sigName, newTrans, "", "")
End Sub

' -------------------------------------------------------
' Digital末端クロックとレベルを取得
' -------------------------------------------------------
Private Sub GetDigitalEnd(ByVal trans As String, _
                           ByVal initLevel As String, _
                           ByVal yHi As Double, _
                           ByVal yLo As Double, _
                           ByRef endClk As Double, _
                           ByRef endLevel As Double)
    endLevel = IIf(initLevel = "1", yHi, yLo)

    If trans = "" Then
        endClk = 0
        Exit Sub
    End If

    Dim toks() As String: toks = Split(trans, ",")

    endClk = IIf(IsNumeric(Trim(toks(UBound(toks)))), _
                 CDbl(Trim(toks(UBound(toks)))) + 1, 1)

    Dim toggleCount As Integer: toggleCount = UBound(toks) + 1
    If toggleCount Mod 2 = 1 Then
        endLevel = IIf(endLevel = yHi, yLo, yHi)
    End If
End Sub

' -------------------------------------------------------
' Bus信号の追記
' -------------------------------------------------------
Private Sub AppendBus(ByVal pg As Visio.Page, _
                       ByVal sigShp As Visio.Shape, _
                       ByVal clkPos As Double, _
                       ByVal unitW As Double, _
                       ByVal sigH As Double)
    Const BUS_GAP As Double = 0.02
 
    Dim sigIdx As Integer: sigIdx = CInt(sigShp.CellsU("Prop.ClkIndex").ResultStr(""))
 
    Dim yHi As Double, yLo As Double, LW As Double
    Call GetLocalBounds(pg, sigIdx, sigH, yHi, yLo, LW)
 
    Dim trans As String: trans = sigShp.CellsU("Prop.Transitions").ResultStrU("")
    Dim bvals As String: bvals = sigShp.CellsU("Prop.BusValues").ResultStrU("")
 
    ' 末端クロックを取得
    Dim endClk As Double: endClk = GetBusEnd(trans)
 
    ' クリック位置が末端以前なら無視
    If clkPos < endClk Then Exit Sub
 
    ' Bus値を入力
    Dim busVal As String
    busVal = InputBox("CLK" & CLng(clkPos) & " のBus値:", "Bus値入力", "")
    
    Dim extPts(3) As Double
    extPts(0) = LW + endClk * unitW: extPts(1) = yHi
    extPts(2) = LW + clkPos * unitW: extPts(3) = yHi
    Dim extTop As Visio.Shape: Set extTop = pg.DrawPolyline(extPts, 0)
    Call SetShapeProp(extTop, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")

    extPts(1) = yLo: extPts(3) = yLo
    Dim extBot As Visio.Shape: Set extBot = pg.DrawPolyline(extPts, 0)
    Call SetShapeProp(extBot, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
    
    ' X（変化記号）を描く
    Const xSize As Double = BUS_GAP * 3
    Dim xMid As Double: xMid = LW + clkPos * unitW
 
    Dim xPts1(3) As Double
    xPts1(0) = xMid - xSize: xPts1(1) = yHi
    xPts1(2) = xMid + xSize: xPts1(3) = yLo
    Dim xShp1 As Visio.Shape: Set xShp1 = pg.DrawPolyline(xPts1, 0)
    Call SetShapeProp(xShp1, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
 
    Dim xPts2(3) As Double
    xPts2(0) = xMid - xSize: xPts2(1) = yLo
    xPts2(2) = xMid + xSize: xPts2(3) = yHi
    Dim xShp2 As Visio.Shape: Set xShp2 = pg.DrawPolyline(xPts2, 0)
    Call SetShapeProp(xShp2, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
 
    ' 1クロック分の =（2本水平線）
    Dim ex1 As Double: ex1 = LW + clkPos * unitW + BUS_GAP
    Dim ex2 As Double: ex2 = LW + (clkPos + 1) * unitW - BUS_GAP
 
    Dim eqTop(3) As Double
    eqTop(0) = ex1: eqTop(1) = yHi
    eqTop(2) = ex2: eqTop(3) = yHi
    Dim eqTopShp As Visio.Shape: Set eqTopShp = pg.DrawPolyline(eqTop, 0)
    Call SetShapeProp(eqTopShp, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
 
    Dim eqBot(3) As Double
    eqBot(0) = ex1: eqBot(1) = yLo
    eqBot(2) = ex2: eqBot(3) = yLo
    Dim eqBotShp As Visio.Shape: Set eqBotShp = pg.DrawPolyline(eqBot, 0)
    Call SetShapeProp(eqBotShp, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
 
    ' 値テキスト
    Dim txtBox As Visio.Shape: Set txtBox = pg.DrawRectangle(ex1, yHi, ex2, yLo)
    txtBox.Text = busVal
    txtBox.CellsU("LinePattern").FormulaU = "0"
    txtBox.CellsU("FillPattern").FormulaU = "0"
    txtBox.CellsU("VerticalAlign").FormulaU = "1"
    txtBox.CellsU("Para.HorzAlign").FormulaU = "1"
    Call SetShapeProp(txtBox, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
 
    ' Prop.Transitions / BusValues 更新
    Dim newTrans As String
    Dim newBvals As String
    If trans = "" Then
        newTrans = CStr(CLng(clkPos))
        newBvals = busVal
    Else
        newTrans = trans & "," & CStr(CLng(clkPos))
        newBvals = bvals & "," & busVal
    End If
    sigShp.CellsU("Prop.Transitions").FormulaU = """" & newTrans & """"
    sigShp.CellsU("Prop.BusValues").FormulaU = """" & newBvals & """"
 
    ' バッファも同時更新
    Dim sigName As String: sigName = sigShp.CellsU("Prop.SigName").ResultStrU("")
    Call frmAddSignal.UpdateTransitions(sigName, newTrans, newBvals, "")
 
    ' 末端クロックを更新（次のクリックの起点）
    mBusPrevClk = clkPos + 1
    mBusPrevVal = busVal
    mBusDrawing = True
    
    ' バッファ更新（末端CLKは未確定なので空）
    Call frmAddSignal.UpdateTransitions(sigName, newTrans, newBvals, "")
End Sub

' -------------------------------------------------------
' Bus末端クロックを取得
' -------------------------------------------------------
Private Function GetBusEnd(ByVal trans As String) As Double
    If trans = "" Then
        GetBusEnd = 0
        Exit Function
    End If
    Dim toks() As String: toks = Split(trans, ",")
    GetBusEnd = CDbl(Trim(toks(UBound(toks)))) + 1
End Function

' -------------------------------------------------------
' 波形図形のみ削除
' -------------------------------------------------------
Private Sub DeleteWaveShapes(ByVal pg As Visio.Page, ByVal sigIdx As Integer)
    Dim shp     As Visio.Shape
    Dim delList() As Visio.Shape
    Dim cnt     As Integer: cnt = 0

    ReDim delList(pg.Shapes.Count)
    For Each shp In pg.Shapes
        If shp.CellExistsU("Prop.WaveSigIdx", visExistsAnywhere) Then
            If CInt(shp.CellsU("Prop.WaveSigIdx").ResultStr("")) = sigIdx Then
                Set delList(cnt) = shp
                cnt = cnt + 1
            End If
        End If
    Next shp

    Dim i As Integer
    For i = 0 To cnt - 1
        delList(i).Delete
    Next i
End Sub

' -------------------------------------------------------
' ハイライト適用 / リセット
' -------------------------------------------------------
Private Sub ApplyHighlight(ByVal shp As Visio.Shape)
    shp.CellsU("FillForegnd").FormulaU = HL_FILL
    shp.CellsU("FillPattern").FormulaU = "1"
    shp.CellsU("Char.Color").FormulaU = HL_TEXT
    shp.CellsU("Char.Style").FormulaU = "1"
End Sub

Private Sub ResetHighlight(ByVal shp As Visio.Shape)
    On Error Resume Next
    shp.CellsU("FillForegnd").FormulaU = NM_FILL
    shp.CellsU("FillPattern").FormulaU = "0"
    shp.CellsU("Char.Color").FormulaU = NM_TEXT
    shp.CellsU("Char.Style").FormulaU = "0"
    On Error GoTo 0
End Sub

' -------------------------------------------------------
' ガイド点線削除
' -------------------------------------------------------
Private Sub RemoveGuide()
    If Not mGuideShape Is Nothing Then
        On Error Resume Next
        mGuideShape.CellsU("LockSelect").FormulaU = "FALSE"
        mGuideShape.Delete
        On Error GoTo 0
        Set mGuideShape = Nothing
    End If
    mLastSnapClk = -1
End Sub

' -------------------------------------------------------
' 数値文字列配列をバブルソート
' -------------------------------------------------------
Private Sub SortStrAsNum(ByRef arr() As String, ByVal n As Integer)
    Dim i As Integer, j As Integer, tmp As String
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If CDbl(arr(j)) > CDbl(arr(j + 1)) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

' -------------------------------------------------------
' 信号の編集機能
' -------------------------------------------------------
Private Sub EditSignal(ByVal pg As Visio.Page, _
                        ByVal sigShp As Visio.Shape, _
                        ByVal clickClk As Integer)
 
    Dim mode   As String:  mode = frmEditSig.EditMode
    Dim amount As Integer: amount = frmEditSig.ShiftAmount
 
    If mode = "" Then
        Debug.Print "mode empty! exiting"
        Exit Sub
    End If
 
    If mode = "" Then Exit Sub
    
    ' ★追加
    Dim buf2 As Variant: buf2 = frmAddSignal.GetAddBufferArray()
    Dim ii As Integer

    Dim sigName As String: sigName = sigShp.CellsU("Prop.SigName").ResultStrU("")
 
    Select Case mode
        Case "Insert"
            Call frmAddSignal.ShiftTransitions(sigName, clickClk, amount)
        Case "Delete"
            Call frmAddSignal.ShiftTransitions(sigName, clickClk, -amount)
        Case "Truncate"
            Call frmAddSignal.TruncateTransitions(sigName, clickClk)
            If sigShp.CellExistsU("Prop.EndClk", visExistsAnywhere) Then
                sigShp.CellsU("Prop.EndClk").FormulaU = """" & CStr(clickClk) & """"
            End If
                
            ' ★追加
            Dim buf3 As Variant: buf3 = frmAddSignal.GetAddBufferArray()
    End Select
 
    ' 該当信号の波形Shapeを削除して再描画
    Dim sigIdx As Integer: sigIdx = CInt(sigShp.CellsU("Prop.ClkIndex").ResultStr(""))
    Call DeleteWaveShapes(pg, sigIdx)
 
    Dim buf As Variant: buf = frmAddSignal.GetAddBufferArray()
    Dim i   As Integer

    For i = 0 To frmAddSignal.GetSignalCount() - 1
        If buf(0, i) = sigName Then
            Dim drawClk As Integer
            drawClk = IIf(buf(3, i) = "", clickClk, mClkCount)
            Dim endClk As String: endClk = buf(5, i)

            Select Case buf(1, i)
                Case "Digital"
                    Call DrawSignalDigital(pg, sigIdx, sigName, buf(3, i), _
                               buf(2, i), mUnitW, mSigH, mClkCount, buf(5, i))
                Case "Bus"
                    Call DrawSignalBus(pg, sigIdx, sigName, buf(3, i), _
                       buf(4, i), mUnitW, mSigH, drawClk, endClk)
            End Select
            Exit For
        End If
    Next i
End Sub



' -------------------------------------------------------
' 【2】ExtendBusToEnd を新規追加（右クリックで末端延長）
' -------------------------------------------------------
 
Private Sub ExtendBusToEnd(ByVal pg As Visio.Page, _
                            ByVal sigShp As Visio.Shape, _
                            ByVal endClk As Integer, _
                            ByVal unitW As Double, _
                            ByVal sigH As Double)
    Debug.Print "ExtendBusToEnd called mBusPrevClk=" & mBusPrevClk & " endClk=" & endClk
    Const BUS_GAP As Double = 0.02
    
 
    Dim sigIdx As Integer: sigIdx = CInt(sigShp.CellsU("Prop.ClkIndex").ResultStr(""))
 
    Dim yHi As Double, yLo As Double, LW As Double
    Call GetLocalBounds(pg, sigIdx, sigH, yHi, yLo, LW)
 
    ' 最後の値を末端まで延長（上下ライン）
    ' 上下ラインを探して EndX を伸ばす
    Dim shp As Visio.Shape
    ' 上下ラインを追加
    Dim extPts(3) As Double
    extPts(0) = LW + mBusPrevClk * unitW: extPts(1) = yHi
    extPts(2) = LW + endClk * unitW:             extPts(3) = yHi
    Dim extTop As Visio.Shape: Set extTop = pg.DrawPolyline(extPts, 0)
    Call SetShapeProp(extTop, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")

    extPts(1) = yLo: extPts(3) = yLo
    Dim extBot As Visio.Shape: Set extBot = pg.DrawPolyline(extPts, 0)
    Call SetShapeProp(extBot, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
  
    ' Prop.Transitions / BusValues 更新
    Dim trans  As String: trans = sigShp.CellsU("Prop.Transitions").ResultStrU("")
    Dim bvals  As String: bvals = sigShp.CellsU("Prop.BusValues").ResultStrU("")
    Dim newTrans As String: newTrans = trans
    Dim newBvals As String: newBvals = bvals  ' 値は追加しない（延長なので）
 
    sigShp.CellsU("Prop.Transitions").FormulaU = """" & newTrans & """"
    sigShp.CellsU("Prop.BusValues").FormulaU = """" & newBvals & """"
 
    Dim sigName As String: sigName = sigShp.CellsU("Prop.SigName").ResultStrU("")
    Call frmAddSignal.UpdateTransitions(sigName, newTrans, newBvals, "")
    
    Debug.Print "UpdateTransitions endClk=" & CStr(endClk)
 
    ' 描画終了
    mBusDrawing = False
    mBusPrevVal = ""
    mBusPrevClk = 0
    pg.Application.ActiveWindow.DeselectAll
    
    ' バッファ更新（末端CLKを確定保存）
    Call frmAddSignal.UpdateTransitions(sigName, newTrans, newBvals, CStr(endClk))
End Sub
 

Private Sub ExtendDigitalToEnd(ByVal pg As Visio.Page, _
                                ByVal sigShp As Visio.Shape, _
                                ByVal endClk As Integer, _
                                ByVal unitW As Double, _
                                ByVal sigH As Double)
    Dim sigIdx As Integer: sigIdx = CInt(sigShp.CellsU("Prop.ClkIndex").ResultStr(""))

    Dim yHi As Double, yLo As Double, LW As Double
    Call GetLocalBounds(pg, sigIdx, sigH, yHi, yLo, LW)

    Dim trans     As String: trans = sigShp.CellsU("Prop.Transitions").ResultStrU("")
    Dim initLevel As String: initLevel = sigShp.CellsU("Prop.InitLevel").ResultStrU("")

    ' 現在の末端クロックと末端レベルを取得
    Dim curEndClk   As Double
    Dim curEndLevel As Double
    Call GetDigitalEnd(trans, initLevel, yHi, yLo, curEndClk, curEndLevel)

    If endClk <= curEndClk Then Exit Sub

    ' 末端レベルのまま endClk まで伸ばす
    Dim pts(3) As Double
    pts(0) = LW + curEndClk * unitW: pts(1) = curEndLevel
    pts(2) = LW + endClk * unitW: pts(3) = curEndLevel

    Dim wshp As Visio.Shape: Set wshp = pg.DrawPolyline(pts, 0)
    Call SetShapeProp(wshp, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")

    ' Digitalも右クリック延長した終端CLKを保存する
    On Error Resume Next
    sigShp.CellsU("Prop.EndClk").FormulaU = """ & CStr(endClk) & """
    On Error GoTo 0

    Dim sigName As String: sigName = sigShp.CellsU("Prop.SigName").ResultStrU("")
    Call frmAddSignal.UpdateTransitions(sigName, trans, "", CStr(endClk))
End Sub



