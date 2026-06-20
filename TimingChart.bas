Attribute VB_Name = "TimingChart"
' =============================================
' Visio Timing Chart Generator
' =============================================
Option Explicit

Public Const TOP_MARGIN As Double = 1.5
Public Const LEFT_MARGIN As Double = 1.5
Public Const LABEL_H    As Double = 0.15

'座標基準をひとまとめに計算する関数
Private Function GetBounds(ByVal sigIdx As Integer, ByVal sigH As Double, _
                           ByRef outYHi As Double, ByRef outYLo As Double, ByRef outXStart As Double)
    ' 縦の計算（インチ）
    Dim pageH As Double: pageH = ActivePage.PageSheet.Cells("PageHeight").Result("Inches")
    Dim sigPitch As Double: sigPitch = sigH * 1.5
    outYHi = (pageH - TOP_MARGIN) - sigIdx * sigPitch
    outYLo = outYHi - sigH
    
    ' 横の計算（インチ）
    outXStart = LEFT_MARGIN
End Function

' -------------------------------------------------------
' エントリポイント
' -------------------------------------------------------
Sub NewTimingChart()
    Call frmAddSignal.ClearBuffer
    frmMain.StartUpPosition = 0
    frmMain.Left = 10
    frmMain.Top = 100
    frmMain.Show
End Sub

' -------------------------------------------------------
' 設定保存（図形のプロパティから直接保存）
' -------------------------------------------------------
Sub SaveSettingsToFile(ByVal fPath As String, _
                       ByVal clkCount As Integer, _
                       ByVal unitW_mm As Double, _
                       ByVal sigH_mm As Double)
    Dim fNum As Integer: fNum = FreeFile
    
    Open fPath For Output As #fNum
    
    Print #fNum, "CLK|" & clkCount & "|" & unitW_mm & "|" & sigH_mm
    
    Dim buf As Variant: buf = frmAddSignal.GetAddBufferArray()
    Dim i As Integer
    For i = 0 To frmAddSignal.GetSignalCount() - 1
        Print #fNum, buf(0, i) & "|" & buf(1, i) & "|" & buf(2, i) & "|" & buf(3, i) & "|" & buf(4, i) & "|" & buf(5, i)
    Next i
    
    Close #fNum
End Sub

' -------------------------------------------------------
' ② 読込：テキストから読み込んで、一旦すべてバッファに保存する
' -------------------------------------------------------
' -------------------------------------------------------
Sub LoadSettingsFromFile(ByVal fPath As String, _
                         ByRef clkCount As Integer, _
                         ByRef unitW_mm As Double, _
                         ByRef sigH_mm As Double)
                         
    If Dir(fPath) = "" Then Exit Sub

    Dim fNum As Integer: fNum = FreeFile
    Open fPath For Input As #fNum
    Dim lineStr As String
    
    ' 最初に追加画面のバッファをきれいにクリアしておく
    Call frmAddSignal.ClearBuffer

    Do While Not EOF(fNum)
        Line Input #fNum, lineStr
        lineStr = Trim(lineStr)
        
        '空行じゃなく、かつ「|」がちゃんと含まれている行だけ処理する！
        If lineStr <> "" And InStr(lineStr, "|") > 0 Then
            Dim d() As String: d = Split(lineStr, "|")
            
            '配列 d がちゃんと作られているかチェック
            If UBound(d) >= 0 Then
                If d(0) = "CLK" Then
                    ' CLK設定を読み込む（最低3つデータがあるかチェック）
                    If UBound(d) >= 3 Then
                        clkCount = CInt(d(1))
                        unitW_mm = CDbl(d(2))
                        sigH_mm = CDbl(d(3))
                    End If
                Else
                ' 信号データ: sName|sigType|initLevel|transitions|busValues
                    If UBound(d) >= 2 Then
                        Dim sigName  As String: sigName = d(0)
                        Dim sigType  As String: sigType = d(1)
                        Dim initLevel As String: initLevel = d(2)
                        Dim trans    As String: trans = IIf(UBound(d) >= 3, d(3), "")
                        Debug.Print "UBound(d)=" & UBound(d)
                        Dim bvals    As String: bvals = IIf(UBound(d) >= 4, d(4), "")
                        Dim endClk As String: endClk = IIf(UBound(d) >= 5, d(5), "")
                        Call frmAddSignal.AppendToBuffer(sigName, sigType, initLevel)
                        If trans <> "" Or bvals <> "" Then
                            Call frmAddSignal.UpdateTransitions(sigName, trans, bvals, endClk)
                            Debug.Print "loaded: sigName=" & sigName & " trans=" & trans & " bvals=" & bvals
                            Debug.Print "loaded: sigName=" & sigName & " trans=" & trans & " bvals=" & bvals & " endClk=" & endClk
                        End If
                    End If
                End If
            End If
        End If
    Loop
    Close #fNum
End Sub

' -------------------------------------------------------
' ページ初期化・再描画
' -------------------------------------------------------
Sub InitTimingChart(ByVal clkCount As Integer, _
                    ByVal unitW_mm As Double, _
                    ByVal sigH_mm As Double, _
                    ByVal signals As Collection)
                    
    Dim unitW As Double: unitW = unitW_mm / 25.4
    Dim sigH  As Double: sigH = sigH_mm / 25.4
    Dim vsoPage As Visio.Page: Set vsoPage = Visio.ActivePage

    ' 既存図形を全削除
    Call ClearPage(vsoPage)

    ' ページプロパティ保存
    Call SetPageProp(vsoPage, "ClkCount", CStr(clkCount), "クロック数")
    Call SetPageProp(vsoPage, "UnitW_mm", CStr(unitW_mm), "1クロック幅(mm)")
    Call SetPageProp(vsoPage, "SigH_mm", CStr(sigH_mm), "信号高さ(mm)")
    ' 信号数は後で数えるので今は仮
    Call SetPageProp(vsoPage, "SigCount", CStr(frmAddSignal.GetSignalCount()), "信号数")
    
    ' クロック番号・CLK描画
    Call DrawClockLabels(vsoPage, clkCount, unitW)
    Call DrawSignalCLK(vsoPage, 0, clkCount, unitW, sigH)

    ' ----------------------------------------------------
    ' 信号描画
    ' ----------------------------------------------------
    Dim i As Integer
    For i = 1 To signals.Count
        Dim sig As Variant: sig = signals(i)
        Select Case sig(1)
            Case "Digital"
                Call DrawSignalDigital(vsoPage, i, sig(0), sig(3), sig(2), unitW, sigH, clkCount, sig(5))
            Case "Bus"
                Call DrawSignalBus(vsoPage, i, sig(0), sig(3), sig(4), unitW, sigH, clkCount, sig(5))
        End Select
    Next i

    ' 設定保存
    Dim savePath As String
    savePath = Left(Visio.ActiveDocument.FullName, InStrRev(Visio.ActiveDocument.FullName, "\")) & vsoPage.Name & ".txt"
    Call SaveSettingsToFile(savePath, clkCount, unitW_mm, sigH_mm)

    ThisDocument.StartEvents
End Sub

' -------------------------------------------------------
' ページ全図形削除
' -------------------------------------------------------
Public Sub ClearPage(ByVal pg As Visio.Page)
    Dim shp As Visio.Shape
    Dim delList() As Visio.Shape
    Dim cnt As Integer: cnt = 0
    ReDim delList(pg.Shapes.Count)
    For Each shp In pg.Shapes
        Set delList(cnt) = shp
        cnt = cnt + 1
    Next shp
    Dim i As Integer
    For i = 0 To cnt - 1
        On Error Resume Next
        delList(i).Delete
        On Error GoTo 0
    Next i
End Sub

' -------------------------------------------------------
' クロック番号ラベル描画
' -------------------------------------------------------
Public Sub DrawClockLabels(ByVal pg As Visio.Page, _
                             ByVal clkCount As Integer, _
                             ByVal unitW As Double)
    ' ?? 共通関数から、1行目の基準の高さ(yHi)と横(xStart)をゲット！
    Dim yHi As Double, yLo As Double, xStart As Double
    Call GetBounds(0, 0, yHi, yLo, xStart) ' sigHは0でOK

    Dim li  As Integer
    Dim shp As Visio.Shape
    For li = 0 To clkCount - 1
        ' ?? yHi から LABEL_H（0.15インチ）分、上に飛び出させて四角を描く！
        Set shp = pg.DrawRectangle( _
            xStart + li * unitW, yHi + 0.15, _
            xStart + (li + 1) * unitW, yHi)
        shp.Text = CStr(li)
        shp.CellsU("LinePattern").FormulaU = "0"
        shp.CellsU("VerticalAlign").FormulaU = "1"
        shp.CellsU("Para.HorzAlign").FormulaU = "0"
    Next li
End Sub

' -------------------------------------------------------
' CLK信号描画
' -------------------------------------------------------
Public Sub DrawSignalCLK(ByVal pg As Visio.Page, _
                           ByVal sigIdx As Integer, _
                           ByVal clkCount As Integer, _
                           ByVal unitW As Double, _
                           ByVal sigH As Double)
    
    Dim sigPitch As Double: sigPitch = sigH * 1.5
    Dim yHi As Double
    Dim yLo As Double
    Dim xStart As Double
    
    Call GetBounds(sigIdx, sigH, yHi, yLo, xStart)
    'Dim yHi As Double: yHi = TOP_MARGIN - sigIdx * sigPitch
    'Dim yLo As Double: yLo = yHi - sigH

    Dim shp As Visio.Shape
    Set shp = pg.DrawRectangle(0, yHi, xStart, yLo)
    shp.Text = "CLK"
    shp.CellsU("LinePattern").FormulaU = "0"
    shp.CellsU("VerticalAlign").FormulaU = "1"
    shp.CellsU("Para.HorzAlign").FormulaU = "0"
    Call SetShapeProp(shp, "SigName", "CLK", "信号名")
    Call SetShapeProp(shp, "SigType", "CLK", "種別")
    Call SetShapeProp(shp, "ClkIndex", "0", "信号行番号")
    Call SetShapeProp(shp, "Transitions", "", "遷移クロック列")
    Call SetShapeProp(shp, "BusValues", "", "Bus値ラベル")

    Dim pts() As Double
    ReDim pts(clkCount * 8 + 3)
    Dim idx As Integer: idx = 0
    pts(idx) = xStart: idx = idx + 1
    pts(idx) = yLo: idx = idx + 1

    Dim ci As Integer
    For ci = 0 To clkCount - 1
        Dim xRise As Double: xRise = xStart + ci * unitW
        Dim xFall As Double: xFall = xStart + (ci + 0.5) * unitW
        pts(idx) = xRise: idx = idx + 1
        pts(idx) = yLo: idx = idx + 1
        pts(idx) = xRise: idx = idx + 1
        pts(idx) = yHi: idx = idx + 1
        pts(idx) = xFall: idx = idx + 1
        pts(idx) = yHi: idx = idx + 1
        pts(idx) = xFall: idx = idx + 1
        pts(idx) = yLo: idx = idx + 1
    Next ci
    pts(idx) = xStart + clkCount * unitW: idx = idx + 1
    pts(idx) = yLo: idx = idx + 1
    pg.DrawPolyline pts, 0
End Sub

' -------------------------------------------------------
' Digital信号描画
' -------------------------------------------------------
Public Function DrawSignalDigital(ByVal pg As Visio.Page, _
                                    ByVal sigIdx As Integer, _
                                    ByVal sigName As String, _
                                    ByVal transitions As String, _
                                    ByVal initLevel As String, _
                                    ByVal unitW As Double, _
                                    ByVal sigH As Double, _
                                    ByVal clkCount As Integer, _
                                    Optional ByVal endClk As String = "") As Visio.Shape
    Dim sigPitch As Double: sigPitch = sigH * 1.5
    Dim yHi As Double
    Dim yLo As Double
    Dim xStart As Double
    
    Call GetBounds(sigIdx, sigH, yHi, yLo, xStart)

    'Dim yHi As Double: yHi = TOP_MARGIN - sigIdx * sigPitch
    'Dim yLo As Double: yLo = yHi - sigH

    Dim shp As Visio.Shape
    Set shp = pg.DrawRectangle(0, yHi, xStart, yLo)
    shp.Text = sigName
    shp.CellsU("LinePattern").FormulaU = "0"
    shp.CellsU("VerticalAlign").FormulaU = "1"
    shp.CellsU("Para.HorzAlign").FormulaU = "0"
    Call SetShapeProp(shp, "SigName", sigName, "信号名")
    Call SetShapeProp(shp, "SigType", "Digital", "種別")
    Call SetShapeProp(shp, "ClkIndex", CStr(sigIdx), "信号行番号")
    Call SetShapeProp(shp, "Transitions", transitions, "遷移クロック列")
    Call SetShapeProp(shp, "BusValues", "", "Bus値ラベル")
    Call SetShapeProp(shp, "InitLevel", initLevel, "初期レベル(0=Lo,1=Hi)")
    Call SetShapeProp(shp, "EndClk", endClk, "End CLK")

    Dim flatY As Double
    If initLevel = "1" Then flatY = yHi Else flatY = yLo

    ' 遷移が空 → 1クロック分だけ描く
    If transitions = "" Then
        Dim flatDrawEnd As Double
        If Trim(endClk) <> "" And IsNumeric(Trim(endClk)) Then
            flatDrawEnd = CDbl(Trim(endClk))
        Else
            flatDrawEnd = clkCount
        End If
        If flatDrawEnd > clkCount Then flatDrawEnd = clkCount

        Dim flatPts(3) As Double
        flatPts(0) = xStart: flatPts(1) = flatY
        flatPts(2) = xStart + flatDrawEnd * unitW: flatPts(3) = flatY
        
        Dim wshpF As Visio.Shape
        Set wshpF = pg.DrawPolyline(flatPts, 0)
        Call SetShapeProp(wshpF, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
        Exit Function
    End If

    ' 遷移リストから波形を再構築
    Dim toks() As String: toks = Split(transitions, ",")
    Dim n As Integer: n = UBound(toks) + 1
    Dim pts() As Double
    ReDim pts((n * 2 + 2) * 2 - 1)

    Dim pidx  As Integer: pidx = 0
    Dim level As Double: level = flatY

    pts(pidx) = xStart: pidx = pidx + 1
    pts(pidx) = level: pidx = pidx + 1

    Dim ti As Integer
    For ti = 0 To n - 1
        'もし中身が空っぽ（ゴミデータ）なら、スキップして次へ！
        If IsNumeric(Trim(toks(ti))) Then
            Dim xTrans As Double: xTrans = xStart + CDbl(toks(ti)) * unitW
            pts(pidx) = xTrans: pidx = pidx + 1
            pts(pidx) = level: pidx = pidx + 1
            If level = yHi Then level = yLo Else level = yHi
                pts(pidx) = xTrans: pidx = pidx + 1
                pts(pidx) = level: pidx = pidx + 1
        End If
    Next ti

    ' 末端：endClkがあればそこまで。なければ最後の遷移点+1クロックで止める
    Dim drawEndClk As Double
    If Trim(endClk) <> "" And IsNumeric(Trim(endClk)) Then
        drawEndClk = CDbl(Trim(endClk))
    ElseIf IsNumeric(Trim(toks(n - 1))) Then
        drawEndClk = CDbl(Trim(toks(n - 1))) + 1
    Else
        drawEndClk = 1
    End If
    If drawEndClk > clkCount Then drawEndClk = clkCount
    pts(pidx) = xStart + drawEndClk * unitW: pidx = pidx + 1
    pts(pidx) = level: pidx = pidx + 1
    
    ReDim Preserve pts(pidx - 1)
    Dim wshpD As Visio.Shape
    Set wshpD = pg.DrawPolyline(pts, 0)
    Call SetShapeProp(wshpD, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
End Function

' -------------------------------------------------------
' Bus信号描画
' -------------------------------------------------------
Public Function DrawSignalBus(ByVal pg As Visio.Page, _
                           ByVal sigIdx As Integer, _
                           ByVal sigName As String, _
                           ByVal transitions As String, _
                           ByVal busValues As String, _
                           ByVal unitW As Double, _
                           ByVal sigH As Double, _
                           ByVal clkCount As Integer, _
                           ByVal endClk As String) As Visio.Shape
    Const BUS_GAP As Double = 0.01
    Dim sigPitch As Double: sigPitch = sigH * 1.5
    Dim yHi As Double
    Dim yLo As Double
    Dim xStart As Double
    
    Call GetBounds(sigIdx, sigH, yHi, yLo, xStart)
    Debug.Print "transitions=" & transitions & " busValues=" & busValues

    Dim shp As Visio.Shape
    Set shp = pg.DrawRectangle(0, yHi, xStart, yLo)
    shp.Text = sigName
    shp.CellsU("LinePattern").FormulaU = "0"
    shp.CellsU("VerticalAlign").FormulaU = "1"
    shp.CellsU("Para.HorzAlign").FormulaU = "0"
    Call SetShapeProp(shp, "SigName", sigName, "信号名")
    Call SetShapeProp(shp, "SigType", "Bus", "種別")
    Call SetShapeProp(shp, "ClkIndex", CStr(sigIdx), "信号行番号")
    Call SetShapeProp(shp, "Transitions", transitions, "遷移クロック列")
    Call SetShapeProp(shp, "BusValues", busValues, "Bus値ラベル")
    Call SetShapeProp(shp, "EndClk", endClk, "End CLK")

    ' 遷移が空 → 1クロック分の=を描く
    If transitions = "" Or busValues = "" Then
        Dim ex1 As Double: ex1 = xStart + BUS_GAP
        Dim ex2 As Double: ex2 = xStart + clkCount * unitW - BUS_GAP
        Dim eqTop(3) As Double
        eqTop(0) = ex1: eqTop(1) = yHi
        eqTop(2) = ex2: eqTop(3) = yHi
        Dim eqTopShp As Visio.Shape
        Set eqTopShp = pg.DrawPolyline(eqTop, 0)
        Call SetShapeProp(eqTopShp, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
        Dim eqBot(3) As Double
        eqBot(0) = ex1: eqBot(1) = yLo
        eqBot(2) = ex2: eqBot(3) = yLo
        Dim eqBotShp As Visio.Shape
        Set eqBotShp = pg.DrawPolyline(eqBot, 0)
        Call SetShapeProp(eqBotShp, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
        Exit Function
    End If

    ' 上下ライン全体
    Dim toks()   As String: toks = Split(transitions, ",")
    Dim labels() As String: labels = Split(busValues, ",")
    Dim segCount As Integer: segCount = UBound(labels) + 1

    Dim bi As Integer
    ' ① xStartから最初のXまで（=だけ、Xなし）
    Dim initPts(3) As Double
    initPts(0) = xStart: initPts(1) = yHi
    initPts(2) = xStart + CDbl(toks(0)) * unitW: initPts(3) = yHi
    Dim initT As Visio.Shape: Set initT = pg.DrawPolyline(initPts, 0)
    Call SetShapeProp(initT, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
    initPts(1) = yLo: initPts(3) = yLo
    Dim initB As Visio.Shape: Set initB = pg.DrawPolyline(initPts, 0)
    Call SetShapeProp(initB, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")

    ' ② 各区間（X＋=＋テキスト）
    For bi = 0 To segCount - 1
        Dim startClk As Double: startClk = CDbl(toks(bi))
        Dim endClkSeg As Double
        If bi < segCount - 1 Then
            endClkSeg = CDbl(toks(bi + 1))
        ElseIf Trim(endClk) <> "" And IsNumeric(Trim(endClk)) Then
            endClkSeg = CDbl(Trim(endClk))
        Else
            endClkSeg = startClk + 1  ' 最後は1CLK分だけ
        End If
        If endClkSeg > clkCount Then endClkSeg = clkCount

        ' Xマーク
        Dim xMid  As Double: xMid = xStart + startClk * unitW
        Dim xSize As Double: xSize = BUS_GAP * 3
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

        ' 2本水平線＋テキスト
        Dim bx1 As Double: bx1 = xStart + startClk * unitW + BUS_GAP
        Dim bx2 As Double: bx2 = xStart + endClkSeg * unitW - BUS_GAP
        Dim topPts(3) As Double
        topPts(0) = bx1: topPts(1) = yHi
        topPts(2) = bx2: topPts(3) = yHi
        Dim topSeg As Visio.Shape: Set topSeg = pg.DrawPolyline(topPts, 0)
        Call SetShapeProp(topSeg, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
        Dim botPts(3) As Double
        botPts(0) = bx1: botPts(1) = yLo
        botPts(2) = bx2: botPts(3) = yLo
        Dim botSeg As Visio.Shape: Set botSeg = pg.DrawPolyline(botPts, 0)
        Call SetShapeProp(botSeg, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
        Dim txtBox As Visio.Shape
        Set txtBox = pg.DrawRectangle(bx1, yHi, bx2, yLo)
        txtBox.Text = labels(bi)
        txtBox.CellsU("LinePattern").FormulaU = "0"
        txtBox.CellsU("FillPattern").FormulaU = "0"
        txtBox.CellsU("VerticalAlign").FormulaU = "1"
        txtBox.CellsU("Para.HorzAlign").FormulaU = "1"
        Call SetShapeProp(txtBox, "WaveSigIdx", CStr(sigIdx), "波形対応信号インデックス")
    Next bi

End Function

' -------------------------------------------------------
' ShapeSheetカスタムプロパティ設定
' -------------------------------------------------------
Public Sub SetShapeProp(ByVal shp As Visio.Shape, _
                         ByVal propName As String, _
                         ByVal propValue As String, _
                         ByVal propLabel As String)
    If Not shp.CellExistsU("Prop." & propName, visExistsAnywhere) Then
        shp.AddNamedRow visSectionProp, propName, visTagDefault
    End If
    shp.CellsU("Prop." & propName).FormulaU = """" & propValue & """"
    shp.CellsU("Prop." & propName & ".Label").FormulaU = """" & propLabel & """"
End Sub

' -------------------------------------------------------
' ページプロパティ設定
' -------------------------------------------------------
Public Sub SetPageProp(ByVal pg As Visio.Page, _
                        ByVal propName As String, _
                        ByVal propValue As String, _
                        ByVal propLabel As String)
    Dim ps As Visio.Shape: Set ps = pg.PageSheet
    If Not ps.CellExistsU("Prop." & propName, visExistsAnywhere) Then
        ps.AddNamedRow visSectionProp, propName, visTagDefault
    End If
    ps.CellsU("Prop." & propName).FormulaU = """" & propValue & """"
    ps.CellsU("Prop." & propName & ".Label").FormulaU = """" & propLabel & """"
End Sub

' -------------------------------------------------------
' ページプロパティ取得
' -------------------------------------------------------
Public Function GetPageProp(ByVal pg As Visio.Page, _
                              ByVal propName As String) As String
    Dim ps As Visio.Shape: Set ps = pg.PageSheet
    If ps.CellExistsU("Prop." & propName, visExistsAnywhere) Then
        GetPageProp = ps.CellsU("Prop." & propName).ResultStr("")
    Else
        GetPageProp = ""
    End If
End Function


