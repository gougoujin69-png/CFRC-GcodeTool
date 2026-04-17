Attribute VB_Name = "CFRC_GcodeExporter"
'==============================================================================
' CFRC G-code Exporter for Excel
'   作者: Albert (冯镜泽) 定制版
'   功能: 读取活动表 A:D 四列(X, Y, Z, Cut)生成 G-code
'         A=X(mm)  B=Y(mm)  C=Z(mm)  D=剪丝标记(1=在该点剪丝, 留空=否)
'   输出: 仅 G1 段 + Header(用户自定义) + 剪丝宏
'   挤出: 相对模式 (M83), E = 段长 × W × H / (π·(D/2)²)
'==============================================================================
Option Explicit

' ---------- 默认参数 (可通过对话框改, 改完会记忆到 Names) ----------
Private Const DEF_WIDTH       As Double = 0.4      ' 线宽 mm
Private Const DEF_HEIGHT      As Double = 0.2      ' 层高 mm
Private Const DEF_DIAMETER    As Double = 1.75     ' 丝材直径 mm
Private Const DEF_F_PRINT     As Double = 1800#    ' 打印进给 mm/min
Private Const DEF_F_TRAVEL    As Double = 3000#    ' 空驶进给 mm/min
Private Const DEF_RETRACT     As Double = 0#       ' 剪丝前回抽 mm (0=不回抽)
Private Const DEF_CUT_MACRO   As String = "CUT_FIBER"   ' 剪丝宏命令
Private Const DEF_HAS_HEADER  As Boolean = True    ' 第1行是否表头
Private Const DEF_Z_LIFT_THR  As Double = 0.05     ' |dZ|>该值视为层间抬升(空驶) mm

'------------------------------------------------------------------------------
' 入口: 功能区按钮点击 / Alt+F8 调用
'------------------------------------------------------------------------------
Public Sub ExportGcodeFromSheet()
    On Error GoTo ErrH

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "请先打开一个工作表", vbExclamation: Exit Sub
    End If

    ' 1. 读取参数 (带记忆)
    Dim p As GcodeParams
    If Not LoadParamsDialog(p) Then Exit Sub

    ' 2. 读取数据 A:D
    Dim startRow As Long
    startRow = IIf(p.HasHeader, 2, 1)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then
        MsgBox "A 列没有数据(从第 " & startRow & " 行起)", vbExclamation: Exit Sub
    End If

    Dim nPts As Long: nPts = lastRow - startRow + 1
    If nPts < 2 Then
        MsgBox "至少需要 2 个点", vbExclamation: Exit Sub
    End If

    Dim arr As Variant
    arr = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, 4)).Value

    ' 3. 选择保存路径
    Dim savePath As String
    savePath = AskSavePath(ws.Name)
    If Len(savePath) = 0 Then Exit Sub

    ' 4. 生成 G-code
    Dim gcode As String
    gcode = BuildGcode(arr, nPts, p)

    ' 5. 写文件 (UTF-8 无 BOM)
    WriteUtf8 savePath, gcode

    MsgBox "已导出: " & vbCrLf & savePath & vbCrLf & vbCrLf & _
           "数据点: " & nPts & "    剪丝次数: " & p.CutCount & vbCrLf & _
           "总挤出长度 E ≈ " & Format(p.TotalE, "0.000") & " mm" & vbCrLf & _
           "总路径长度 ≈ " & Format(p.TotalLen, "0.000") & " mm", vbInformation, "导出完成"
    Exit Sub

ErrH:
    MsgBox "错误: " & Err.Description, vbCritical, "CFRC G-code Exporter"
End Sub

'------------------------------------------------------------------------------
' 核心: 构造 G-code
'------------------------------------------------------------------------------
Private Function BuildGcode(arr As Variant, nPts As Long, p As GcodeParams) As String
    Dim sb As Object
    Set sb = CreateObject("System.Collections.ArrayList")

    ' ---- Header ----
    Dim hdrLines() As String
    hdrLines = Split(p.Header, vbLf)
    Dim i As Long
    For i = LBound(hdrLines) To UBound(hdrLines)
        sb.Add Replace(hdrLines(i), vbCr, "")
    Next i

    ' ---- 起点: 第一行 ----
    Dim x1 As Double, y1 As Double, z1 As Double
    x1 = CDbl(arr(1, 1)): y1 = CDbl(arr(1, 2)): z1 = CDbl(arr(1, 3))

    sb.Add ""
    sb.Add "; ===== Toolpath start (from XLSX) ====="
    sb.Add "G92 E0                      ; 重置挤出"
    sb.Add "G1 X" & F(x1) & " Y" & F(y1) & " Z" & F(z1) & " F" & F(p.FTravel) & " ; 移动到起点"

    ' E 系数: e_per_mm = W*H / (π·(D/2)²)
    Dim ePerMm As Double
    ePerMm = (p.LineWidth * p.LayerHeight) / (3.14159265358979 * (p.FilamentDia / 2#) ^ 2)

    Dim travelNext As Boolean: travelNext = False  ' 上一点剪丝则下一段空驶
    Dim totalE As Double: totalE = 0#
    Dim totalLen As Double: totalLen = 0#
    Dim cutCount As Long: cutCount = 0
    Dim curLayerZ As Double: curLayerZ = z1
    Dim layerIdx As Long: layerIdx = 1

    ' ---- 是否第一行就剪丝(罕见但允许) ----
    If IsCutCell(arr(1, 4)) Then
        sb.Add "; ---- Cut at start point ----"
        EmitCut sb, p
        cutCount = cutCount + 1
        travelNext = True
    End If

    ' ---- 主循环 ----
    Dim xPrev As Double, yPrev As Double, zPrev As Double
    xPrev = x1: yPrev = y1: zPrev = z1

    Dim r As Long
    For r = 2 To nPts
        Dim xc As Double, yc As Double, zc As Double, cutFlag As Boolean
        xc = CDbl(arr(r, 1)): yc = CDbl(arr(r, 2)): zc = CDbl(arr(r, 3))
        cutFlag = IsCutCell(arr(r, 4))

        Dim dx As Double, dy As Double, dz As Double, segLen As Double
        dx = xc - xPrev: dy = yc - yPrev: dz = zc - zPrev
        segLen = Sqr(dx * dx + dy * dy + dz * dz)

        Dim isLayerLift As Boolean
        isLayerLift = (Abs(dz) > p.ZLiftThreshold)

        ' 换层注释
        If isLayerLift And zc <> curLayerZ Then
            layerIdx = layerIdx + 1
            sb.Add ""
            sb.Add "; ---- Layer " & layerIdx & "  Z=" & F(zc) & " ----"
            curLayerZ = zc
        End If

        Dim asTravel As Boolean
        asTravel = travelNext Or isLayerLift Or (segLen <= 0.000001)

        If asTravel Then
            ' 空驶 / 抬升
            sb.Add "G1 X" & F(xc) & " Y" & F(yc) & " Z" & F(zc) & _
                   " F" & F(p.FTravel) & IIf(travelNext, " ; travel after cut", IIf(isLayerLift, " ; layer move", ""))
        Else
            ' 挤出
            Dim de As Double: de = segLen * ePerMm
            totalE = totalE + de
            totalLen = totalLen + segLen
            If Abs(dz) < 0.000001 Then
                sb.Add "G1 X" & F(xc) & " Y" & F(yc) & _
                       " E" & F(de) & " F" & F(p.FPrint)
            Else
                sb.Add "G1 X" & F(xc) & " Y" & F(yc) & " Z" & F(zc) & _
                       " E" & F(de) & " F" & F(p.FPrint)
            End If
        End If

        travelNext = False

        ' 该点剪丝
        If cutFlag Then
            sb.Add "; ---- Cut at point " & r & " ----"
            EmitCut sb, p
            cutCount = cutCount + 1
            travelNext = True
        End If

        xPrev = xc: yPrev = yc: zPrev = zc
    Next r

    sb.Add ""
    sb.Add "; ===== End of toolpath ====="

    ' 回填统计
    p.TotalE = totalE
    p.TotalLen = totalLen
    p.CutCount = cutCount

    ' 拼字符串
    Dim outArr() As String
    ReDim outArr(0 To sb.Count - 1)
    Dim k As Long
    For k = 0 To sb.Count - 1
        outArr(k) = CStr(sb(k))
    Next k
    BuildGcode = Join(outArr, vbCrLf) & vbCrLf
End Function

Private Sub EmitCut(sb As Object, p As GcodeParams)
    If p.RetractLen > 0 Then
        sb.Add "G1 E-" & F(p.RetractLen) & " F" & F(p.FTravel) & " ; retract before cut"
    End If
    sb.Add p.CutMacro
    sb.Add "G92 E0                      ; 剪丝后重置 E"
End Sub

'------------------------------------------------------------------------------
' 工具函数
'------------------------------------------------------------------------------
Private Function F(v As Double) As String
    ' 5 位有效精度, 去掉无意义的尾零
    Dim s As String
    s = Format(v, "0.#####")
    If s = "" Or s = "-" Then s = "0"
    F = s
End Function

Private Function IsCutCell(v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then IsCutCell = False: Exit Function
    Dim s As String: s = Trim(CStr(v))
    If Len(s) = 0 Then IsCutCell = False: Exit Function
    If IsNumeric(s) Then
        IsCutCell = (CDbl(s) <> 0)
    Else
        IsCutCell = (UCase(s) = "Y" Or UCase(s) = "YES" Or UCase(s) = "TRUE" Or UCase(s) = "CUT")
    End If
End Function

Private Function AskSavePath(defName As String) As String
    Dim sfd As FileDialog
    Set sfd = Application.FileDialog(msoFileDialogSaveAs)
    With sfd
        .Title = "保存 G-code 文件"
        .InitialFileName = defName & ".gcode"
        .FilterIndex = 1
        If .Show = -1 Then
            AskSavePath = .SelectedItems(1)
        Else
            AskSavePath = ""
        End If
    End With
End Function

Private Sub WriteUtf8(path As String, content As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2          ' Text
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content
    ' 去 BOM
    Dim binStm As Object
    Set binStm = CreateObject("ADODB.Stream")
    binStm.Type = 1
    binStm.Mode = 3
    binStm.Open
    stm.Position = 3      ' 跳过 UTF-8 BOM 三字节
    stm.CopyTo binStm
    stm.Flush
    stm.Close
    binStm.SaveToFile path, 2
    binStm.Close
End Sub

'==============================================================================
' 参数对话框 (用 InputBox 链, 简单可靠, 不依赖 UserForm)
'   参数会以 Names 形式记忆在加载项里, 下次打开默认值就是上次的输入
'==============================================================================
Public Type GcodeParams
    LineWidth      As Double
    LayerHeight    As Double
    FilamentDia    As Double
    FPrint         As Double
    FTravel        As Double
    RetractLen     As Double
    CutMacro       As String
    Header         As String
    HasHeader      As Boolean
    ZLiftThreshold As Double
    ' 输出统计
    TotalE         As Double
    TotalLen       As Double
    CutCount       As Long
End Type

Private Function LoadParamsDialog(ByRef p As GcodeParams) As Boolean
    LoadParamsDialog = False

    p.LineWidth      = GetStored("CFRC_W",       DEF_WIDTH)
    p.LayerHeight    = GetStored("CFRC_H",       DEF_HEIGHT)
    p.FilamentDia    = GetStored("CFRC_D",       DEF_DIAMETER)
    p.FPrint         = GetStored("CFRC_FP",      DEF_F_PRINT)
    p.FTravel        = GetStored("CFRC_FT",      DEF_F_TRAVEL)
    p.RetractLen     = GetStored("CFRC_RT",      DEF_RETRACT)
    p.ZLiftThreshold = GetStored("CFRC_ZTHR",    DEF_Z_LIFT_THR)
    p.CutMacro       = GetStoredStr("CFRC_MACRO", DEF_CUT_MACRO)
    p.HasHeader      = (GetStored("CFRC_HDR_ROW", IIf(DEF_HAS_HEADER, 1, 0)) <> 0)
    p.Header         = GetStoredStr("CFRC_HEADER", DefaultHeader())

    ' 一次性表单不方便, 这里用一个紧凑的 InputBox 链.
    Dim s As String
    s = InputBox( _
        "请输入打印参数 (用逗号分隔, 共 7 项):" & vbCrLf & _
        "线宽W, 层高H, 丝径D, 打印速度FP, 空驶速度FT, 剪丝回抽RT, Z抬升阈值" & vbCrLf & vbCrLf & _
        "示例: 0.4, 0.2, 1.75, 1800, 3000, 0, 0.05", _
        "CFRC G-code 参数 (1/3)", _
        Format(p.LineWidth, "0.###") & ", " & Format(p.LayerHeight, "0.###") & ", " & _
        Format(p.FilamentDia, "0.###") & ", " & Format(p.FPrint, "0") & ", " & _
        Format(p.FTravel, "0") & ", " & Format(p.RetractLen, "0.###") & ", " & _
        Format(p.ZLiftThreshold, "0.###"))
    If Len(s) = 0 Then Exit Function
    Dim parts() As String: parts = Split(s, ",")
    If UBound(parts) < 6 Then MsgBox "参数项数不足", vbExclamation: Exit Function
    p.LineWidth      = CDbl(Trim(parts(0)))
    p.LayerHeight    = CDbl(Trim(parts(1)))
    p.FilamentDia    = CDbl(Trim(parts(2)))
    p.FPrint         = CDbl(Trim(parts(3)))
    p.FTravel        = CDbl(Trim(parts(4)))
    p.RetractLen     = CDbl(Trim(parts(5)))
    p.ZLiftThreshold = CDbl(Trim(parts(6)))

    Dim macroIn As String
    macroIn = InputBox("剪丝宏命令(将作为单独一行写入 G-code):", _
                       "CFRC G-code 参数 (2/3)", p.CutMacro)
    If Len(macroIn) = 0 Then Exit Function
    p.CutMacro = macroIn

    Dim hdrAns As VbMsgBoxResult
    hdrAns = MsgBox("是否使用上次保存的 Header? (否=重新粘贴)" & vbCrLf & vbCrLf & _
                    "当前 Header 前 3 行预览:" & vbCrLf & _
                    HeadPreview(p.Header, 3), _
                    vbYesNoCancel + vbQuestion, "CFRC G-code 参数 (3/3) - Header")
    If hdrAns = vbCancel Then Exit Function
    If hdrAns = vbNo Then
        ' 弹出大文本编辑窗 (用临时表)
        Dim newHdr As String
        newHdr = EditHeaderViaTempSheet(p.Header)
        If Len(newHdr) = 0 Then Exit Function
        p.Header = newHdr
    End If

    ' 是否表头
    Dim hasHdrAns As VbMsgBoxResult
    hasHdrAns = MsgBox("活动表第 1 行是表头吗? (是=从第 2 行读数据)", _
                       vbYesNoCancel + vbQuestion, "数据起始行")
    If hasHdrAns = vbCancel Then Exit Function
    p.HasHeader = (hasHdrAns = vbYes)

    ' 持久化
    SetStored "CFRC_W",        p.LineWidth
    SetStored "CFRC_H",        p.LayerHeight
    SetStored "CFRC_D",        p.FilamentDia
    SetStored "CFRC_FP",       p.FPrint
    SetStored "CFRC_FT",       p.FTravel
    SetStored "CFRC_RT",       p.RetractLen
    SetStored "CFRC_ZTHR",     p.ZLiftThreshold
    SetStored "CFRC_HDR_ROW",  IIf(p.HasHeader, 1, 0)
    SetStoredStr "CFRC_MACRO", p.CutMacro
    SetStoredStr "CFRC_HEADER", p.Header

    LoadParamsDialog = True
End Function

Private Function HeadPreview(s As String, n As Long) As String
    Dim arr() As String: arr = Split(Replace(s, vbCr, ""), vbLf)
    Dim i As Long, lines As String
    For i = 0 To Application.Min(n - 1, UBound(arr))
        lines = lines & arr(i) & vbCrLf
    Next i
    HeadPreview = lines
End Function

' 用一个临时工作簿表来编辑长 Header (比 InputBox 友好得多)
Private Function EditHeaderViaTempSheet(curHdr As String) As String
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Dim ws As Worksheet: Set ws = wb.Sheets(1)
    ws.Name = "EditHeader"
    ws.Range("A1").Value = "请把 G-code Header 完整粘贴到 A2(整段一格), 然后按本提示框'确定'. 取消则放弃."
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").ColumnWidth = 100
    ws.Range("A2").WrapText = True
    ws.Range("A2").Value = curHdr
    ws.Range("A2").Select
    Dim ans As VbMsgBoxResult
    ans = MsgBox("已打开临时表, 请在 A2 单元格内编辑/粘贴你的 Header." & vbCrLf & _
                 "完成后回到本对话框, 点 ' 确定 ' 应用; 点 ' 取消 ' 放弃.", _
                 vbOKCancel + vbInformation, "编辑 Header")
    If ans = vbOK Then
        EditHeaderViaTempSheet = CStr(ws.Range("A2").Value)
    Else
        EditHeaderViaTempSheet = ""
    End If
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function

'------------------------------------------------------------------------------
' Names 持久化 (写在加载项工作簿里, 不污染用户数据)
'------------------------------------------------------------------------------
Private Function StoreBook() As Workbook
    On Error Resume Next
    Set StoreBook = ThisWorkbook
    If StoreBook Is Nothing Then Set StoreBook = ActiveWorkbook
    On Error GoTo 0
End Function

Private Function GetStored(nm As String, def As Double) As Double
    On Error GoTo Fb
    Dim v As String
    v = StoreBook.Names(nm).Value     ' 形如 "=0.4"
    v = Replace(v, "=", "")
    v = Replace(v, """", "")
    GetStored = CDbl(v)
    Exit Function
Fb: GetStored = def
End Function

Private Sub SetStored(nm As String, v As Double)
    On Error Resume Next
    StoreBook.Names.Add Name:=nm, RefersTo:="=" & CStr(v), Visible:=False
End Sub

Private Function GetStoredStr(nm As String, def As String) As String
    On Error GoTo Fb
    Dim v As String
    v = StoreBook.Names(nm).Value
    ' Names 存字符串需要带引号; 这里反推
    If Left(v, 2) = "=""" And Right(v, 1) = """" Then
        v = Mid(v, 3, Len(v) - 3)
        v = Replace(v, """""", """")
    Else
        v = Replace(v, "=", "")
    End If
    GetStoredStr = v
    Exit Function
Fb: GetStoredStr = def
End Function

Private Sub SetStoredStr(nm As String, v As String)
    On Error Resume Next
    Dim esc As String: esc = Replace(v, """", """""")
    StoreBook.Names.Add Name:=nm, RefersTo:="=""" & esc & """", Visible:=False
End Sub

'------------------------------------------------------------------------------
' 默认 Header (你给定的那段)
'------------------------------------------------------------------------------
Private Function DefaultHeader() As String
    Dim s As String
    s = "; HEADER_BLOCK_START" & vbLf
    s = s & "; generated by OrcaSlicer 2.2.0 on 2026-04-15 at 15:15:01" & vbLf
    s = s & "; total layer number: 50" & vbLf
    s = s & "; filament_density: 1.04" & vbLf
    s = s & "; filament_diameter: 1.75" & vbLf
    s = s & "; max_z_height: 10.00" & vbLf
    s = s & "; HEADER_BLOCK_END" & vbLf
    s = s & "; external perimeters extrusion width = 0.40mm" & vbLf
    s = s & "; perimeters extrusion width = 0.44mm" & vbLf
    s = s & "; infill extrusion width = 0.44mm" & vbLf
    s = s & "; solid infill extrusion width = 0.48mm" & vbLf
    s = s & "; top infill extrusion width = 0.38mm" & vbLf
    s = s & "; first layer extrusion width = 0.48mm" & vbLf
    s = s & "M73 P0 R37                  ; 更新打印时间" & vbLf
    s = s & "M190 S70                    ; 等待热床温度" & vbLf
    s = s & "M109 S240 T0                ; 等待T0喷头温度" & vbLf
    s = s & "PRINT_START EXTRUDER=240 BED=70" & vbLf
    s = s & "G90                         ; 绝对坐标模式" & vbLf
    s = s & "G21                         ; 毫米单位" & vbLf
    s = s & "M83                         ; 挤出机相对模式" & vbLf
    s = s & "T0                          ; 选择T0喷头" & vbLf
    s = s & "M106 S100                   ; 开启风扇" & vbLf
    s = s & "G92 E0                      ; 重置挤出机"
    DefaultHeader = s
End Function

'==============================================================================
' 功能区 / 加载项启动钩子
'==============================================================================
Public Sub OnRibbonExportClick(control As IRibbonControl)
    ExportGcodeFromSheet
End Sub

Public Sub OnRibbonAboutClick(control As IRibbonControl)
    MsgBox "CFRC G-code Exporter v1.0" & vbCrLf & _
           "为 Albert (冯镜泽) 定制" & vbCrLf & vbCrLf & _
           "数据格式: A=X  B=Y  C=Z  D=Cut(1=剪丝)" & vbCrLf & _
           "挤出公式: E = L × W × H / (π·(D/2)²) (相对模式 M83)" & vbCrLf & _
           "剪丝逻辑: 到达点后插入宏命令, 下一段自动空驶", _
           vbInformation, "About"
End Sub
