Attribute VB_Name = "CFRC_GcodeExporter"
'==============================================================================
' CFRC G-code Exporter for Excel  (v1.1 - no external references needed)
'   Author: Jazz Feng
'   Data:   A=X(mm)  B=Y(mm)  C=Z(mm)  D=Cut flag(1=cut, blank=no)
'   Output: G1 moves + custom Header + cut macros
'   Extrusion: relative mode (M83), E = seg_len * W * H / (pi*(D/2)^2)
'==============================================================================
Option Explicit

' ---------- Parameter struct (must be declared first) ----------
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
    TotalE         As Double
    TotalLen       As Double
    CutCount       As Long
End Type

' ---------- Defaults ----------
Private Const DEF_WIDTH       As Double = 0.4
Private Const DEF_HEIGHT      As Double = 0.2
Private Const DEF_DIAMETER    As Double = 1.75
Private Const DEF_F_PRINT     As Double = 1800#
Private Const DEF_F_TRAVEL    As Double = 3000#
Private Const DEF_RETRACT     As Double = 0#
Private Const DEF_CUT_MACRO   As String = "CUT_FIBER"
Private Const DEF_HAS_HEADER  As Boolean = True
Private Const DEF_Z_LIFT_THR  As Double = 0.05

'------------------------------------------------------------------------------
' Entry point:  Alt+F8 -> ExportGcodeFromSheet
'------------------------------------------------------------------------------
Public Sub ExportGcodeFromSheet()
    On Error GoTo ErrH

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "Please open a worksheet first.", vbExclamation: Exit Sub
    End If

    Dim p As GcodeParams
    If Not LoadParamsDialog(p) Then Exit Sub

    Dim startRow As Long
    startRow = IIf(p.HasHeader, 2, 1)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < startRow Then
        MsgBox "A column has no data (from row " & startRow & ").", vbExclamation: Exit Sub
    End If

    Dim nPts As Long: nPts = lastRow - startRow + 1
    If nPts < 2 Then
        MsgBox "Need at least 2 points.", vbExclamation: Exit Sub
    End If

    Dim arr As Variant
    arr = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, 4)).Value

    ' Save dialog (pure VBA, no Office library needed)
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=ws.Name & ".gcode", _
        FileFilter:="G-code (*.gcode), *.gcode, All (*.*), *.*", _
        Title:="Save G-code")
    If savePath = False Then Exit Sub

    Dim gcode As String
    gcode = BuildGcode(arr, nPts, p)

    WriteUtf8 CStr(savePath), gcode

    MsgBox "Exported: " & vbCrLf & CStr(savePath) & vbCrLf & vbCrLf & _
           "Points: " & nPts & "    Cuts: " & p.CutCount & vbCrLf & _
           "Total E = " & Format(p.TotalE, "0.000") & " mm" & vbCrLf & _
           "Total path = " & Format(p.TotalLen, "0.000") & " mm", _
           vbInformation, "Export done"
    Exit Sub

ErrH:
    MsgBox "Error: " & Err.Description, vbCritical, "CFRC G-code Exporter"
End Sub

'------------------------------------------------------------------------------
' Core: build G-code string
'------------------------------------------------------------------------------
Private Function BuildGcode(arr As Variant, nPts As Long, p As GcodeParams) As String
    Dim lines() As String
    Dim lc As Long: lc = 0
    Dim cap As Long: cap = nPts * 3 + 100
    ReDim lines(1 To cap)

    ' Header
    Dim hdrParts() As String
    hdrParts = Split(p.Header, vbLf)
    Dim h As Long
    For h = LBound(hdrParts) To UBound(hdrParts)
        lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
        lines(lc) = Replace(hdrParts(h), vbCr, "")
    Next h

    ' First point
    Dim x1 As Double, y1 As Double, z1 As Double
    x1 = CDbl(arr(1, 1)): y1 = CDbl(arr(1, 2)): z1 = CDbl(arr(1, 3))

    lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
    lines(lc) = ""
    lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
    lines(lc) = "; ===== Toolpath start (from XLSX) ====="
    lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
    lines(lc) = "G92 E0                      ; reset E"
    lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
    lines(lc) = "G1 X" & F(x1) & " Y" & F(y1) & " Z" & F(z1) & " F" & F(p.FTravel) & " ; move to start"

    Dim ePerMm As Double
    ePerMm = (p.LineWidth * p.LayerHeight) / (3.14159265358979 * (p.FilamentDia / 2#) ^ 2)

    Dim travelNext As Boolean: travelNext = False
    Dim totalE As Double: totalE = 0#
    Dim totalLen As Double: totalLen = 0#
    Dim cutCount As Long: cutCount = 0
    Dim curLayerZ As Double: curLayerZ = z1
    Dim layerIdx As Long: layerIdx = 1

    If IsCutCell(arr(1, 4)) Then
        lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
        lines(lc) = "; ---- Cut at start ----"
        EmitCut lines, lc, cap, p
        cutCount = cutCount + 1
        travelNext = True
    End If

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

        Dim isLift As Boolean
        isLift = (Abs(dz) > p.ZLiftThreshold)

        If isLift And zc <> curLayerZ Then
            layerIdx = layerIdx + 1
            lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
            lines(lc) = ""
            lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
            lines(lc) = "; ---- Layer " & layerIdx & "  Z=" & F(zc) & " ----"
            curLayerZ = zc
        End If

        Dim asTravel As Boolean
        asTravel = travelNext Or isLift Or (segLen <= 0.000001)

        lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)

        If asTravel Then
            Dim cmt As String
            If travelNext Then
                cmt = " ; travel after cut"
            ElseIf isLift Then
                cmt = " ; layer move"
            Else
                cmt = ""
            End If
            lines(lc) = "G1 X" & F(xc) & " Y" & F(yc) & " Z" & F(zc) & " F" & F(p.FTravel) & cmt
        Else
            Dim de As Double: de = segLen * ePerMm
            totalE = totalE + de
            totalLen = totalLen + segLen
            If Abs(dz) < 0.000001 Then
                lines(lc) = "G1 X" & F(xc) & " Y" & F(yc) & " E" & F(de) & " F" & F(p.FPrint)
            Else
                lines(lc) = "G1 X" & F(xc) & " Y" & F(yc) & " Z" & F(zc) & " E" & F(de) & " F" & F(p.FPrint)
            End If
        End If

        travelNext = False

        If cutFlag Then
            lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
            lines(lc) = "; ---- Cut at point " & r & " ----"
            EmitCut lines, lc, cap, p
            cutCount = cutCount + 1
            travelNext = True
        End If

        xPrev = xc: yPrev = yc: zPrev = zc
    Next r

    lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
    lines(lc) = ""
    lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
    lines(lc) = "; ===== End of toolpath ====="

    p.TotalE = totalE
    p.TotalLen = totalLen
    p.CutCount = cutCount

    ReDim Preserve lines(1 To lc)
    BuildGcode = Join(lines, vbCrLf) & vbCrLf
End Function

Private Sub EmitCut(ByRef lines() As String, ByRef lc As Long, ByRef cap As Long, p As GcodeParams)
    If p.RetractLen > 0 Then
        lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
        lines(lc) = "G1 E-" & F(p.RetractLen) & " F" & F(p.FTravel) & " ; retract"
    End If
    lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
    lines(lc) = p.CutMacro
    lc = lc + 1: If lc > cap Then cap = cap * 2: ReDim Preserve lines(1 To cap)
    lines(lc) = "G92 E0                      ; reset E after cut"
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Function F(v As Double) As String
    Dim s As String: s = Format(v, "0.#####")
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

Private Sub WriteUtf8(path As String, content As String)
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open: stm.WriteText content
    Dim bin As Object: Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1: bin.Mode = 3: bin.Open
    stm.Position = 3: stm.CopyTo bin
    stm.Flush: stm.Close
    bin.SaveToFile path, 2: bin.Close
End Sub

'==============================================================================
' Parameter dialog
'==============================================================================
Private Function LoadParamsDialog(ByRef p As GcodeParams) As Boolean
    LoadParamsDialog = False

    p.LineWidth = GetStored("CFRC_W", DEF_WIDTH)
    p.LayerHeight = GetStored("CFRC_H", DEF_HEIGHT)
    p.FilamentDia = GetStored("CFRC_D", DEF_DIAMETER)
    p.FPrint = GetStored("CFRC_FP", DEF_F_PRINT)
    p.FTravel = GetStored("CFRC_FT", DEF_F_TRAVEL)
    p.RetractLen = GetStored("CFRC_RT", DEF_RETRACT)
    p.ZLiftThreshold = GetStored("CFRC_ZTHR", DEF_Z_LIFT_THR)
    p.CutMacro = GetStoredStr("CFRC_MACRO", DEF_CUT_MACRO)
    p.HasHeader = (GetStored("CFRC_HDR_ROW", IIf(DEF_HAS_HEADER, 1, 0)) <> 0)
    p.Header = GetStoredStr("CFRC_HEADER", DefaultHeader())

    Dim s As String
    s = InputBox( _
        "Print params (comma-separated, 7 items):" & vbCrLf & _
        "W, H, D, F_print, F_travel, Retract, Z_threshold" & vbCrLf & vbCrLf & _
        "Example: 0.4, 0.2, 1.75, 1800, 3000, 0, 0.05", _
        "CFRC G-code Params (1/3)", _
        Format(p.LineWidth, "0.###") & ", " & Format(p.LayerHeight, "0.###") & ", " & _
        Format(p.FilamentDia, "0.###") & ", " & Format(p.FPrint, "0") & ", " & _
        Format(p.FTravel, "0") & ", " & Format(p.RetractLen, "0.###") & ", " & _
        Format(p.ZLiftThreshold, "0.###"))
    If Len(s) = 0 Then Exit Function
    Dim parts() As String: parts = Split(s, ",")
    If UBound(parts) < 6 Then MsgBox "Need 7 items.", vbExclamation: Exit Function
    p.LineWidth = CDbl(Trim(parts(0)))
    p.LayerHeight = CDbl(Trim(parts(1)))
    p.FilamentDia = CDbl(Trim(parts(2)))
    p.FPrint = CDbl(Trim(parts(3)))
    p.FTravel = CDbl(Trim(parts(4)))
    p.RetractLen = CDbl(Trim(parts(5)))
    p.ZLiftThreshold = CDbl(Trim(parts(6)))

    Dim macroIn As String
    macroIn = InputBox("Cut macro command:", "CFRC G-code Params (2/3)", p.CutMacro)
    If Len(macroIn) = 0 Then Exit Function
    p.CutMacro = macroIn

    Dim hdrAns As VbMsgBoxResult
    hdrAns = MsgBox("Use saved Header?" & vbCrLf & vbCrLf & _
                    "Preview (first 3 lines):" & vbCrLf & _
                    HeadPreview(p.Header, 3) & vbCrLf & _
                    "[Yes] = use saved    [No] = paste new    [Cancel] = abort", _
                    vbYesNoCancel + vbQuestion, "CFRC G-code Params (3/3)")
    If hdrAns = vbCancel Then Exit Function
    If hdrAns = vbNo Then
        Dim newHdr As String
        newHdr = InputBox("Paste full G-code Header:", "Edit Header", p.Header)
        If Len(newHdr) = 0 Then Exit Function
        p.Header = newHdr
    End If

    Dim rowAns As VbMsgBoxResult
    rowAns = MsgBox("Does Row 1 contain column headers?" & vbCrLf & _
                    "[Yes] = data starts at Row 2" & vbCrLf & _
                    "[No]  = data starts at Row 1", _
                    vbYesNoCancel + vbQuestion, "Data start row")
    If rowAns = vbCancel Then Exit Function
    p.HasHeader = (rowAns = vbYes)

    SetStored "CFRC_W", p.LineWidth
    SetStored "CFRC_H", p.LayerHeight
    SetStored "CFRC_D", p.FilamentDia
    SetStored "CFRC_FP", p.FPrint
    SetStored "CFRC_FT", p.FTravel
    SetStored "CFRC_RT", p.RetractLen
    SetStored "CFRC_ZTHR", p.ZLiftThreshold
    SetStored "CFRC_HDR_ROW", IIf(p.HasHeader, 1, 0)
    SetStoredStr "CFRC_MACRO", p.CutMacro
    SetStoredStr "CFRC_HEADER", p.Header

    LoadParamsDialog = True
End Function

Private Function HeadPreview(s As String, n As Long) As String
    Dim arr() As String: arr = Split(Replace(s, vbCr, ""), vbLf)
    Dim i As Long, result As String
    For i = 0 To Application.Min(n - 1, UBound(arr))
        result = result & arr(i) & vbCrLf
    Next i
    HeadPreview = result
End Function

'------------------------------------------------------------------------------
' Persistent storage via Names
'------------------------------------------------------------------------------
Private Function StoreBook() As Workbook
    On Error Resume Next
    Set StoreBook = ThisWorkbook
    If StoreBook Is Nothing Then Set StoreBook = ActiveWorkbook
    On Error GoTo 0
End Function

Private Function GetStored(nm As String, def As Double) As Double
    On Error GoTo Fb
    Dim v As String: v = StoreBook.Names(nm).Value
    v = Replace(Replace(v, "=", ""), """", "")
    GetStored = CDbl(v): Exit Function
Fb: GetStored = def
End Function

Private Sub SetStored(nm As String, v As Double)
    On Error Resume Next
    StoreBook.Names.Add Name:=nm, RefersTo:="=" & CStr(v), Visible:=False
End Sub

Private Function GetStoredStr(nm As String, def As String) As String
    On Error GoTo Fb
    Dim v As String: v = StoreBook.Names(nm).Value
    If Left(v, 2) = "=""" And Right(v, 1) = """" Then
        v = Mid(v, 3, Len(v) - 3): v = Replace(v, """""", """")
    Else: v = Replace(v, "=", "")
    End If
    GetStoredStr = v: Exit Function
Fb: GetStoredStr = def
End Function

Private Sub SetStoredStr(nm As String, v As String)
    On Error Resume Next
    StoreBook.Names.Add Name:=nm, RefersTo:="=""" & Replace(v, """", """""") & """", Visible:=False
End Sub

'------------------------------------------------------------------------------
' Default Header
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
    s = s & "M73 P0 R37" & vbLf
    s = s & "M190 S70" & vbLf
    s = s & "M109 S240 T0" & vbLf
    s = s & "PRINT_START EXTRUDER=240 BED=70" & vbLf
    s = s & "G90" & vbLf
    s = s & "G21" & vbLf
    s = s & "M83" & vbLf
    s = s & "T0" & vbLf
    s = s & "M106 S100" & vbLf
    s = s & "G92 E0"
    DefaultHeader = s
End Function
