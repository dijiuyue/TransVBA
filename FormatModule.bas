Option Private Module
Option Explicit
' 需要引用 TypeModule 中的 gSettings / FormatSettings
' 也需要在“工具-引用”里保证引用了 Word 对象库（在 WPS 中同理）

' 启动前自检：验证当前工程的 FormatSettings 是否支持 1~5 级标题数组
Public Function ValidateTitleLevel5Support() As Boolean
    On Error Resume Next
    Dim tmp As Variant
    
    tmp = gSettings.TitleAlignment(5)
    If Err.Number <> 0 Then
        ValidateTitleLevel5Support = False
        Err.Clear
        Exit Function
    End If
    
    tmp = gSettings.TitleFont(5)
    If Err.Number <> 0 Then
        ValidateTitleLevel5Support = False
        Err.Clear
        Exit Function
    End If
    
    tmp = gSettings.TitleBeforeLines(5)
    If Err.Number <> 0 Then
        ValidateTitleLevel5Support = False
        Err.Clear
        Exit Function
    End If
    
    tmp = gSettings.TitleAfterLines(5)
    If Err.Number <> 0 Then
        ValidateTitleLevel5Support = False
        Err.Clear
        Exit Function
    End If
    
    tmp = gSettings.TitleLineSpacing(5)
    If Err.Number <> 0 Then
        ValidateTitleLevel5Support = False
        Err.Clear
        Exit Function
    End If
    
    ValidateTitleLevel5Support = True
    On Error GoTo 0
End Function

' 手动测试入口：在立即窗口/宏列表执行，用于快速确认是否仍是旧版 typemodule
Public Sub TestTitleArrayBound()
    If ValidateTitleLevel5Support() Then
        MsgBox "检测通过：当前工程支持 1~5 级标题数组。", vbInformation
    Else
        MsgBox "检测失败：当前工程中的 FormatSettings 仍为旧版本（仅 1~4 级）。请在 VBE 的 typemodule 中把标题数组改为 (1 To 5)。", vbExclamation
    End If
End Sub

'========================
' 1. 从窗体读取设置并应用
'========================

Public Sub ApplyFormatting()
    Dim prevScreenUpdating As Boolean
    Dim hasPrevScreenUpdating As Boolean
    
    ' 性能优化：运行期间关闭界面刷新（结束时恢复）
    On Error Resume Next
    prevScreenUpdating = Application.ScreenUpdating
    hasPrevScreenUpdating = (Err.Number = 0)
    Err.Clear
    Application.ScreenUpdating = False
    On Error GoTo EH
    
    ' 若当前工程实际引用的 FormatSettings 仍是 1~4（历史版本），
    ' 对第 5 级的访问会触发“下标越界”。这里先做容错，避免直接崩溃。
    On Error Resume Next
    Dim testVal As Variant
    testVal = gSettings.TitleAlignment(5)
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "当前工程的 typemodule/FormatSettings 仍为旧版本（仅支持1~4级标题）。请更新 typemodule.bas 后再使用5级标题功能。", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' 从当前 UserForm 控件读取设置到 gSettings
    GetSettingsFromForm

    ' 目录设置已移除：始终使用固定默认值
    SetDefaultDirectorySettings

    ' 如需记忆，写入注册表（本地）
    If gSettings.RememberSettings Then
        SaveSettingsToRegistry
    End If

    ' 按任务书要求刷新当前文档格式
    ApplySettingsToDocument

    On Error Resume Next
    If hasPrevScreenUpdating Then
        Application.ScreenUpdating = prevScreenUpdating
    Else
        Application.ScreenUpdating = True
    End If
    On Error GoTo 0

    MsgBox "格式刷新完成。", vbInformation
    Exit Sub
EH:
    ' 出错也要恢复界面刷新
    On Error Resume Next
    If hasPrevScreenUpdating Then
        Application.ScreenUpdating = prevScreenUpdating
    Else
        Application.ScreenUpdating = True
    End If
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub GetSettingsFromForm()
    On Error GoTo EH
    Dim uf As UserForm1
        Dim mp As Object
    Dim frame As Object
    Dim i As Integer

    Set uf = UserForm1
    Set mp = uf.Controls("mpSettings")

    ' 记忆勾选框在窗体顶部
    On Error Resume Next
    gSettings.RememberSettings = (uf.Controls("chkRemember").Value = True)
    On Error GoTo 0

    ' ---------- 正文页（第 0 页） ----------
    Set frame = mp.Pages(0).Controls("frameContent")
    With frame
        gSettings.BodyFont = .Controls("cmbBodyFont").Value
        gSettings.BodySize = .Controls("txtBodySize").Text
        gSettings.BodySpacing = Val(.Controls("txtBodySpacing").Text)
        gSettings.BodyBeforeLines = Val(.Controls("txtBodyBefore").Text)
        gSettings.BodyAfterLines = Val(.Controls("txtBodyAfter").Text)
        gSettings.BodyAlignment = .Controls("cmbBodyAlign").Value
        gSettings.BodyLeftIndentCm = Val(.Controls("txtBodyLeftIndent").Text)
        gSettings.BodyRightIndentCm = Val(.Controls("txtBodyRightIndent").Text)
        gSettings.BodySpecialIndent = .Controls("cmbBodySpecial").Value
        gSettings.BodySpecialIndentCm = Val(.Controls("txtBodySpecialIndent").Text)
        End With
        
    ' ---------- 标题页（第 1 页） ----------
    Set frame = mp.Pages(1).Controls("frameTitleConfig")

    ' 读取 1~4 级标题设置
    Dim iTitle As Integer
    Dim grp As Object
    For iTitle = 1 To 5
        Set grp = frame.Controls("frameTitle" & iTitle)
        gSettings.TitleAlignment(iTitle) = grp.Controls("cmbAlign" & iTitle).Value
        gSettings.TitleFont(iTitle) = grp.Controls("cmbTitleFont" & iTitle).Value
        gSettings.TitleSize(iTitle) = grp.Controls("txtTitleSize" & iTitle).Text
        gSettings.TitleBold(iTitle) = (grp.Controls("chkBold" & iTitle).Value = True)
        gSettings.TitleBeforeLines(iTitle) = Val(grp.Controls("txtBeforeLines" & iTitle).Text)
        gSettings.TitleAfterLines(iTitle) = Val(grp.Controls("txtAfterLines" & iTitle).Text)
        gSettings.TitleLineSpacing(iTitle) = Val(grp.Controls("txtTitleSpacing" & iTitle).Text)
    Next iTitle

    ' 自动识别数字标题（大纲级别）
    On Error Resume Next
    gSettings.AutoDetectNumericTitles = (frame.Controls("chkAutoDetectNumericTitles").Value = True)
    gSettings.AutoDetectIncludeListParagraphs = (frame.Controls("chkAutoDetectIncludeList").Value = True)
    On Error GoTo 0
        
    ' ---------- 表格页（第 2 页） ----------
    Set frame = mp.Pages(2).Controls("frameTable")
    With frame
        gSettings.TableTitleFont = .Controls("cmbTableTitleFont").Value
        gSettings.TableTitleSize = .Controls("txtTableTitleSize").Text
        gSettings.TableTitleBold = (.Controls("chkTableTitleBold").Value = True)
        gSettings.TableTitleSpacing = Val(.Controls("txtTableTitleSpacing").Text)
        gSettings.TableTitleLeftIndentCm = Val(.Controls("txtTableTitleLeftIndent").Text)
        gSettings.TableTitleRightIndentCm = Val(.Controls("txtTableTitleRightIndent").Text)
        gSettings.TableTitleSpecialIndent = .Controls("cmbTableTitleSpecial").Value
        gSettings.TableTitleSpecialIndentCm = Val(.Controls("txtTableTitleSpecialIndent").Text)
        If gSettings.TableTitleSpecialIndent = "" Then gSettings.TableTitleSpecialIndent = "无"

        gSettings.TableFont = .Controls("cmbTableFont").Value
        gSettings.TableSize = .Controls("txtTableSize").Text
        gSettings.TableLineWidth = Val(.Controls("txtLineWidth").Text)
        gSettings.TableRowHeight = Val(.Controls("txtRowHeight").Text)
        gSettings.TableSpacing = Val(.Controls("txtTableSpacing").Text)
        gSettings.TableAutoFitWindow = (.Controls("chkTableAutoFitWindow").Value = True)
    End With

    ' ---------- 图片标题页（第 3 页） ----------
    Set frame = mp.Pages(3).Controls("frameFigure")
    With frame
        gSettings.FigureTitleFont = .Controls("cmbFigureTitleFont").Value
        gSettings.FigureTitleSize = .Controls("txtFigureTitleSize").Text
        gSettings.FigureTitleBold = (.Controls("chkFigureTitleBold").Value = True)
        gSettings.FigureTitleSpacing = Val(.Controls("txtFigureTitleSpacing").Text)
        gSettings.FigureTitleLeftIndentCm = Val(.Controls("txtFigureTitleLeftIndent").Text)
        gSettings.FigureTitleRightIndentCm = Val(.Controls("txtFigureTitleRightIndent").Text)
        gSettings.FigureTitleSpecialIndent = .Controls("cmbFigureTitleSpecial").Value
        gSettings.FigureTitleSpecialIndentCm = Val(.Controls("txtFigureTitleSpecialIndent").Text)
        If gSettings.FigureTitleSpecialIndent = "" Then gSettings.FigureTitleSpecialIndent = "无"
    End With

    ' ---------- 为 1~4 级标题设置缺省值 ----------
    For i = 1 To 5
        ' 默认对齐：1、2 级居中，其它左对齐
            If gSettings.TitleAlignment(i) = "" Then
            If i <= 2 Then
                gSettings.TitleAlignment(i) = "居中"
            Else
                gSettings.TitleAlignment(i) = "左对齐"
                End If
            End If
            
        If gSettings.TitleFont(i) = "" Then gSettings.TitleFont(i) = gSettings.BodyFont
        If gSettings.TitleSize(i) = "" Then gSettings.TitleSize(i) = gSettings.BodySize
        If gSettings.TitleLineSpacing(i) = 0 Then gSettings.TitleLineSpacing(i) = gSettings.BodySpacing

        ' 默认段前/段后：可按需要自行调整，这里默认 0
        ' 若用户没有填，保持 0

        ' 默认加粗：1~3 级加粗，4 级不加粗（若未手动勾选）
            If gSettings.TitleBold(i) = False Then
            If i <= 3 Then gSettings.TitleBold(i) = True
            End If
        Next i
        
    Exit Sub
EH:
    MsgBox "读取窗体设置时发生错误（下标越界/控件不存在）。" & vbCrLf & _
           "Err=" & CStr(Err.Number) & " " & Err.Description & vbCrLf & _
           "常见原因：标题页缺少第5级控件 frameTitle5 或内部控件名不一致。" & vbCrLf & _
           "请把调试窗口高亮的那一行截图发我，我就能精确修复。", vbExclamation
    Err.Clear
End Sub

Private Sub SetDefaultDirectorySettings()
    ' 固定默认目录设置（界面已移除，不再可配置）
    gSettings.Level1Font = "宋体"
    gSettings.Level1Size = "小四"
    gSettings.Level1Spacing = 1
    gSettings.Level1Bold = True

    gSettings.Level2Font = "宋体"
    gSettings.Level2Size = "小四"
    gSettings.Level2IndentChars = 2
    gSettings.Level2Spacing = 1
    gSettings.Level2Bold = False

    gSettings.Level3Font = "宋体"
    gSettings.Level3Size = "小四"
    gSettings.Level3IndentChars = 2
    gSettings.Level3Spacing = 1
    gSettings.Level3Bold = False
End Sub

'========================
' 2. 记忆设置（本地注册表）
'========================


Public Sub LoadSettingsFromRegistry()
    On Error Resume Next

    ' 只加载关键字段，没存的用缺省
    gSettings.BodyFont = GetSetting(APP_NAME, SECTION_NAME, "BodyFont", "宋体")
    gSettings.BodySize = GetSetting(APP_NAME, SECTION_NAME, "BodySize", "小四")
    gSettings.BodySpacing = Val(GetSetting(APP_NAME, SECTION_NAME, "BodySpacing", "1.5"))
    gSettings.BodyBeforeLines = Val(GetSetting(APP_NAME, SECTION_NAME, "BodyBeforeLines", "0"))
    gSettings.BodyAfterLines = Val(GetSetting(APP_NAME, SECTION_NAME, "BodyAfterLines", "0"))
    gSettings.BodyAlignment = GetSetting(APP_NAME, SECTION_NAME, "BodyAlign", "左对齐")
    gSettings.BodyLeftIndentCm = Val(GetSetting(APP_NAME, SECTION_NAME, "BodyLeftIndentCm", "0"))
    gSettings.BodyRightIndentCm = Val(GetSetting(APP_NAME, SECTION_NAME, "BodyRightIndentCm", "0"))
    gSettings.BodySpecialIndent = GetSetting(APP_NAME, SECTION_NAME, "BodySpecialIndent", "无")
    gSettings.BodySpecialIndentCm = Val(GetSetting(APP_NAME, SECTION_NAME, "BodySpecialIndentCm", "0"))
    
    ' 目录设置已移除：始终使用固定默认值
    SetDefaultDirectorySettings
    
    ' 各级正文标题的对齐/段前段后/字体/字号/加粗（仅 1~5）
    Dim i As Integer
    For i = 1 To 5
        gSettings.TitleAlignment(i) = GetSetting(APP_NAME, SECTION_NAME, "TitleAlign" & CStr(i), _
                                IIf(i <= 2, "居中", "左对齐"))
        gSettings.TitleBold(i) = (GetSetting(APP_NAME, SECTION_NAME, "TitleBold" & CStr(i), _
                                IIf(i <= 3, "1", "0")) = "1")
        gSettings.TitleFont(i) = GetSetting(APP_NAME, SECTION_NAME, "TitleFont" & CStr(i), gSettings.BodyFont)
        gSettings.TitleSize(i) = GetSetting(APP_NAME, SECTION_NAME, "TitleSize" & CStr(i), gSettings.BodySize)
        gSettings.TitleBeforeLines(i) = Val(GetSetting(APP_NAME, SECTION_NAME, "TitleBeforeLines" & CStr(i), "0"))
        gSettings.TitleAfterLines(i) = Val(GetSetting(APP_NAME, SECTION_NAME, "TitleAfterLines" & CStr(i), "0"))
        gSettings.TitleLineSpacing(i) = Val(GetSetting(APP_NAME, SECTION_NAME, "TitleLineSpacing" & CStr(i), CStr(gSettings.BodySpacing)))
    Next i

    gSettings.AutoDetectNumericTitles = (GetSetting(APP_NAME, SECTION_NAME, "AutoDetectNumericTitles", "1") = "1")
    gSettings.AutoDetectIncludeListParagraphs = (GetSetting(APP_NAME, SECTION_NAME, "AutoDetectIncludeListParagraphs", "1") = "1")

    gSettings.TableTitleFont = GetSetting(APP_NAME, SECTION_NAME, "TableTitleFont", "宋体")
    gSettings.TableTitleSize = GetSetting(APP_NAME, SECTION_NAME, "TableTitleSize", "五号")
    gSettings.TableTitleBold = (GetSetting(APP_NAME, SECTION_NAME, "TableTitleBold", "1") = "1")
    gSettings.TableTitleSpacing = Val(GetSetting(APP_NAME, SECTION_NAME, "TableTitleSpacing", "1.0"))
    gSettings.TableTitleLeftIndentCm = Val(GetSetting(APP_NAME, SECTION_NAME, "TableTitleLeftIndentCm", "0"))
    gSettings.TableTitleRightIndentCm = Val(GetSetting(APP_NAME, SECTION_NAME, "TableTitleRightIndentCm", "0"))
    gSettings.TableTitleSpecialIndent = GetSetting(APP_NAME, SECTION_NAME, "TableTitleSpecialIndent", "无")
    gSettings.TableTitleSpecialIndentCm = Val(GetSetting(APP_NAME, SECTION_NAME, "TableTitleSpecialIndentCm", "0"))

    gSettings.TableFont = GetSetting(APP_NAME, SECTION_NAME, "TableFont", "宋体")
    gSettings.TableSize = GetSetting(APP_NAME, SECTION_NAME, "TableSize", "五号")
    gSettings.TableLineWidth = Val(GetSetting(APP_NAME, SECTION_NAME, "TableLineWidth", "0.25"))
    gSettings.TableRowHeight = Val(GetSetting(APP_NAME, SECTION_NAME, "TableRowHeight", "0.6"))
    gSettings.TableSpacing = Val(GetSetting(APP_NAME, SECTION_NAME, "TableSpacing", "1.0"))
    gSettings.TableAutoFitWindow = (GetSetting(APP_NAME, SECTION_NAME, "TableAutoFitWindow", "0") = "1")

    ' 图片标题：默认与表格标题一致
    gSettings.FigureTitleFont = GetSetting(APP_NAME, SECTION_NAME, "FigureTitleFont", gSettings.TableTitleFont)
    gSettings.FigureTitleSize = GetSetting(APP_NAME, SECTION_NAME, "FigureTitleSize", gSettings.TableTitleSize)
    gSettings.FigureTitleBold = (GetSetting(APP_NAME, SECTION_NAME, "FigureTitleBold", IIf(gSettings.TableTitleBold, "1", "0")) = "1")
    gSettings.FigureTitleSpacing = Val(GetSetting(APP_NAME, SECTION_NAME, "FigureTitleSpacing", CStr(IIf(gSettings.TableTitleSpacing > 0, gSettings.TableTitleSpacing, 1#))))
    gSettings.FigureTitleLeftIndentCm = Val(GetSetting(APP_NAME, SECTION_NAME, "FigureTitleLeftIndentCm", "0"))
    gSettings.FigureTitleRightIndentCm = Val(GetSetting(APP_NAME, SECTION_NAME, "FigureTitleRightIndentCm", "0"))
    gSettings.FigureTitleSpecialIndent = GetSetting(APP_NAME, SECTION_NAME, "FigureTitleSpecialIndent", "无")
    gSettings.FigureTitleSpecialIndentCm = Val(GetSetting(APP_NAME, SECTION_NAME, "FigureTitleSpecialIndentCm", "0"))

    gSettings.RememberSettings = (GetSetting(APP_NAME, SECTION_NAME, "Remember", "1") = "1")
End Sub

Private Sub SaveSettingsToRegistry()
    On Error Resume Next

    SaveSetting APP_NAME, SECTION_NAME, "BodyFont", gSettings.BodyFont
    SaveSetting APP_NAME, SECTION_NAME, "BodySize", gSettings.BodySize
    SaveSetting APP_NAME, SECTION_NAME, "BodySpacing", CStr(gSettings.BodySpacing)
    SaveSetting APP_NAME, SECTION_NAME, "BodyBeforeLines", CStr(gSettings.BodyBeforeLines)
    SaveSetting APP_NAME, SECTION_NAME, "BodyAfterLines", CStr(gSettings.BodyAfterLines)
    SaveSetting APP_NAME, SECTION_NAME, "BodyAlign", gSettings.BodyAlignment
    SaveSetting APP_NAME, SECTION_NAME, "BodyLeftIndentCm", CStr(gSettings.BodyLeftIndentCm)
    SaveSetting APP_NAME, SECTION_NAME, "BodyRightIndentCm", CStr(gSettings.BodyRightIndentCm)
    SaveSetting APP_NAME, SECTION_NAME, "BodySpecialIndent", gSettings.BodySpecialIndent
    SaveSetting APP_NAME, SECTION_NAME, "BodySpecialIndentCm", CStr(gSettings.BodySpecialIndentCm)
    
    ' 目录设置已移除：不再保存目录相关键
    
    ' 各级正文标题的对齐/段前段后/字体/字号/加粗（仅 1~4）
    Dim i As Integer
    For i = 1 To 5
        SaveSetting APP_NAME, SECTION_NAME, "TitleAlign" & CStr(i), gSettings.TitleAlignment(i)
        SaveSetting APP_NAME, SECTION_NAME, "TitleBold" & CStr(i), IIf(gSettings.TitleBold(i), "1", "0")
        SaveSetting APP_NAME, SECTION_NAME, "TitleFont" & CStr(i), gSettings.TitleFont(i)
        SaveSetting APP_NAME, SECTION_NAME, "TitleSize" & CStr(i), gSettings.TitleSize(i)
        SaveSetting APP_NAME, SECTION_NAME, "TitleBeforeLines" & CStr(i), CStr(gSettings.TitleBeforeLines(i))
        SaveSetting APP_NAME, SECTION_NAME, "TitleAfterLines" & CStr(i), CStr(gSettings.TitleAfterLines(i))
        SaveSetting APP_NAME, SECTION_NAME, "TitleLineSpacing" & CStr(i), CStr(gSettings.TitleLineSpacing(i))
    Next i

    SaveSetting APP_NAME, SECTION_NAME, "AutoDetectNumericTitles", IIf(gSettings.AutoDetectNumericTitles, "1", "0")
    SaveSetting APP_NAME, SECTION_NAME, "AutoDetectIncludeListParagraphs", IIf(gSettings.AutoDetectIncludeListParagraphs, "1", "0")

    SaveSetting APP_NAME, SECTION_NAME, "TableTitleFont", gSettings.TableTitleFont
    SaveSetting APP_NAME, SECTION_NAME, "TableTitleSize", gSettings.TableTitleSize
    SaveSetting APP_NAME, SECTION_NAME, "TableTitleBold", IIf(gSettings.TableTitleBold, "1", "0")
    SaveSetting APP_NAME, SECTION_NAME, "TableTitleSpacing", CStr(gSettings.TableTitleSpacing)
    SaveSetting APP_NAME, SECTION_NAME, "TableTitleLeftIndentCm", CStr(gSettings.TableTitleLeftIndentCm)
    SaveSetting APP_NAME, SECTION_NAME, "TableTitleRightIndentCm", CStr(gSettings.TableTitleRightIndentCm)
    SaveSetting APP_NAME, SECTION_NAME, "TableTitleSpecialIndent", gSettings.TableTitleSpecialIndent
    SaveSetting APP_NAME, SECTION_NAME, "TableTitleSpecialIndentCm", CStr(gSettings.TableTitleSpecialIndentCm)

    SaveSetting APP_NAME, SECTION_NAME, "TableFont", gSettings.TableFont
    SaveSetting APP_NAME, SECTION_NAME, "TableSize", gSettings.TableSize
    SaveSetting APP_NAME, SECTION_NAME, "TableLineWidth", CStr(gSettings.TableLineWidth)
    SaveSetting APP_NAME, SECTION_NAME, "TableRowHeight", CStr(gSettings.TableRowHeight)
    SaveSetting APP_NAME, SECTION_NAME, "TableSpacing", CStr(gSettings.TableSpacing)
    SaveSetting APP_NAME, SECTION_NAME, "TableAutoFitWindow", IIf(gSettings.TableAutoFitWindow, "1", "0")

    SaveSetting APP_NAME, SECTION_NAME, "FigureTitleFont", gSettings.FigureTitleFont
    SaveSetting APP_NAME, SECTION_NAME, "FigureTitleSize", gSettings.FigureTitleSize
    SaveSetting APP_NAME, SECTION_NAME, "FigureTitleBold", IIf(gSettings.FigureTitleBold, "1", "0")
    SaveSetting APP_NAME, SECTION_NAME, "FigureTitleSpacing", CStr(gSettings.FigureTitleSpacing)
    SaveSetting APP_NAME, SECTION_NAME, "FigureTitleLeftIndentCm", CStr(gSettings.FigureTitleLeftIndentCm)
    SaveSetting APP_NAME, SECTION_NAME, "FigureTitleRightIndentCm", CStr(gSettings.FigureTitleRightIndentCm)
    SaveSetting APP_NAME, SECTION_NAME, "FigureTitleSpecialIndent", gSettings.FigureTitleSpecialIndent
    SaveSetting APP_NAME, SECTION_NAME, "FigureTitleSpecialIndentCm", CStr(gSettings.FigureTitleSpecialIndentCm)

    SaveSetting APP_NAME, SECTION_NAME, "Remember", IIf(gSettings.RememberSettings, "1", "0")
End Sub

'========================
' 3. 按任务书刷新当前文档
'========================

Public Sub ApplySettingsToDocument()
    Dim doc As Document
    Set doc = ActiveDocument

    ' ---------- 正文默认：小四、1.5 倍行距 ----------
    With doc.Styles(wdStyleNormal).Font
        .NameFarEast = gSettings.BodyFont
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = ConvertSizeToPoints(gSettings.BodySize)
    End With

    With doc.Styles(wdStyleNormal).ParagraphFormat
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = 12 * gSettings.BodySpacing

        ' 段前/段后（按“行”计算）：优先用 LineUnitBefore/After，兼容 WPS/Word
        Dim pf As Object
        Set pf = doc.Styles(wdStyleNormal).ParagraphFormat
        SafeSetPfProp pf, "LineUnitBefore", gSettings.BodyBeforeLines
        SafeSetPfProp pf, "LineUnitAfter", gSettings.BodyAfterLines

        Select Case gSettings.BodyAlignment
            Case "左对齐": .Alignment = wdAlignParagraphLeft
            Case "居中": .Alignment = wdAlignParagraphCenter
            Case "右对齐": .Alignment = wdAlignParagraphRight
            Case "两端对齐": .Alignment = wdAlignParagraphJustify
        End Select
        
        ApplyParagraphIndentByChars pf, gSettings.BodyLeftIndentCm, gSettings.BodyRightIndentCm, gSettings.BodySpecialIndent, gSettings.BodySpecialIndentCm

        ' 若宿主不支持 LineUnitBefore/After，则退回到 SpaceBefore/After
        Dim oneLinePt As Single
        oneLinePt = ConvertSizeToPoints(gSettings.BodySize) * gSettings.BodySpacing
        If gSettings.BodyBeforeLines <> 0 Then
            If SafeGetPfPropNumber(pf, "LineUnitBefore") = 0 Then
                pf.SpaceBefore = oneLinePt * gSettings.BodyBeforeLines
            End If
        End If
        If gSettings.BodyAfterLines <> 0 Then
            If SafeGetPfPropNumber(pf, "LineUnitAfter") = 0 Then
                pf.SpaceAfter = oneLinePt * gSettings.BodyAfterLines
            End If
        End If
    End With

    ' ---------- 自动识别“1.1/1.1.1...”标题并设置大纲级别（可选） ----------
    If gSettings.AutoDetectNumericTitles Then
        AutoDetectAndFormatNumericTitles doc
    End If

    ' ---------- 正文各级标题（按多级编号 1~7 级） ----------
    ' 保留自动编号，不将其转为纯文本
    RefreshContentFormat doc

    ' ---------- 表格 ----------
    RefreshTableFormat doc
    
    ' ---------- 图片标题 ----------
    RefreshFigureCaptions doc
    
    ' ---------- 统一英文和数字字体（正文/标题/目录/表格/图片标题等） ----------
    NormalizeAsciiFont doc
End Sub

'========================
' 4. 目录识别与样式
'========================

Private Sub RefreshDirectoryFormat(doc As Document)
    Dim para As Paragraph
    Dim txt As String
    Dim level As Integer
    
    For Each para In doc.Paragraphs
        txt = Trim$(para.Range.Text)

        ' 目录标题（“目录/目 录”）固定默认格式：宋体、小四、加粗、1.5 倍行距
        If IsDirectoryTitleLine(txt) Then
            ApplyDirectoryTitleStyle para
            GoTo NextPara
        End If

        level = IdentifyDirectoryLevel(txt)
        If level > 0 And level <= 3 Then
            ' 只要被视为目录行（自动 TOC 或手工目录）就按目录样式处理
            ' - IsTocEntryLine(txt): 手工目录（编号 + 文字 + 点引导符 + 页码）
            ' - IsTocParagraph(para): 使用 TOC1/TOC2/TOC3 等目录样式的自动目录
            If IsTocEntryLine(txt) Or IsTocParagraph(para) Then
                ' 先设置样式（TOC1~TOC3），再用直接格式覆盖字号/字体
                ApplyTocStyleToParagraph doc, para, level
                ApplyDirectoryStyle para, level
            End If
        End If
NextPara:
    Next para
End Sub

Private Function IsDirectoryTitleLine(ByVal txt As String) As Boolean
    Dim t As String
    t = Replace$(txt, vbCr, "")
    t = Replace$(t, vbLf, "")
    t = Trim$(t)
    t = Replace$(t, " ", "")
    t = Replace$(t, ChrW(12288), "") ' 全角空格
    IsDirectoryTitleLine = (t = "目录")
End Function

Private Sub ApplyDirectoryTitleStyle(para As Paragraph)
    With para.Range
        .Font.NameFarEast = "宋体"
        .Font.NameAscii = "Times New Roman"
        .Font.NameOther = "Times New Roman"
        .Font.Size = ConvertSizeToPoints("小四")
        .Font.Bold = True

        .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
        .ParagraphFormat.LineSpacing = 12 * 1.5
    End With
End Sub

' 判断是否为目录条目行：包含 Tab 且以页码结尾（如：1.1 XXXX<Tab>2）
' 说明：
'   - 这一版逻辑在 WPS 中经过验证是稳定可用的；
'   - 在 Word 中，同样适用于自动目录（编号 + 目录文字 + 制表符 + 页码）。
Private Function IsTocEntryLine(ByVal txt As String) As Boolean
    Dim t As String
    Dim lastToken As String
    Dim posTab As Long
    Dim i As Long
    Dim j As Long

    ' 清理换行符并去掉首尾空白
    t = Replace$(txt, vbCr, "")
    t = Replace$(t, vbLf, "")
    t = Trim$(t)

    If t = "" Then
        IsTocEntryLine = False
        Exit Function
    End If

    ' 必须包含一个制表符（编号/标题 与 页码 的分隔符）
    posTab = InStr(t, vbTab)
    If posTab = 0 Then
        IsTocEntryLine = False
        Exit Function
    End If

    ' 取最后一个“非空白 token”，判断是否为数字页码
    i = Len(t)
    Do While i > 0 And Mid$(t, i, 1) = " "
        i = i - 1
    Loop
    If i <= 0 Then
        IsTocEntryLine = False
        Exit Function
    End If

    j = i
    Do While j > 0 And Mid$(t, j, 1) <> " " And Mid$(t, j, 1) <> vbTab
        j = j - 1
    Loop
    lastToken = Mid$(t, j + 1, i - j)

    IsTocEntryLine = IsNumeric(lastToken)
End Function

' 判断某段落是否属于“目录”（应保持单倍行距、不按正文/标题处理）
Private Function IsTocParagraph(para As Paragraph) As Boolean
    Dim t As String
    Dim styleName As String
    
    t = Trim$(para.Range.Text)
    If IsTocEntryLine(t) Then
        IsTocParagraph = True
        Exit Function
    End If
    
    On Error Resume Next
    styleName = CStr(para.Style)
    On Error GoTo 0
    
    ' 样式名中包含 "TOC"（TOC1~TOC3 等）视为目录
    If styleName <> "" Then
        If InStr(1, styleName, "TOC", vbTextCompare) > 0 Then
            IsTocParagraph = True
            Exit Function
        End If
    End If
End Function

' 强制设置目录样式：目录 1/2/3（TOC1~TOC3）
Private Sub ApplyTocStyleToParagraph(doc As Document, para As Paragraph, ByVal level As Integer)
    On Error Resume Next
    Select Case level
        Case 1: para.Style = doc.Styles(wdStyleTOC1)
        Case 2: para.Style = doc.Styles(wdStyleTOC2)
        Case 3: para.Style = doc.Styles(wdStyleTOC3)
    End Select
    On Error GoTo 0
End Sub

' 识别 1.0 / 1.1 / 1.1.1 / 1.1.1.1 ... 对应级别
Private Function IdentifyDirectoryLevel(ByVal txt As String) As Integer
    Dim regEx As Object, m As Object, numberPart As String
    Dim dotCount As Integer

    Set regEx = CreateObject("VBScript.Regexp")
    regEx.Pattern = "^\d+(\.\d+)*"
    regEx.Global = False

    If regEx.Test(txt) Then
        Set m = regEx.Execute(txt)(0)
        numberPart = m.Value
        dotCount = Len(numberPart) - Len(Replace(numberPart, ".", ""))

        ' 1.0 / 2.0 判定为一级
        If Right$(numberPart, 2) = ".0" And dotCount = 1 Then
            IdentifyDirectoryLevel = 1
        Else
            IdentifyDirectoryLevel = dotCount + 1
        End If
    Else
    IdentifyDirectoryLevel = 0
    End If
End Function

Private Sub ApplyDirectoryStyle(para As Paragraph, ByVal level As Integer)
    With para.Range
        Select Case level
            Case 1
                .Font.NameFarEast = gSettings.Level1Font
                .Font.NameAscii = "Times New Roman"
                .Font.NameOther = "Times New Roman"
                .Font.Size = ConvertSizeToPoints(gSettings.Level1Size)
                .Font.Bold = gSettings.Level1Bold
                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                
            Case 2
                .Font.NameFarEast = gSettings.Level2Font
                .Font.NameAscii = "Times New Roman"
                .Font.NameOther = "Times New Roman"
                .Font.Size = ConvertSizeToPoints(gSettings.Level2Size)
                .Font.Bold = gSettings.Level2Bold
                .ParagraphFormat.LeftIndent = CentimetersToPoints(gSettings.Level2IndentChars * 0.35)
                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                
            Case 3
                .Font.NameFarEast = gSettings.Level3Font
                .Font.NameAscii = "Times New Roman"
                .Font.NameOther = "Times New Roman"
                .Font.Size = ConvertSizeToPoints(gSettings.Level3Size)
                .Font.Bold = gSettings.Level3Bold
                .ParagraphFormat.LeftIndent = CentimetersToPoints(gSettings.Level3IndentChars * 0.35)
                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        End Select
    End With
End Sub

'========================
' 5. 正文多级标题样式
'========================

Private Sub RefreshContentFormat(doc As Document)
    Dim para As Paragraph
    Dim level As Integer
    Dim txt As String

    ' 性能优化：合并原先两次全量遍历为一次遍历
    For Each para In doc.Paragraphs
        If Not IsTocParagraph(para) Then
            ' 正文段落
            If para.OutlineLevel = wdOutlineLevelBodyText Then
                With para.Range
                    .Font.NameFarEast = gSettings.BodyFont
                    .Font.NameAscii = "Times New Roman"
                    .Font.NameOther = "Times New Roman"
                    .Font.Size = ConvertSizeToPoints(gSettings.BodySize)
                    .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                    .ParagraphFormat.LineSpacing = 12 * gSettings.BodySpacing
                    ApplyParagraphIndentByChars .ParagraphFormat, gSettings.BodyLeftIndentCm, gSettings.BodyRightIndentCm, gSettings.BodySpecialIndent, gSettings.BodySpecialIndentCm
                    Select Case gSettings.BodyAlignment
                        Case "左对齐": .ParagraphFormat.Alignment = wdAlignParagraphLeft
                        Case "居中": .ParagraphFormat.Alignment = wdAlignParagraphCenter
                        Case "右对齐": .ParagraphFormat.Alignment = wdAlignParagraphRight
                        Case "两端对齐": .ParagraphFormat.Alignment = wdAlignParagraphJustify
                    End Select
                End With
            End If

            ' 标题段落
            level = CInt(para.OutlineLevel)
            If level >= wdOutlineLevel1 And level <= wdOutlineLevel5 Then
                txt = Trim$(para.Range.Text)
                ApplyContentTitleStyle para, level, txt
            End If
        End If
    Next para
End Sub

'========================
' 5.5 自动识别正文数字标题（基于文本）
'========================
' 多级列表识别能力：
' - IsMultiLevelListParagraph: 判断段落是否属于多级列表，并返回当前级别（1~9）
' - ReportAllMultiLevelListLevels: 扫描当前文档并输出识别结果
Public Function IsMultiLevelListParagraph(ByVal para As Paragraph, ByRef level As Integer, Optional ByRef listNumberText As String = "") As Boolean
    On Error Resume Next
    
    Dim lf As ListFormat
    level = 0
    listNumberText = ""
    
    Set lf = para.Range.ListFormat
    If lf Is Nothing Then Exit Function
    If lf.ListType = wdListNoNumbering Then Exit Function
    
    ' 编号文本（如 1 / 1.2 / 1.2.3）
    listNumberText = NormalizeNumberString(CStr(lf.ListString))
    
    ' 优先取 Word/WPS 给出的真实列表级别
    If lf.ListLevelNumber >= 1 And lf.ListLevelNumber <= 9 Then
        level = CInt(lf.ListLevelNumber)
        IsMultiLevelListParagraph = True
        Exit Function
    End If
    
    ' 兜底：由编号文本估算级别
    level = IdentifyContentTitleLevelFromNumber(listNumberText)
    If level >= 1 And level <= 9 Then
        IsMultiLevelListParagraph = True
    End If
    
    On Error GoTo 0
End Function

Public Sub ReportAllMultiLevelListLevels()
    On Error GoTo EH
    
    Dim doc As Document
    Dim para As Paragraph
    Dim lvl As Integer
    Dim numText As String
    Dim reportDoc As Document
    Dim lineText As String
    Dim hitCount As Long
    
    Set doc = ActiveDocument
    Set reportDoc = Documents.Add
    
    reportDoc.Range.Text = "多级列表识别结果" & vbCrLf & _
                           "文档：" & doc.Name & vbCrLf & _
                           "----------------------------------------" & vbCrLf
    
    hitCount = 0
    For Each para In doc.Paragraphs
        If IsMultiLevelListParagraph(para, lvl, numText) Then
            lineText = CleanParaText(para.Range.Text)
            If Len(lineText) > 80 Then lineText = Left$(lineText, 80) & "..."
            
            reportDoc.Range.InsertAfter _
                "段落序号=" & CStr(para.Range.Paragraphs(1).Range.ListFormat.ListValue) & _
                " | 列表级别=" & CStr(lvl) & _
                " | 编号=" & IIf(numText = "", "(空)", numText) & _
                " | 文本=" & lineText & vbCrLf
            hitCount = hitCount + 1
        End If
    Next para
    
    reportDoc.Range.InsertAfter "----------------------------------------" & vbCrLf & _
                                "共识别到多级列表段落：" & CStr(hitCount) & " 条"
    MsgBox "识别完成：共找到 " & CStr(hitCount) & " 条多级列表段落。结果已输出到新文档。", vbInformation
    Exit Sub
EH:
    MsgBox "识别多级列表时出错：" & Err.Number & " " & Err.Description, vbExclamation
    Err.Clear
End Sub

' 规则：
'   - "1.1 空格 XXX"      => 2 级标题
'   - "1.1.2 空格 XXX"    => 3 级标题
'   - "1.1.2.4 空格 XXX"  => 4 级标题
' 说明：
'   - 仅针对正文段落（OutlineLevel 为 BodyText）做提升
'   - 会跳过目录段落（TOC）
'   - 不改动“编号文本”，只做标题识别与格式刷
Private Sub AutoDetectAndFormatNumericTitles(doc As Document)
    Dim para As Paragraph
    Dim txt As String
    Dim level As Integer
    Dim listStr As String
    
    For Each para In doc.Paragraphs
        ' 只跳过目录（TOC）段落
        If Not IsTocParagraph(para) Then
            ' 是否包含“多级列表”段落：不勾选时，只处理 BodyText
            If (Not gSettings.AutoDetectIncludeListParagraphs) Then
                If para.OutlineLevel <> wdOutlineLevelBodyText Then GoTo NextPara
            End If

            txt = Trim$(para.Range.Text)

            ' 优先：若该段落是多级列表自动编号，则优先用真实列表级别识别
            level = 0
            If IsMultiLevelListParagraph(para, level, listStr) Then
                ' 这里得到的是真实多级列表级别（1~9）
                ' 为标题自动识别仅使用 1~5
                If level > 5 Then level = 0
                If level = 0 Then
                    level = IdentifyContentTitleLevelFromNumber(listStr)
                End If
            End If

            ' 次选：手工输入的 “1.1.2 空格 标题” 这种，把编号当成纯文本
            If level = 0 Then
                level = IdentifyContentTitleLevel(txt)
            End If

            If level >= 1 And level <= 5 Then
                On Error Resume Next
                para.OutlineLevel = level
                On Error GoTo 0

                ' 识别后立即刷一次标题格式（避免后续逻辑遗漏）
                ApplyContentTitleStyle para, level, txt
            End If
        End If
NextPara:
    Next para
End Sub

' 规范化编号字符串：兼容多级列表 ListString 可能带尾部点号/全角点号/空白等情况
Private Function NormalizeNumberString(ByVal s As String) As String
    Dim t As String
    t = Replace$(Replace$(Replace$(Trim$(s), vbCr, ""), vbLf, ""), "。", ".")
    ' 去掉末尾多余的点号（如 "1.1."）
    Do While Len(t) > 0 And Right$(t, 1) = "."
        t = Left$(t, Len(t) - 1)
    Loop
    NormalizeNumberString = t
End Function

' 从段落文本识别正文标题级别：返回 0 表示不匹配
Private Function IdentifyContentTitleLevel(ByVal txt As String) As Integer
    Dim t As String
    Dim regEx As Object, m As Object, numberPart As String
    Dim dotCount As Integer

    t = Replace$(txt, vbCr, "")
    t = Replace$(t, vbLf, "")
    t = Trim$(t)
    If t = "" Then
        IdentifyContentTitleLevel = 0
        Exit Function
    End If

    ' 必须是：数字( .数字 )... +（可选空格/制表符）+ 标题文字
    Set regEx = CreateObject("VBScript.Regexp")
    ' 允许：1 工程概况 / 1.1 工程概况 / 1.1工程概况（从一级开始）
    regEx.Pattern = "^(\d+(\.\d+){0,6})[ \t]*.+$"
    regEx.Global = False
    regEx.IgnoreCase = True

    If Not regEx.Test(t) Then
        IdentifyContentTitleLevel = 0
        Exit Function
    End If

    Set m = regEx.Execute(t)(0)
    numberPart = m.SubMatches(0)

    dotCount = Len(numberPart) - Len(Replace(numberPart, ".", ""))

    ' 若是纯一级编号（如 "1"），要求后面至少有空格/Tab，避免误识别 "2026年..." 这类正文
    If dotCount = 0 Then
        Dim nextPos As Integer
        Dim nextCh As String
        nextPos = Len(numberPart) + 1
        If nextPos <= Len(t) Then
            nextCh = Mid$(t, nextPos, 1)
            If nextCh <> " " And nextCh <> vbTab Then
                IdentifyContentTitleLevel = 0
                Exit Function
            End If
        Else
            IdentifyContentTitleLevel = 0
            Exit Function
        End If
    End If

    ' 兼容 "1.0 标题" 这种写法：作为 1 级标题
    If dotCount = 1 And Right$(numberPart, 2) = ".0" Then
        IdentifyContentTitleLevel = 1
    Else
        IdentifyContentTitleLevel = dotCount + 1
    End If
End Function

' 从“编号字符串”识别正文标题级别：返回 0 表示不匹配
' 例如：
'   "1.1" -> 2
'   "1.1.2" -> 3
'   "1.1.2.4" -> 4
'   "1.0" -> 1
Private Function IdentifyContentTitleLevelFromNumber(ByVal numStr As String) As Integer
    Dim t As String
    Dim regEx As Object, m As Object, numberPart As String
    Dim dotCount As Integer

    t = NormalizeNumberString(numStr)
    If t = "" Then
        IdentifyContentTitleLevelFromNumber = 0
        Exit Function
    End If

    Set regEx = CreateObject("VBScript.Regexp")
    regEx.Pattern = "^\d+(\.\d+){0,6}$"
    regEx.Global = False
    regEx.IgnoreCase = True

    If regEx.Test(t) Then
        Set m = regEx.Execute(t)(0)
        numberPart = m.Value
        dotCount = Len(numberPart) - Len(Replace(numberPart, ".", ""))
    Else
        ' 兜底：遇到少量非标准字符时，尽量只保留数字和点号再判断
        Dim cleaned As String
        cleaned = t
        cleaned = Replace$(cleaned, " ", "")
        cleaned = Replace$(cleaned, vbTab, "")
        ' 若清理后仍然包含非数字/点号，放弃
        Set regEx = CreateObject("VBScript.Regexp")
        regEx.Pattern = "^[0-9.]+$"
        regEx.Global = False
        If Not regEx.Test(cleaned) Then
            IdentifyContentTitleLevelFromNumber = 0
            Exit Function
        End If
        cleaned = NormalizeNumberString(cleaned)
        If cleaned = "" Then
            IdentifyContentTitleLevelFromNumber = 0
            Exit Function
        End If
        dotCount = Len(cleaned) - Len(Replace(cleaned, ".", ""))
        numberPart = cleaned
    End If

    If dotCount = 1 And Right$(numberPart, 2) = ".0" Then
        IdentifyContentTitleLevelFromNumber = 1
    Else
        IdentifyContentTitleLevelFromNumber = dotCount + 1
    End If
End Function

Private Sub ApplyContentTitleStyle(para As Paragraph, ByVal level As Integer, ByVal txt As String)
        With para.Range
        ' 字体/字号：按 1~4 级各自设置
        .Font.NameFarEast = gSettings.TitleFont(level)
        .Font.NameAscii = "Times New Roman"
        .Font.NameOther = "Times New Roman"
        .Font.Size = ConvertSizeToPoints(gSettings.TitleSize(level))
            
        ' 加粗：按设置
            .Font.Bold = gSettings.TitleBold(level)
            
        ' 行距：按该级标题设置的倍数
        .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
        .ParagraphFormat.LineSpacing = 12 * IIf(gSettings.TitleLineSpacing(level) > 0, gSettings.TitleLineSpacing(level), gSettings.BodySpacing)
            
        ' 对齐方式
            Select Case gSettings.TitleAlignment(level)
            Case "左对齐": .ParagraphFormat.Alignment = wdAlignParagraphLeft
            Case "居中": .ParagraphFormat.Alignment = wdAlignParagraphCenter
            Case "右对齐": .ParagraphFormat.Alignment = wdAlignParagraphRight
            Case "两端对齐": .ParagraphFormat.Alignment = wdAlignParagraphJustify
            End Select
            
        ' 段前/段后：优先用“行单位”属性（更贴近 WPS/Word 的 UI：0.5 行、1 行等）
        ' 注意：WPS 不同版本对象库不一定有 SnapToGrid/DisableLineHeightGrid 等属性，必须用 CallByName 规避编译错误
        Dim pf As Object
        Set pf = .ParagraphFormat

        ' 标题缩进固定：左侧=0，右侧=0，特殊格式=无
        pf.LeftIndent = 0
        pf.RightIndent = 0
        pf.FirstLineIndent = 0
        SafeSetPfProp pf, "CharacterUnitLeftIndent", 0
        SafeSetPfProp pf, "CharacterUnitRightIndent", 0
        SafeSetPfProp pf, "CharacterUnitFirstLineIndent", 0

        SafeSetPfProp pf, "SpaceBeforeAuto", False
        SafeSetPfProp pf, "SpaceAfterAuto", False
        SafeSetPfProp pf, "SnapToGrid", False
        SafeSetPfProp pf, "DisableLineHeightGrid", True

        SafeSetPfProp pf, "LineUnitBefore", gSettings.TitleBeforeLines(level)
        SafeSetPfProp pf, "LineUnitAfter", gSettings.TitleAfterLines(level)

        ' 兜底：若宿主不支持 LineUnitBefore/After，则按磅值折算
        If gSettings.TitleBeforeLines(level) <> 0 Or gSettings.TitleAfterLines(level) <> 0 Then
            Dim oneLinePt As Single
            oneLinePt = ConvertSizeToPoints(gSettings.TitleSize(level)) * IIf(gSettings.TitleLineSpacing(level) > 0, gSettings.TitleLineSpacing(level), gSettings.BodySpacing)
            On Error Resume Next
            If SafeGetPfPropNumber(pf, "LineUnitBefore") = 0 And gSettings.TitleBeforeLines(level) <> 0 Then
                pf.SpaceBefore = oneLinePt * gSettings.TitleBeforeLines(level)
            End If
            If SafeGetPfPropNumber(pf, "LineUnitAfter") = 0 And gSettings.TitleAfterLines(level) <> 0 Then
                pf.SpaceAfter = oneLinePt * gSettings.TitleAfterLines(level)
            End If
            On Error GoTo 0
        End If
    End With

    ' 同步多级列表“编号”的字体和字号为正文字体/字号
    SyncNumberFontWithBody para
End Sub

Private Sub SafeSetPfProp(ByVal pf As Object, ByVal propName As String, ByVal propValue As Variant)
    On Error Resume Next
    CallByName pf, propName, VbLet, propValue
    On Error GoTo 0
End Sub

Private Function SafeGetPfPropNumber(ByVal pf As Object, ByVal propName As String) As Double
    On Error Resume Next
    SafeGetPfPropNumber = CDbl(CallByName(pf, propName, VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        SafeGetPfPropNumber = 0
            End If
    On Error GoTo 0
End Function

' 按“字符”应用缩进（兼容 Word/WPS）：
' - 优先使用 CharacterUnit* 属性
' - 若宿主不支持，则按 1 字符≈0.35cm 折算为磅值兜底
Private Sub ApplyParagraphIndentByChars(ByVal pf As Object, ByVal leftChars As Single, ByVal rightChars As Single, ByVal specialType As String, ByVal specialChars As Single)
    Dim specialCharsSigned As Single
    
    Select Case specialType
        Case "首行缩进"
            specialCharsSigned = specialChars
        Case "悬挂缩进"
            specialCharsSigned = -specialChars
        Case Else
            specialCharsSigned = 0
    End Select
    
    SafeSetPfProp pf, "CharacterUnitLeftIndent", leftChars
    SafeSetPfProp pf, "CharacterUnitRightIndent", rightChars
    SafeSetPfProp pf, "CharacterUnitFirstLineIndent", specialCharsSigned
    
    ' 兜底：若 CharacterUnit* 不生效，则换算为 points
    If leftChars = 0 Then
        pf.LeftIndent = 0
    ElseIf SafeGetPfPropNumber(pf, "CharacterUnitLeftIndent") = 0 Then
        pf.LeftIndent = CentimetersToPoints(leftChars * 0.35)
    End If
    
    If rightChars = 0 Then
        pf.RightIndent = 0
    ElseIf SafeGetPfPropNumber(pf, "CharacterUnitRightIndent") = 0 Then
        pf.RightIndent = CentimetersToPoints(rightChars * 0.35)
    End If
    
    If specialCharsSigned = 0 Then
        pf.FirstLineIndent = 0
    ElseIf SafeGetPfPropNumber(pf, "CharacterUnitFirstLineIndent") = 0 Then
        pf.FirstLineIndent = CentimetersToPoints(specialCharsSigned * 0.35)
    End If
End Sub

Private Sub SyncNumberFontWithBody(para As Paragraph)
    On Error Resume Next

    Dim lf As ListFormat
    Dim lvl As ListLevel
    Dim lvlNo As Long
    Dim titleLevel As Long
    Dim indentChars As Long
    Dim indentPt As Single

    Set lf = para.Range.ListFormat
    If lf Is Nothing Then Exit Sub
    If lf.ListType = wdListNoNumbering Then Exit Sub

    ' 当前段落所处的多级列表级别
    lvlNo = lf.ListLevelNumber
    If lvlNo < 1 Or lvlNo > 9 Then Exit Sub

    Set lvl = lf.ListTemplate.ListLevels(lvlNo)
    If lvl Is Nothing Then Exit Sub

    ' 把编号（多级列表）的字体和字号改成与正文一致
    With lvl.Font
        .NameFarEast = gSettings.BodyFont
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = ConvertSizeToPoints(gSettings.BodySize)
    End With

    ' 不在此处统一修改多级列表缩进/对齐（避免覆盖用户在 WPS “定义新多级列表”里的缩进设置）

    On Error GoTo 0
End Sub

Private Sub NormalizeAsciiFont(doc As Document)
    On Error Resume Next

    Dim storyRng As Range
    Dim rng As Range

    ' 遍历正文、脚注、尾注、页眉页脚等所有故事范围
    For Each storyRng In doc.StoryRanges
        Set rng = storyRng.Duplicate
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            ' 匹配所有半角英文字符和数字
            .Text = "[0-9A-Za-z]"
            .Replacement.Text = "^&"
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = True

            ' 将西文字体统一为 Times New Roman
            .Replacement.Font.NameAscii = "Times New Roman"
            .Replacement.Font.NameOther = "Times New Roman"

            .Execute Replace:=wdReplaceAll
        End With
    Next storyRng

    On Error GoTo 0
End Sub

'========================
' 6. 括号和句号（中英文判断）
'========================

Private Sub ApplyBrackets(para As Paragraph, ByVal txt As String)
    Dim numberPartEnd As Long
    Dim titleText As String
    Dim isChinese As Boolean
    Dim i As Long
    Dim ch As String

    ' 已存在括号就不处理
    If InStr(txt, "(") > 0 Or InStr(txt, ")") > 0 _
       Or InStr(txt, "（") > 0 Or InStr(txt, "）") > 0 Then
        Exit Sub
    End If

    ' 取出编号后正文部分（假定“编号+空格+标题”）
    numberPartEnd = InStr(txt, " ")
    If numberPartEnd = 0 Then Exit Sub

    titleText = Mid$(txt, numberPartEnd + 1)

    ' 简单判断是否包含中文字符
    isChinese = False
    For i = 1 To Len(titleText)
        ch = Mid$(titleText, i, 1)
        If AscW(ch) < 0 Then
            isChinese = True
            Exit For
        End If
    Next i

    If isChinese Then
        para.Range.Text = left$(txt, numberPartEnd) & "（" & Trim$(titleText) & "）"
    Else
        para.Range.Text = left$(txt, numberPartEnd) & " (" & Trim$(titleText) & ")"
    End If
End Sub

Private Sub AddPeriodIfNeeded(para As Paragraph)
    Dim txt As String
    txt = para.Range.Text
    
    If Right$(Trim$(txt), 1) <> "。" Then
        para.Range.Text = RTrim$(txt) & "。"
    End If
End Sub

'========================
' 7. 表格格式
'========================

Private Sub RefreshTableFormat(doc As Document)
    Dim tbl As Table
    Dim idx As Long

    idx = 1
    For Each tbl In doc.Tables
        ' 允许表格随内容/窗口自动调整（行高/列宽）
        On Error Resume Next
        tbl.AllowAutoFit = True
        ' 使用数值枚举避免 WPS 缺少常量名导致编译失败：
        '   1 = wdAutoFitContent（根据内容）
        '   2 = wdAutoFitWindow（根据窗口）
        If gSettings.TableAutoFitWindow Then
            tbl.AutoFitBehavior 2
        Else
            tbl.AutoFitBehavior 1
        End If
        On Error GoTo 0

        ' 设置表格内文字
        With tbl.Range.Font
            .NameFarEast = gSettings.TableFont
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
            .Size = ConvertSizeToPoints(gSettings.TableSize)
            .Bold = False
        End With

        ' 行高：改为“自动/至少”，使之随文字大小与内容长度自适应
        Dim iRow As Long
        On Error Resume Next
        For iRow = 1 To tbl.Rows.Count
            ' 优先：至少高度（保留一个最小值），内容更多时可自动变高
            ' 使用数值枚举避免 WPS 缺少常量名导致编译失败：wdRowHeightAtLeast=1
            tbl.Rows(iRow).HeightRule = 1
            tbl.Rows(iRow).height = CentimetersToPoints(gSettings.TableRowHeight)
        Next iRow
        On Error GoTo 0

        ' 表格内段落行距：按设置倍数（默认 1.0），Word/WPS 均支持
        With tbl.Range.ParagraphFormat
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = 12 * IIf(gSettings.TableSpacing > 0, gSettings.TableSpacing, 1#)
            ' 统一段前/段后，避免同一行文字上下不齐
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With

        ' 线宽：使用底层枚举值 2（等价于 Word 的 wdLineWidth025pt），避免在 WPS 中找不到常量名
        On Error Resume Next
        tbl.Borders.OutsideLineWidth = 2
        tbl.Borders.InsideLineWidth = 2
        On Error GoTo 0
            
        ' 单元格内容居中（水平）+ 垂直居中，避免文字高低不统一
            tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        On Error Resume Next
        ' wdCellAlignVerticalCenter=1（用数值避免 WPS 常量缺失）
        tbl.Range.Cells.VerticalAlignment = 1
        On Error GoTo 0

        ' 表格标题（简单用“表 1、表 2、…”）
        SetTableTitle tbl, idx
        idx = idx + 1
    Next tbl
End Sub

Private Sub SetTableTitle(tbl As Table, ByVal index As Long)
    ' 目标：格式化“表格名称/题注”这一行（例如：表 1.4-1 文件构成表）
    ' - 优先识别表格前一段（更常见的题注位置）
    ' - 找不到时才退回表格起始段
    ' - 不修改题注文本（不再强制写入“表 1、表 2…”），避免破坏 1.4-1 这种章节编号
    Dim capRng As Range
    Set capRng = FindTableCaptionRange(tbl, 10)

    ' 找不到题注就直接退出：宁可不刷，也不要误把表格内容当题注
    If capRng Is Nothing Then Exit Sub

    With capRng
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
        .ParagraphFormat.LineSpacing = 12 * IIf(gSettings.TableTitleSpacing > 0, gSettings.TableTitleSpacing, 1#)
        ApplyParagraphIndentByChars .ParagraphFormat, gSettings.TableTitleLeftIndentCm, gSettings.TableTitleRightIndentCm, gSettings.TableTitleSpecialIndent, gSettings.TableTitleSpecialIndentCm
        .Font.NameFarEast = gSettings.TableTitleFont
        .Font.NameAscii = "Times New Roman"
        .Font.NameOther = "Times New Roman"
        .Font.Size = ConvertSizeToPoints(gSettings.TableTitleSize)
        .Font.Bold = gSettings.TableTitleBold
    End With
End Sub

' 查找表格题注的 Range：先找普通段落，再找文本框(Shape.TextFrame)中的题注
' maxUpParas: 向上最多检查多少个段落
Private Function FindTableCaptionRange(tbl As Table, ByVal maxUpParas As Integer) As Range
    Dim p As Paragraph
    Dim i As Integer
    Dim t As String
    Dim doc As Document
    Dim startPos As Long

    Set doc = tbl.Parent

    ' 1) 普通段落：从表格起始段落向上找
    On Error Resume Next
    Set p = tbl.Range.Paragraphs(1)
    On Error GoTo 0

    If Not p Is Nothing Then
        For i = 1 To maxUpParas
            On Error Resume Next
            Set p = p.Previous
            On Error GoTo 0
            If p Is Nothing Then Exit For

            t = CleanParaText(p.Range.Text)
            If t = "" Then
                ' 空段跳过
            ElseIf IsTableCaptionLine(t) Then
                Set FindTableCaptionRange = p.Range
                Exit Function
            End If
        Next i
    End If

    ' 2) 文本框/形状：很多 WPS 模板把“表1.5-1 …”放在文本框里
    startPos = 0
    On Error Resume Next
    startPos = tbl.Range.Start
    On Error GoTo 0

    Set FindTableCaptionRange = FindCaptionInShapes(doc, startPos)
End Function

Private Function FindCaptionInShapes(doc As Document, ByVal tableStartPos As Long) As Range
    On Error Resume Next

    Dim shp As Shape
    Dim txt As String
    Dim rng As Range
    Dim best As Range
    Dim bestDist As Long
    Dim dist As Long

    bestDist = 2147483647

    For Each shp In doc.Shapes
        If shp.TextFrame.HasText Then
            txt = shp.TextFrame.TextRange.Text
            txt = CleanParaText(txt)

            If IsTableCaptionLine(txt) Then
                ' 用锚点与表格起始位置的距离做“就近匹配”
                dist = Abs(CLng(shp.Anchor.Start) - CLng(tableStartPos))
                If dist < bestDist Then
                    bestDist = dist
                    Set rng = shp.TextFrame.TextRange.Duplicate
                    Set best = rng
                End If
            End If
        End If
    Next shp

    Set FindCaptionInShapes = best
End Function

'========================
' 8.1 图片标题格式
'========================
Private Sub RefreshFigureCaptions(doc As Document)
    Dim para As Paragraph
    Dim t As String

    For Each para In doc.Paragraphs
        t = CleanParaText(para.Range.Text)
        If IsFigureCaptionLine(t) Then
            With para.Range
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
                .ParagraphFormat.LineSpacing = 12 * IIf(gSettings.FigureTitleSpacing > 0, gSettings.FigureTitleSpacing, 1#)
                ApplyParagraphIndentByChars .ParagraphFormat, gSettings.FigureTitleLeftIndentCm, gSettings.FigureTitleRightIndentCm, gSettings.FigureTitleSpecialIndent, gSettings.FigureTitleSpecialIndentCm
                .Font.NameFarEast = gSettings.FigureTitleFont
                .Font.NameAscii = "Times New Roman"
                .Font.NameOther = "Times New Roman"
                .Font.Size = ConvertSizeToPoints(gSettings.FigureTitleSize)
                .Font.Bold = gSettings.FigureTitleBold
            End With
        End If
    Next para
End Sub

Private Function CleanParaText(ByVal s As String) As String
    Dim t As String
    t = Replace$(s, vbCr, "")
    t = Replace$(t, vbLf, "")
    t = Trim$(t)
    ' 兼容全角空格
    t = Replace$(t, ChrW(12288), " ")
    ' 合并多空格/Tab 不在这里做，交给正则
    CleanParaText = t
End Function

' 表格题注识别：
'   前方有回车（即为单独段落）
'   组成模式："表" + "数字(可含 . 和 - )" + "空白(空格或Tab)" + "文字"
' 示例：表 1.4-1 文件构成表
Private Function IsTableCaptionLine(ByVal paraText As String) As Boolean
    Dim t As String
    Static re As Object

    t = CleanParaText(paraText)

    If t = "" Then
        IsTableCaptionLine = False
        Exit Function
    End If

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = False
        re.IgnoreCase = False
        ' 允许：表1.5-1  标题 / 表 1.5-1\t标题 / 表1.4-1 文件构成表
        ' 数字部分：章号.节号-序号（也兼容更多点号）
        re.Pattern = "^表\s*\d+(\.\d+)*-\d+[\t ]+.+$"
    End If

    IsTableCaptionLine = re.Test(t)
End Function

' 图片题注识别：
'   组成模式："图" + "数字(可含 . 和 - )" + "空白(空格或Tab)" + "文字"
' 示例：图1.5-1  西气东输三线闽粤支干线（潮州-27#阀室段）线路走向示意图
Private Function IsFigureCaptionLine(ByVal paraText As String) As Boolean
    Dim t As String
    Static re As Object

    t = CleanParaText(paraText)

    If t = "" Then
        IsFigureCaptionLine = False
        Exit Function
    End If

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = False
        re.IgnoreCase = False
        ' 允许：图1.5-1  标题 / 图 1.5-1\t标题 / 图1.5-1 标题
        re.Pattern = "^图\s*\d+(\.\d+)*-\d+[\t ]+.+$"
    End If

    IsFigureCaptionLine = re.Test(t)
End Function

'========================
' 8. 公共辅助函数
'========================
Private Function ConvertSizeToPoints(ByVal sizeStr As String) As Single
    Select Case Trim$(sizeStr)
        Case "初号": ConvertSizeToPoints = 42
        Case "小初": ConvertSizeToPoints = 36
        Case "一号": ConvertSizeToPoints = 26
        Case "小一": ConvertSizeToPoints = 24
        Case "二号": ConvertSizeToPoints = 22
        Case "小二": ConvertSizeToPoints = 18
        Case "三号": ConvertSizeToPoints = 16
        Case "小三": ConvertSizeToPoints = 15
        Case "四号": ConvertSizeToPoints = 14
        Case "小四": ConvertSizeToPoints = 12
        Case "五号": ConvertSizeToPoints = 10.5
        Case "小五": ConvertSizeToPoints = 9
        Case "六号": ConvertSizeToPoints = 7.5
        Case "小六": ConvertSizeToPoints = 6.5
        Case "七号": ConvertSizeToPoints = 5.5
        Case "八号": ConvertSizeToPoints = 5
        Case Else
            If IsNumeric(sizeStr) Then
            ConvertSizeToPoints = Val(sizeStr)
            Else
                ConvertSizeToPoints = 12
            End If
    End Select
End Function

Private Function CentimetersToPoints(ByVal cm As Single) As Single
    CentimetersToPoints = cm * 28.35
End Function

