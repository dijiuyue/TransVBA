
Option Explicit

' 是否允许编辑设置（默认不允许，防止误触）
Private Const EDIT_TOGGLE_NAME As String = "chkEnableEdit"
Private m_LastEditToggleValue As Boolean
Private m_TitleFrameScrollInited As Boolean
Private Const REMEMBER_TOGGLE_NAME As String = "chkRemember"

' 通过页签标题安全获取 MultiPage.Page（避免 Pages(index) 下标越界）
' - 优先按 caption 精确匹配
' - 找不到则用 fallbackIndex（若存在）
' - 仍不存在则自动补建该页
Private Function GetOrCreatePage(ByVal mp As Object, ByVal caption As String, ByVal fallbackIndex As Integer) As Object
    Dim i As Long
    On Error Resume Next
    For i = 0 To mp.Pages.Count - 1
        If mp.Pages(i).caption = caption Then
            Set GetOrCreatePage = mp.Pages(i)
            On Error GoTo 0
            Exit Function
        End If
    Next i
    On Error GoTo 0
    
    If fallbackIndex >= 0 Then
        On Error Resume Next
        If mp.Pages.Count - 1 >= fallbackIndex Then
            Set GetOrCreatePage = mp.Pages(fallbackIndex)
            If Not GetOrCreatePage Is Nothing Then Exit Function
        End If
        On Error GoTo 0
    End If
    
    Set GetOrCreatePage = mp.Pages.Add(caption)
End Function

' 获取窗体上的主 MultiPage 实例：
' - 优先使用控件名 `mpSettings`
' - 若不存在（例如设计器里原本叫 MultiPage1），则取第一个 MultiPage
Private Function GetMainMultiPage() As Object
    Dim i As Integer
    On Error Resume Next
    Set GetMainMultiPage = Me.Controls("mpSettings")
    On Error GoTo 0
    
    If GetMainMultiPage Is Nothing Then
        For i = 0 To Me.Controls.Count - 1
            If TypeName(Me.Controls(i)) = "MultiPage" Then
                Set GetMainMultiPage = Me.Controls(i)
                Exit For
            End If
        Next i
    End If
End Function

'（移除 DbgLog 调试日志，避免宿主差异导致编译异常）

' 在UserForm的Initialize事件中设置界面
Private Sub UserForm_Initialize()
    Dim stepName As String
    On Error GoTo EH
    stepName = "设置窗体属性"
    ' 设置UserForm属性
    Me.caption = "Word文档格式自动刷新设置"
    ' 按比例继续缩小整体尺寸：减少右侧空隙
    Me.width = 610
    Me.height = 455
    
    ' 查找现有的MultiPage控件或创建新的
    Dim mpSettings As Object
    Dim i As Integer
    Dim found As Boolean
    
    ' 先尝试查找现有的MultiPage控件
    found = False
    For i = 0 To UserForm1.Controls.Count - 1
        If TypeName(UserForm1.Controls(i)) = "MultiPage" Then
            Set mpSettings = UserForm1.Controls(i)
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        ' 如果没有找到MultiPage控件，创建一个新的
        Set mpSettings = UserForm1.Controls.Add("Forms.MultiPage.1", "mpSettings", True)
    End If
    
    stepName = "配置 MultiPage（宽高/清空页签）"
    With mpSettings
        ' 内框跟随窗体宽度缩放，避免右侧出现大空隙
        .width = Me.width - 30
        ' 下移一点，给顶部“更改设置选项”留位置
        .top = 35
        .left = 10

        .height = 360
        
        ' 清除所有现有页面（如果有）
        ' 部分环境下 Pages.Remove(0) 在循环中可能触发异常，这里用“从后往前删 + 失败兜底”
        On Error Resume Next
        Do While .Pages.Count > 0
            .Pages.Remove .Pages.Count - 1
            If Err.Number <> 0 Then
                Err.Clear
                Exit Do
            End If
        Loop
        On Error GoTo EH
    End With
    
    ' 创建页面（删除“目录设置”页，仅保留默认目录规则）
    stepName = "创建正文页"
    CreateContentPage mpSettings
    stepName = "创建标题页"
    CreateTitlePage mpSettings
    stepName = "创建表格页"
    CreateTablePage mpSettings
    stepName = "创建图片标题页"
    CreateFigurePage mpSettings

    ' --- 顶部“更改设置选项”开关：默认锁定所有设置框 ---
    stepName = "创建顶部开关"
    EnsureEditToggle mpSettings
    EnsureRememberToggle mpSettings

    ' --- 调整已在设计器中放置的按钮位置，确保不被 MultiPage 遮挡且在可视区域内 ---
    On Error Resume Next
    With Me
        ' 将底部按钮上移，避免在小屏幕上被挤出
        .btnApply.top = mpSettings.top + mpSettings.height + 10
        .btnOK.top = .btnApply.top
        .btnCancel.top = .btnApply.top
        
        .btnApply.left = 220
        .btnOK.left = 310
        .btnCancel.left = 400
        
        .btnApply.Visible = True
        .btnOK.Visible = True
        .btnCancel.Visible = True

        ' 外层窗体纵向滚动条兜底：小屏时确保按钮区/页头不被裁切
        .ScrollBars = fmScrollBarsVertical
        .ScrollTop = 0
        .ScrollHeight = .btnCancel.top + .btnCancel.height + 30
    End With
    On Error GoTo 0

    ' 初始化下拉框选项和默认值
    stepName = "初始化下拉框"
    InitializeComboBoxes
    
    ' 从注册表加载上次保存的设置到全局变量
    stepName = "加载注册表设置"
    LoadSettingsFromRegistry
    
    ' 把设置回填到界面
    stepName = "回填设置到界面"
    LoadSettingsToForm

    ' 默认不允许编辑（只读预览），需勾选“更改设置选项”才解锁
    SetEditingEnabled mpSettings, False
    m_LastEditToggleValue = False
    
    Exit Sub
EH:
    MsgBox "UserForm_Initialize 发生错误：" & vbCrLf & _
           "Err=" & Err.Number & " " & Err.Description & vbCrLf & _
           "发生步骤：" & stepName, vbExclamation
    Err.Clear
End Sub

Private Sub EnsureRememberToggle(ByVal mpSettings As Object)
    Dim ctl As Object

    On Error Resume Next
    Set ctl = Me.Controls(REMEMBER_TOGGLE_NAME)
    On Error GoTo 0

    If ctl Is Nothing Then
        Set ctl = Me.Controls.Add("Forms.CheckBox.1", REMEMBER_TOGGLE_NAME, True)
    End If

    With ctl
        .caption = "记忆本次设置"
        .left = mpSettings.left + 160
        .top = 10
        .width = 120
        .Value = True
        .Visible = True
        .Enabled = True
    End With
End Sub

Private Sub EnsureEditToggle(ByVal mpSettings As Object)
    Dim ctl As Object

    On Error Resume Next
    Set ctl = Me.Controls(EDIT_TOGGLE_NAME)
    On Error GoTo 0

    If ctl Is Nothing Then
        Set ctl = Me.Controls.Add("Forms.CheckBox.1", EDIT_TOGGLE_NAME, True)
    End If

    With ctl
        .caption = "更改设置选项"
        .left = mpSettings.left
        .top = 10
        .width = 140
        .Value = False
        .Visible = True
        .Enabled = True
    End With
End Sub

' 兼容 WPS/Word：动态创建的 MSForms 控件在不同宿主下事件绑定不稳定，
' 这里用“轻量轮询”的方式同步复选框状态（鼠标移动/键盘操作时触发）。
Private Sub SyncEditToggleState()
    Dim v As Boolean
    Dim mp As Object

    On Error Resume Next
    v = (Me.Controls(EDIT_TOGGLE_NAME).Value = True)
    Set mp = GetMainMultiPage()
    On Error GoTo 0

    If v <> m_LastEditToggleValue Then
        SetEditingEnabled mp, v
        m_LastEditToggleValue = v
    End If
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SyncEditToggleState
End Sub

' 在窗体激活后重新计算内框滚动范围：
' MSForms 在创建时/布局尚未稳定时，Frame 的 ScrollWidth/ScrollHeight 可能导致滚动条位置异常（看得到但无法正确拖动）。
Private Sub UserForm_Activate()
    On Error Resume Next
    If m_TitleFrameScrollInited Then Exit Sub
    m_TitleFrameScrollInited = True

    Dim mp As Object
    Dim frame As Object
    Dim maxRight As Single, maxBottom As Single
    Dim c As Object

    Set mp = GetMainMultiPage()
    Set frame = GetOrCreatePage(mp, "标题设置", 1).Controls("frameTitleConfig")

    If Not frame Is Nothing Then
        ' 若标题页不使用滚动条，则不执行滚动范围重算
        If frame.ScrollBars <> fmScrollBarsVertical Then
            On Error GoTo 0
            Exit Sub
        End If
        maxRight = 0
        maxBottom = 0
        For Each c In frame.Controls
            If (c.Left + c.Width) > maxRight Then maxRight = c.Left + c.Width
            If (c.Top + c.Height) > maxBottom Then maxBottom = c.Top + c.Height
        Next c

        frame.ScrollBars = fmScrollBarsVertical
        frame.ScrollHeight = CLng(maxBottom + 240)
        frame.ScrollTop = 0
    End If
    On Error GoTo 0
End Sub

Private Sub SetEditingEnabled(ByVal mp As Object, ByVal enabled As Boolean)
    Dim p As Object
    Dim c As Object
    Dim subCtl As Object

    ' MultiPage 内的交互控件统一启用/禁用；Label 不处理
    For Each p In mp.Pages
        For Each c In p.Controls
            ' Label 直接跳过
            If TypeName(c) <> "Label" Then
                ' 保持外层 Frame 启用：否则滚动条会“看得到但无法拖动”
                If TypeName(c) = "Frame" Then
                    On Error Resume Next
                    c.Enabled = True
                    ' 但 Frame 内部的控件仍按 enabled 进行启用/禁用
                    For Each subCtl In c.Controls
                        If TypeName(subCtl) <> "Label" Then
                            subCtl.Enabled = enabled
                        End If
                    Next subCtl
                    On Error GoTo 0
                Else
                    On Error Resume Next
                    c.Enabled = enabled
                    On Error GoTo 0
                End If
            End If
        Next c
    Next p
End Sub

Private Sub CreateContentPage(ByRef mp As Object)
    Dim page As Object
    Set page = GetOrCreatePage(mp, "正文设置", -1)
    
    Dim frame As Object
    On Error Resume Next
    page.Controls.Remove "frameContent"
    On Error GoTo 0
    Set frame = page.Controls.Add("Forms.Frame.1", "frameContent", True)
    With frame
        .caption = "正文格式设置"
        .width = mp.Width - 40
        .height = 300
        .top = 10
        .left = 10
    End With
    
    ' =====================
    ' 正文默认格式（两行排版，避免被截断）
    ' =====================
    AddLabel frame, "lblBody", "正文默认格式:", 10, 20
    
    ' 第 1 行：字体 / 字号 / 行间距
    AddLabel frame, "lblBodyFont", "字体:", 20, 50
    AddComboBox frame, "cmbBodyFont", 60, 45, 100, 20
    
    AddLabel frame, "lblBodySize", "字号:", 180, 50
    AddTextBox frame, "txtBodySize", 220, 45, 40, 20, "小四"
    
    AddLabel frame, "lblBodySpacing", "行间距:", 280, 50
    AddTextBox frame, "txtBodySpacing", 330, 45, 40, 20, "1.5"
    
    ' 第 2 行：段前 / 段后 / 对齐
    AddLabel frame, "lblBodyBefore", "段前(行):", 20, 80
    AddTextBox frame, "txtBodyBefore", 80, 75, 40, 20, "0"
    
    AddLabel frame, "lblBodyAfter", "段后(行):", 140, 80
    AddTextBox frame, "txtBodyAfter", 200, 75, 40, 20, "0"
    
    AddLabel frame, "lblBodyAlign", "对齐:", 260, 80
    AddComboBox frame, "cmbBodyAlign", 300, 75, 80, 20
    
    ' 第 3 行：左/右缩进 + 特殊格式 + 缩进值
    AddLabel frame, "lblBodyLeftIndent", "左缩进(字符):", 20, 110
    AddTextBox frame, "txtBodyLeftIndent", 90, 105, 40, 20, "0"
    
    AddLabel frame, "lblBodyRightIndent", "右缩进(字符):", 150, 110
    AddTextBox frame, "txtBodyRightIndent", 220, 105, 40, 20, "0"
    
    AddLabel frame, "lblBodySpecial", "特殊格式:", 20, 140
    AddComboBox frame, "cmbBodySpecial", 80, 135, 90, 20
    
    AddLabel frame, "lblBodySpecialIndent", "缩进值(字符):", 200, 140
    AddTextBox frame, "txtBodySpecialIndent", 270, 135, 40, 20, "0"

    '（调试日志已移除）
    
    ' 在“正文设置”内框启用滚动条（放到正文页）
    On Error Resume Next
    Dim maxRightC As Single, maxBottomC As Single
    Dim cc As Object
    maxRightC = 0
    maxBottomC = 0
    For Each cc In frame.Controls
        If (cc.Left + cc.Width) > maxRightC Then maxRightC = cc.Left + cc.Width
        If (cc.Top + cc.Height) > maxBottomC Then maxBottomC = cc.Top + cc.Height
    Next cc
    
    frame.ScrollBars = fmScrollBarsVertical
    ' 最少保证滚动范围足够，避免只“显示”不“可拖动”
    frame.ScrollHeight = CLng(IIf(maxBottomC + 120 > frame.Height + 20, maxBottomC + 120, frame.Height + 80))
    frame.ScrollTop = 0
    On Error GoTo 0
    
End Sub

Private Sub CreateTitlePage(ByRef mp As Object)
    Dim page As Object
    Set page = GetOrCreatePage(mp, "标题设置", -1)

    Dim frame As Object
    On Error Resume Next
    page.Controls.Remove "frameTitleConfig"
    On Error GoTo 0
    Set frame = page.Controls.Add("Forms.Frame.1", "frameTitleConfig", True)
    With frame
        .caption = "标题格式设置"
        ' 外框（可视区域）尺寸：用整数避免 MSForms 在小数单位下的滚动计算异常
        .width = mp.Width - 40
        .height = 360
        .top = 0
        .left = 10
        ' 滚动条放在内框：横向+纵向（后面创建完子控件后再重新计算范围并设置）
    End With

    Dim i As Integer
    For i = 1 To 5
        ' 整体上移：相对正文页减少 topOffset
        CreateTitleLevelGroup frame, i, 20
    Next i

    ' 同步上移底部开关（保证在不使用滚动条的情况下也能完整显示）
    AddCheckBox frame, "chkAutoDetectNumericTitles", 20, 310, 360, 20, "自动识别 1/1.1/1.1.1/... 并设置为 1~4 级标题（大纲级别）"
    AddCheckBox frame, "chkAutoDetectIncludeList", 40, 330, 360, 20, "包含“多级列表”段落（也自动改为对应大纲级别）"

    ' 标题页：仅纵向滚动条
    frame.ScrollBars = fmScrollBarsVertical
    ' 计算滚动范围，确保滚动条可拖动到隐藏内容
    On Error Resume Next
    Dim maxRight As Single, maxBottom As Single
    Dim tc As Object
    maxRight = 0
    maxBottom = 0
    For Each tc In frame.Controls
        If (tc.Left + tc.Width) > maxRight Then maxRight = tc.Left + tc.Width
        If (tc.Top + tc.Height) > maxBottom Then maxBottom = tc.Top + tc.Height
    Next tc
    frame.ScrollHeight = CLng(maxBottom + 240)
    frame.ScrollTop = 0
    On Error GoTo 0
End Sub

Private Sub CreateTitleLevelGroup(ByRef parent As Object, ByVal level As Integer, Optional ByVal topOffset As Integer = 110)
    Dim grp As Object
    Dim topBase As Integer

    ' 每级标题设置区使用两行控件；默认 topOffset=110（用于“正文格式设置”）
    ' 注意：这里的间距用于给底部“自动识别/包含多级列表”开辟显示空间
    topBase = topOffset + (level - 1) * 70

    '（调试日志已移除）

    Set grp = parent.Controls.Add("Forms.Frame.1", "frameTitle" & level, True)
    With grp
        .caption = CStr(level) & "级标题设置"
        .left = 10
        .top = topBase
        .width = parent.Width - 20
        .height = 74
    End With

    AddLabel grp, "lblAlign" & level, "对齐:", 10, 18
    AddComboBox grp, "cmbAlign" & level, 45, 14, 70, 20

    AddLabel grp, "lblBefore" & level, "段前(行):", 135, 18
    AddTextBox grp, "txtBeforeLines" & level, 200, 14, 40, 20, "0"

    AddLabel grp, "lblAfter" & level, "段后(行):", 250, 18
    AddTextBox grp, "txtAfterLines" & level, 315, 14, 40, 20, "0"

    AddLabel grp, "lblFont" & level, "字体:", 370, 18
    AddComboBox grp, "cmbTitleFont" & level, 410, 14, 90, 20

    ' 第二行：字号 + 加粗
    AddLabel grp, "lblSize" & level, "字号:", 10, 48
    AddTextBox grp, "txtTitleSize" & level, 45, 44, 40, 20, "小四"
    AddCheckBox grp, "chkBold" & level, 95, 44, 60, 20, "加粗"

    AddLabel grp, "lblLineSpacing" & level, "行距:", 170, 48
    AddTextBox grp, "txtTitleSpacing" & level, 205, 44, 40, 20, "1.5"
End Sub

Private Sub CreateTablePage(ByRef mp As Object)
    Dim page As Object
    Set page = GetOrCreatePage(mp, "表格设置", -1)
    
    Dim frame As Object
    On Error Resume Next
    page.Controls.Remove "frameTable"
    On Error GoTo 0
    Set frame = page.Controls.Add("Forms.Frame.1", "frameTable", True)
    With frame
        .caption = "表格格式设置"
        .width = mp.Width - 40
        .height = 250
        .top = 10
        .left = 10
    End With
    
    ' 表格标题设置（二级标题） - 保持2级
    AddLabel frame, "lblTableTitle", "表格标题（二级标题）:", 10, 20  ' 改为二级标题
    AddLabel frame, "lblTableTitleFont", "字体:", 20, 50
    AddComboBox frame, "cmbTableTitleFont", 70, 45, 100, 20
    AddLabel frame, "lblTableTitleSize", "字号:", 200, 50
    AddTextBox frame, "txtTableTitleSize", 240, 45, 40, 20, "五号"
    AddCheckBox frame, "chkTableTitleBold", 300, 45, 60, 20, "加粗"
    AddLabel frame, "lblTableTitleSpacing", "行距:", 380, 50
    AddTextBox frame, "txtTableTitleSpacing", 420, 45, 40, 20, "1.0"
    
    AddLabel frame, "lblTableTitleLeftIndent", "左缩进(字符):", 20, 80
    AddTextBox frame, "txtTableTitleLeftIndent", 90, 75, 40, 20, "0"
    AddLabel frame, "lblTableTitleRightIndent", "右缩进(字符):", 150, 80
    AddTextBox frame, "txtTableTitleRightIndent", 220, 75, 40, 20, "0"
    
    AddLabel frame, "lblTableTitleSpecial", "特殊格式:", 20, 110
    AddComboBox frame, "cmbTableTitleSpecial", 80, 105, 90, 20
    AddLabel frame, "lblTableTitleSpecialIndent", "缩进值(字符):", 200, 110
    AddTextBox frame, "txtTableTitleSpecialIndent", 270, 105, 40, 20, "0"
    
    ' 表格内容设置
    AddLabel frame, "lblTableContent", "表格内容:", 10, 140
    AddLabel frame, "lblTableFont", "字体:", 20, 170
    AddComboBox frame, "cmbTableFont", 70, 165, 100, 20
    AddLabel frame, "lblTableSize", "字号:", 200, 170
    AddTextBox frame, "txtTableSize", 240, 165, 40, 20, "五号"
    
    ' 表格样式
    AddLabel frame, "lblTableStyle", "表格样式:", 10, 200
    AddLabel frame, "lblLineWidth", "线宽:", 20, 230
    AddTextBox frame, "txtLineWidth", 70, 225, 40, 20, "0.25"
    AddLabel frame, "lblRowHeight", "行高(cm):", 150, 230
    AddTextBox frame, "txtRowHeight", 210, 225, 40, 20, "0.6"
    
    ' 表格行距（可编辑，默认 1.0 倍）
    AddLabel frame, "lblTableSpacing", "行距:", 300, 230
    AddTextBox frame, "txtTableSpacing", 340, 225, 40, 20, "1.0"
    
    ' 根据窗口自动调整表格宽度
    AddCheckBox frame, "chkTableAutoFitWindow", 400, 225, 220, 20, "根据窗口自动调整表格宽度"

    ' 在“表格设置”内框启用滚动条（放到表格页）
    On Error Resume Next
    Dim maxRightT As Single, maxBottomT As Single
    Dim ct As Object
    maxRightT = 0
    maxBottomT = 0
    For Each ct In frame.Controls
        If (ct.Left + ct.Width) > maxRightT Then maxRightT = ct.Left + ct.Width
        If (ct.Top + ct.Height) > maxBottomT Then maxBottomT = ct.Top + ct.Height
    Next ct
    ' 表格设置：仅纵向滚动兜底
    frame.ScrollBars = fmScrollBarsVertical
    frame.ScrollHeight = CLng(maxBottomT + 200)
    frame.ScrollTop = 0
    On Error GoTo 0
End Sub

Private Sub CreateFigurePage(ByRef mp As Object)
    Dim page As Object
    Set page = GetOrCreatePage(mp, "图片标题设置", -1)
    
    Dim frame As Object
    On Error Resume Next
    page.Controls.Remove "frameFigure"
    On Error GoTo 0
    Set frame = page.Controls.Add("Forms.Frame.1", "frameFigure", True)
    With frame
        .caption = "图片标题格式设置"
        .width = mp.Width - 40
        .height = 170
        .top = 10
        .left = 10
    End With

    ' 图片标题：字体 / 字号 / 加粗 / 行距
    AddLabel frame, "lblFigureTitleFont", "字体:", 10, 30
    AddComboBox frame, "cmbFigureTitleFont", 50, 26, 100, 20
    
    AddLabel frame, "lblFigureTitleSize", "字号:", 170, 30
    AddTextBox frame, "txtFigureTitleSize", 210, 26, 40, 20, "五号"
    
    AddCheckBox frame, "chkFigureTitleBold", 270, 26, 60, 20, "加粗"
    
    AddLabel frame, "lblFigureTitleSpacing", "行距:", 350, 30
    AddTextBox frame, "txtFigureTitleSpacing", 390, 26, 40, 20, "1.0"
    
    AddLabel frame, "lblFigureTitleLeftIndent", "左缩进(字符):", 10, 65
    AddTextBox frame, "txtFigureTitleLeftIndent", 80, 60, 40, 20, "0"
    AddLabel frame, "lblFigureTitleRightIndent", "右缩进(字符):", 140, 65
    AddTextBox frame, "txtFigureTitleRightIndent", 210, 60, 40, 20, "0"
    
    AddLabel frame, "lblFigureTitleSpecial", "特殊格式:", 10, 95
    AddComboBox frame, "cmbFigureTitleSpecial", 70, 90, 90, 20
    AddLabel frame, "lblFigureTitleSpecialIndent", "缩进值(字符):", 190, 95
    AddTextBox frame, "txtFigureTitleSpecialIndent", 260, 90, 40, 20, "0"

    ' 在“图片标题设置”内框启用纵向滚动条
    On Error Resume Next
    Dim maxRightF As Single, maxBottomF As Single
    Dim cf As Object
    maxRightF = 0
    maxBottomF = 0
    For Each cf In frame.Controls
        If (cf.Left + cf.Width) > maxRightF Then maxRightF = cf.Left + cf.Width
        If (cf.Top + cf.Height) > maxBottomF Then maxBottomF = cf.Top + cf.Height
    Next cf
    frame.ScrollBars = fmScrollBarsVertical
    frame.ScrollHeight = CLng(IIf(maxBottomF + 80 > frame.Height + 20, maxBottomF + 80, frame.Height + 60))
    frame.ScrollTop = 0
    On Error GoTo 0
End Sub

' ========== 初始化下拉框、加载/保存设置 ==========

Private Sub InitializeComboBoxes()
    Dim mp As Object
    Dim frame As Object
    Dim fonts As Variant
    Dim i As Long
    
    fonts = Array("宋体", "仿宋", "黑体", "微软雅黑", "Times New Roman", "Arial")
    Set mp = GetMainMultiPage()
    
    ' 正文页
    Set frame = GetOrCreatePage(mp, "正文设置", 0).Controls("frameContent")
    Call FillFontCombo(frame.Controls("cmbBodyFont"), fonts, "宋体")

    With frame.Controls("cmbBodyAlign")
        .Clear
        .AddItem "左对齐"
        .AddItem "居中"
        .AddItem "右对齐"
        .AddItem "两端对齐"
        .ListIndex = 0
    End With
    
    With frame.Controls("cmbBodySpecial")
        .Clear
        .AddItem "无"
        .AddItem "首行缩进"
        .AddItem "悬挂缩进"
        .ListIndex = 0
    End With

    ' 标题页：1~4 级标题字体下拉框 + 对齐下拉框
    Dim lvl As Integer
    Dim grp As Object
    Set frame = GetOrCreatePage(mp, "标题设置", 1).Controls("frameTitleConfig")
    For lvl = 1 To 5
        Set grp = frame.Controls("frameTitle" & lvl)
        Call FillFontCombo(grp.Controls("cmbTitleFont" & lvl), fonts, "宋体")

        With grp.Controls("cmbAlign" & lvl)
            .Clear
            .AddItem "左对齐"
            .AddItem "居中"
            .AddItem "右对齐"
            .AddItem "两端对齐"
            .ListIndex = 0
        End With
    Next lvl

    ' 自动识别数字标题开关（checkbox 不需要填充下拉内容，这里仅确保存在不会报错）
    
    ' 表格页
    Set frame = GetOrCreatePage(mp, "表格设置", 2).Controls("frameTable")
    Call FillFontCombo(frame.Controls("cmbTableTitleFont"), fonts, "宋体")
    Call FillFontCombo(frame.Controls("cmbTableFont"), fonts, "宋体")
    With frame.Controls("cmbTableTitleSpecial")
        .Clear
        .AddItem "无"
        .AddItem "首行缩进"
        .AddItem "悬挂缩进"
        .ListIndex = 0
    End With
    
    ' 图片标题页
    Set frame = GetOrCreatePage(mp, "图片标题设置", 3).Controls("frameFigure")
    Call FillFontCombo(frame.Controls("cmbFigureTitleFont"), fonts, "宋体")
    With frame.Controls("cmbFigureTitleSpecial")
        .Clear
        .AddItem "无"
        .AddItem "首行缩进"
        .AddItem "悬挂缩进"
        .ListIndex = 0
    End With
End Sub

Private Sub FillFontCombo(cmb As Object, fonts As Variant, ByVal defaultFont As String)
    Dim i As Long
    cmb.Clear
    For i = LBound(fonts) To UBound(fonts)
        cmb.AddItem fonts(i)
    Next i
    cmb.Value = defaultFont
End Sub

Private Sub LoadSettingsToForm()
    Dim mp As Object
    Dim frame As Object
    
    Set mp = GetMainMultiPage()
    
    ' 正文
    Set frame = GetOrCreatePage(mp, "正文设置", 0).Controls("frameContent")
    With frame
        .Controls("cmbBodyFont").Value = gSettings.BodyFont
        .Controls("txtBodySize").Text = gSettings.BodySize
        If gSettings.BodySpacing <> 0 Then .Controls("txtBodySpacing").Text = CStr(gSettings.BodySpacing)
        If gSettings.BodyBeforeLines <> 0 Then .Controls("txtBodyBefore").Text = CStr(gSettings.BodyBeforeLines)
        If gSettings.BodyAfterLines <> 0 Then .Controls("txtBodyAfter").Text = CStr(gSettings.BodyAfterLines)
        If gSettings.BodyAlignment <> "" Then .Controls("cmbBodyAlign").Value = gSettings.BodyAlignment
        If gSettings.BodyLeftIndentCm <> 0 Then .Controls("txtBodyLeftIndent").Text = CStr(gSettings.BodyLeftIndentCm)
        If gSettings.BodyRightIndentCm <> 0 Then .Controls("txtBodyRightIndent").Text = CStr(gSettings.BodyRightIndentCm)
        If gSettings.BodySpecialIndent <> "" Then .Controls("cmbBodySpecial").Value = gSettings.BodySpecialIndent
        If gSettings.BodySpecialIndentCm <> 0 Then .Controls("txtBodySpecialIndent").Text = CStr(gSettings.BodySpecialIndentCm)
    End With
    ' 标题页：1~4 级标题设置回填
    Dim i As Integer
    Dim grp As Object
    Set frame = GetOrCreatePage(mp, "标题设置", 1).Controls("frameTitleConfig")
    For i = 1 To 5
        Set grp = frame.Controls("frameTitle" & i)
        grp.Controls("cmbAlign" & i).Value = gSettings.TitleAlignment(i)
        grp.Controls("txtBeforeLines" & i).Text = CStr(gSettings.TitleBeforeLines(i))
        grp.Controls("txtAfterLines" & i).Text = CStr(gSettings.TitleAfterLines(i))
        grp.Controls("cmbTitleFont" & i).Value = gSettings.TitleFont(i)
        grp.Controls("txtTitleSize" & i).Text = gSettings.TitleSize(i)
        grp.Controls("chkBold" & i).Value = gSettings.TitleBold(i)
        grp.Controls("txtTitleSpacing" & i).Text = CStr(gSettings.TitleLineSpacing(i))
    Next i

    ' 自动识别数字标题（开关）
    On Error Resume Next
    frame.Controls("chkAutoDetectNumericTitles").Value = gSettings.AutoDetectNumericTitles
    frame.Controls("chkAutoDetectIncludeList").Value = gSettings.AutoDetectIncludeListParagraphs
    On Error GoTo 0
    
    ' 表格
    Set frame = GetOrCreatePage(mp, "表格设置", 2).Controls("frameTable")
    With frame
        .Controls("cmbTableTitleFont").Value = gSettings.TableTitleFont
        .Controls("txtTableTitleSize").Text = gSettings.TableTitleSize
        .Controls("chkTableTitleBold").Value = gSettings.TableTitleBold
        If gSettings.TableTitleSpacing <> 0 Then .Controls("txtTableTitleSpacing").Text = CStr(gSettings.TableTitleSpacing)
        If gSettings.TableTitleLeftIndentCm <> 0 Then .Controls("txtTableTitleLeftIndent").Text = CStr(gSettings.TableTitleLeftIndentCm)
        If gSettings.TableTitleRightIndentCm <> 0 Then .Controls("txtTableTitleRightIndent").Text = CStr(gSettings.TableTitleRightIndentCm)
        If gSettings.TableTitleSpecialIndent <> "" Then .Controls("cmbTableTitleSpecial").Value = gSettings.TableTitleSpecialIndent
        If gSettings.TableTitleSpecialIndentCm <> 0 Then .Controls("txtTableTitleSpecialIndent").Text = CStr(gSettings.TableTitleSpecialIndentCm)
        
        .Controls("cmbTableFont").Value = gSettings.TableFont
        .Controls("txtTableSize").Text = gSettings.TableSize
        If gSettings.TableLineWidth <> 0 Then .Controls("txtLineWidth").Text = CStr(gSettings.TableLineWidth)
        If gSettings.TableRowHeight <> 0 Then .Controls("txtRowHeight").Text = CStr(gSettings.TableRowHeight)
        If gSettings.TableSpacing <> 0 Then .Controls("txtTableSpacing").Text = CStr(gSettings.TableSpacing)
        
        ' 根据窗口自动调整表格宽度
        On Error Resume Next
        .Controls("chkTableAutoFitWindow").Value = gSettings.TableAutoFitWindow
        On Error GoTo 0
    End With
    
    ' 图片标题
    Set frame = GetOrCreatePage(mp, "图片标题设置", 3).Controls("frameFigure")
    With frame
        .Controls("cmbFigureTitleFont").Value = gSettings.FigureTitleFont
        .Controls("txtFigureTitleSize").Text = gSettings.FigureTitleSize
        .Controls("chkFigureTitleBold").Value = gSettings.FigureTitleBold
        If gSettings.FigureTitleSpacing <> 0 Then .Controls("txtFigureTitleSpacing").Text = CStr(gSettings.FigureTitleSpacing)
        If gSettings.FigureTitleLeftIndentCm <> 0 Then .Controls("txtFigureTitleLeftIndent").Text = CStr(gSettings.FigureTitleLeftIndentCm)
        If gSettings.FigureTitleRightIndentCm <> 0 Then .Controls("txtFigureTitleRightIndent").Text = CStr(gSettings.FigureTitleRightIndentCm)
        If gSettings.FigureTitleSpecialIndent <> "" Then .Controls("cmbFigureTitleSpecial").Value = gSettings.FigureTitleSpecialIndent
        If gSettings.FigureTitleSpecialIndentCm <> 0 Then .Controls("txtFigureTitleSpecialIndent").Text = CStr(gSettings.FigureTitleSpecialIndentCm)
    End With

    ' 记忆设置（在窗体顶部）
    On Error Resume Next
    Me.Controls(REMEMBER_TOGGLE_NAME).Value = gSettings.RememberSettings
    On Error GoTo 0
End Sub

Private Sub btnApply_Click()
    ApplyFormatting
End Sub

Private Sub btnOK_Click()
    ApplyFormatting
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

' 辅助函数：查找或创建按钮
'Private Function FindOrCreateButton(btnName As String) As Object
   ' Dim btn As Object
    'Dim i As Integer
    'Dim found As Boolean
    
    'found = False
    'For i = 0 To UserForm1.Controls.Count - 1
     '   If UserForm1.Controls(i).name = btnName Then
      '      Set btn = UserForm1.Controls(i)
       '     found = True
        '    Exit For
        'End If
    'Next i
    
    'If Not found Then
     '   Set btn = UserForm1.Controls.Add("Forms.CommandButton.1", btnName, True)
    'End If
    
    'Set FindOrCreateButton = btn
'End Function

' 辅助函数：添加标签
Private Sub AddLabel(ByRef container As Object, name As String, caption As String, left As Integer, top As Integer)
    Dim lbl As Object
    Set lbl = container.Controls.Add("Forms.Label.1", name, True)
    With lbl
        .caption = caption
        .left = left
        .top = top
        .width = 100

        
        .height = 18
    End With
End Sub

' 辅助函数：添加文本框
Private Sub AddTextBox(ByRef container As Object, name As String, left As Integer, top As Integer, width As Integer, height As Integer, Optional defaultValue As String = "")
    Dim txt As Object
    Set txt = container.Controls.Add("Forms.TextBox.1", name, True)
    With txt
        .left = left
        .top = top
        .width = width
        .height = height
        If defaultValue <> "" Then .Text = defaultValue
    End With
End Sub

' 辅助函数：添加组合框
Private Sub AddComboBox(ByRef container As Object, name As String, left As Integer, top As Integer, width As Integer, height As Integer)
    Dim cmb As Object
    Set cmb = container.Controls.Add("Forms.ComboBox.1", name, True)
    With cmb
        .left = left
        .top = top
        .width = width
        .height = height
        .Style = 2 ' 下拉列表
    End With
End Sub

' 辅助函数：添加复选框
Private Sub AddCheckBox(ByRef container As Object, name As String, left As Integer, top As Integer, width As Integer, height As Integer, caption As String)
    Dim chk As Object
    Set chk = container.Controls.Add("Forms.CheckBox.1", name, True)
    With chk
        .caption = caption
        .left = left
        .top = top
        .width = width
        .height = height
    End With
End Sub


