Option Explicit

' 入口：显示设置窗口（带配置界面）
Public Sub ShowFormatSettings()
    On Error GoTo EH
50: ' 启动前自检：若数组仍是 1~4，直接给出明确提示并停止打开窗体
    If Not ValidateTitleLevel5Support() Then
        MsgBox "检测到当前 VBA 工程仍在使用旧版 typemodule（标题数组仅支持 1~4 级）。" & vbCrLf & _
               "请在 VBE 工程中的 typemodule 把以下数组改为 (1 To 5)：" & vbCrLf & _
               "TitleAlignment/TitleBeforeLines/TitleAfterLines/TitleFont/TitleSize/TitleBold/TitleLineSpacing", vbExclamation
        Exit Sub
    End If
100: ' Erl 用于定位：注册表加载
    LoadSettingsFromRegistry
    
200: ' Erl 用于定位：显示窗体
    UserForm1.Show vbModeless
    Exit Sub
EH:
    MsgBox "ShowFormatSettings 出错：" & vbCrLf & _
           "Err=" & Err.Number & " Erl=" & Erl & " " & Err.Description & vbCrLf & _
           "定位：如果 Erl=100 为 `LoadSettingsFromRegistry`，Erl=200 为 `UserForm1.Show`。", vbExclamation
    Err.Clear
End Sub

 ' 快速应用：不弹窗，直接按任务书要求刷新当前文档
Private Sub QuickApplyFormat()
    ' 先加载记忆的设置
    LoadSettingsFromRegistry
    
    ' 直接对当前文档应用格式（保留自动编号）
    ApplySettingsToDocument
End Sub
