Option Explicit

' 全局常量（注册表存储用，仅本地，不上传任何数据）
Public Const APP_NAME As String = "WordFormatPlugin"
Public Const SECTION_NAME As String = "FormatSettings"

' =======================
' 全局格式设置结构与变量
' =======================

Public Type FormatSettings
    ' 目录（1~3 级数字标题）
    Level1Font As String
    Level1Size As String
    Level1Spacing As Single
    Level1Bold As Boolean

    Level2Font As String
    Level2Size As String
    Level2IndentChars As Integer   ' 以“字符数”计的左缩进
    Level2Spacing As Single
    Level2Bold As Boolean

    Level3Font As String
    Level3Size As String
    Level3IndentChars As Integer
    Level3Spacing As Single
    Level3Bold As Boolean

    ' 正文默认
    BodyFont As String
    BodySize As String
    BodySpacing As Single          ' 行间距倍数，如 1.5
    BodyBeforeLines As Single      ' 段前（行）
    BodyAfterLines As Single       ' 段后（行）
    BodyAlignment As String        ' 对齐方式：左对齐/居中/右对齐/两端对齐
    BodyLeftIndentCm As Single     ' 左缩进（厘米）
    BodyRightIndentCm As Single    ' 右缩进（厘米）
    BodySpecialIndent As String    ' 特殊格式：无/首行缩进/悬挂缩进
    BodySpecialIndentCm As Single  ' 特殊格式缩进值（厘米）

    ' 正文中 1~5 级标题属性（支持 1~5）
    TitleAlignment(1 To 5) As String       ' 左对齐/居中/右对齐/两端对齐
    TitleBeforeLines(1 To 5) As Single     ' 段前（行）
    TitleAfterLines(1 To 5) As Single      ' 段后（行）
    TitleFont(1 To 5) As String            ' 中文字体（NameFarEast）
    TitleSize(1 To 5) As String            ' 字号（如 小四/12）
    TitleBold(1 To 5) As Boolean
    TitleLineSpacing(1 To 5) As Single     ' 行间距倍数（如 1.5）

    ' 自动识别“1.1/1.1.1...”并设置大纲级别
    AutoDetectNumericTitles As Boolean
    AutoDetectIncludeListParagraphs As Boolean

    ' 表格标题与内容
    TableTitleFont As String
    TableTitleSize As String
    TableTitleBold As Boolean
    TableTitleSpacing As Single     ' 表格题注（标题）行距倍数，如 1.0/1.5
    TableTitleLeftIndentCm As Single
    TableTitleRightIndentCm As Single
    TableTitleSpecialIndent As String
    TableTitleSpecialIndentCm As Single

    TableFont As String
    TableSize As String
    TableLineWidth As Single       ' 磅
    TableRowHeight As Single       ' cm
    TableSpacing As Single         ' 行距倍数，如 1.0
    
    ' 表格宽度：是否根据窗口自动调整
    TableAutoFitWindow As Boolean
    
    ' 图片标题（图1.5-1 …）
    FigureTitleFont As String
    FigureTitleSize As String
    FigureTitleBold As Boolean
    FigureTitleSpacing As Single    ' 图片标题行距倍数，如 1.0/1.5
    FigureTitleLeftIndentCm As Single
    FigureTitleRightIndentCm As Single
    FigureTitleSpecialIndent As String
    FigureTitleSpecialIndentCm As Single

    ' 是否记忆设置
    RememberSettings As Boolean
End Type

' 全局当前设置
Public gSettings As FormatSettings
