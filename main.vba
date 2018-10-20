'模块1启动窗体显示

Option Explicit
Public King As Integer
Sub 启动VlookupTool()
    UserForm1.Show
    King = -1 '定义全局变量King用于标记填充列是否存在较多数据处理标示，默认为-1，为1时覆盖单元格，为0时退出程序
End Sub
