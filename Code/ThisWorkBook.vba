'ThisWorkBook代码，主要实现快捷键及欢迎界面
Private Sub Workbook_Open()
    Application.OnKey "%{F1}", "启动VlookupTool"
     UserForm4.Show
End Sub
