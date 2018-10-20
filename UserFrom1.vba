''''主窗体代码

Private Sub ComboBox1_Change() '工作簿改动时添加工作表名称
    Set SourceBook = Excel.Workbooks(ComboBox1.Value)
    Me.Controls("ComboBox2").Clear
    For Index = 1 To SourceBook.Sheets.Count
        Me.Controls("ComboBox2").AddItem SourceBook.Sheets(Index).Name
    Next
End Sub


Private Sub ComboBox3_Change() '工作簿改动时添加工作表名称
    Set SourceBook = Excel.Workbooks(ComboBox3.Value)
    Me.Controls("ComboBox4").Clear
    For Index = 1 To SourceBook.Sheets.Count
        Me.Controls("ComboBox4").AddItem SourceBook.Sheets(Index).Name
    Next
End Sub

Private Sub CommandButton2_Click()
    UserForm3.Show
End Sub

Private Sub TextBox1_Change()
    If IsNumeric(TextBox1.Value) Then
        MsgBox "输入值为数字，请输入所在列的字母，例如A列请输入“A”"
    Else
        If ComboBox1.Value <> "" And ComboBox2.Value <> "" And TextBox1.Value <> "" Then
            Tip1.Value = Excel.Workbooks(ComboBox1.Value).Sheets(ComboBox2.Value).Range(TextBox1.Value & 1)
        End If
    End If
End Sub
Private Sub TextBox2_Change()
    If IsNumeric(TextBox2.Value) Then
        MsgBox "输入值为数字，请输入所在列的字母，例如A列请输入“A”"
    Else
        If ComboBox1.Value <> "" And ComboBox2.Value <> "" And TextBox2.Value <> "" Then
            Tip2.Value = Excel.Workbooks(ComboBox1.Value).Sheets(ComboBox2.Value).Range(TextBox2.Value & 1)
        End If
    End If
End Sub
Private Sub TextBox3_Change()
    If IsNumeric(TextBox3.Value) Then
        MsgBox "输入值为数字，请输入所在列的字母，例如A列请输入“A”"
    Else
        If ComboBox3.Value <> "" And ComboBox4.Value <> "" And TextBox3.Value <> "" Then
            Tip3.Value = Excel.Workbooks(ComboBox3.Value).Sheets(ComboBox4.Value).Range(TextBox3.Value & 1)
        End If
    End If
End Sub
Private Sub TextBox4_Change()
    If IsNumeric(TextBox4.Value) Then
        MsgBox "输入值为数字，请输入所在列的字母，例如A列请输入“A”"
    Else
        If ComboBox3.Value <> "" And ComboBox4.Value <> "" And TextBox4.Value <> "" Then
            Tip4.Value = Excel.Workbooks(ComboBox3.Value).Sheets(ComboBox4.Value).Range(TextBox4.Value & 1)
        End If
    End If
End Sub


Private Sub UserForm_Initialize() '窗体默认启动时添加工作簿名称
    For Index = 1 To Excel.Workbooks.Count
        Me.Controls("ComboBox1").AddItem Excel.Workbooks(Index).Name
        Me.Controls("ComboBox3").AddItem Excel.Workbooks(Index).Name
        TextBox5.Value = 2
        TextBox6.Value = "#N/A"
        TextBox7.Value = "&"
    Next
End Sub
'主函数（按键）入口
Private Sub CommandButton1_Click()
 
    Dim SourceKeyCol, SourceValueCol As String, ThisKeyCol As String, ThisValueCol As String, WriteNum As Integer
    t0 = Timer
    SourceKeyCol = TextBox1.Value
    SourceValueCol = TextBox2.Value
    ThisKeyCol = TextBox3.Value
    ThisValueCol = TextBox4.Value
    WriteNum = TextBox5.Value
    Set SourceSheet = Excel.Workbooks(ComboBox1.Value).Sheets(ComboBox2.Value)
    Set ThisSheet = Excel.Workbooks(ComboBox3.Value).Sheets(ComboBox4.Value)
    Set MapDict = CreateObject("Scripting.Dictionary")
    CheckNum = Application.CountA(ThisSheet.Range(TextBox4.Value & ":" & TextBox4.Value))  '统计有效行数
    If CheckNum >= WriteNum Then
        UserForm2.Show
        While King = -1
        Wend
        If King = 0 Then
            Exit Sub
        End If
    End If
    t0 = Timer '统计时间
    Source_rows = SourceSheet.UsedRange.Rows.Count '统计整个表最大行数
    This_rows = ThisSheet.UsedRange.Rows.Count
    S_Key = SourceSheet.Range(SourceKeyCol & "1").Resize(Source_rows)
    S_Value = SourceSheet.Range(SourceValueCol & "1").Resize(Source_rows)
    T_Key = ThisSheet.Range(ThisKeyCol & "1").Resize(This_rows)
    If CheckBox2.Value = False Then '将源数据写入字典
        For n = 1 To Source_rows
            If Not MapDict.Exists(S_Key(n, 1)) Then
                MapDict.Add S_Key(n, 1), S_Value(n, 1)
            End If
        Next n
    Else
        For n = 1 To Source_rows
            If Not MapDict.Exists(S_Key(n, 1)) Then
                MapDict.Add S_Key(n, 1), S_Value(n, 1)
            Else
                MapDict(S_Key(n, 1)) = MapDict(S_Key(n, 1)) & TextBox7.Value & S_Value(n, 1)
            End If
        Next n
    End If
    Dim Result()
    ReDim Result(WriteNum To This_rows, 1 To 1)
    If CheckBox1.Value = False Then '索引结果写入到数组
        For n = WriteNum To This_rows
            Result(n, 1) = MapDict(T_Key(n, 1))
        Next
    Else
        For n = WriteNum To This_rows
            If Not MapDict.Exists(T_Key(n, 1)) Then
                Result(n, 1) = TextBox6.Value
            Else
                Result(n, 1) = MapDict(T_Key(n, 1))
            End If
        Next
    End If
    ThisSheet.Range(ThisValueCol & WriteNum).Resize(This_rows - WriteNum + 1, 1) = Result
    MsgBox "共处理 " & This_rows - WriteNum + 1 & " 条记录，耗时" & Format(Timer - t0, "0.00") & "秒。"
    Set MapDict = Nothing
End Sub

