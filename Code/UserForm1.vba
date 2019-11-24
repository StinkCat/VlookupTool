Function StrtoW(ByVal SrtText As String) As String '字符转列号再转列字母
    Dim ColNum As Long, Index As Long
    Index = Application.WorksheetFunction.Search(".", SrtText)
    ColNum = Left(SrtText, Index)
    StrtoW = Replace(Cells(1, ColNum).Address(False, False), "1", "")
End Function
Private Sub Col1_Change()
    If Col1 <> "" Then
        Set SourceSheet = Excel.Workbooks(ComboBox1.Value).Sheets(ComboBox2.Value)
        ColStr = StrtoW(Col1.Value)
        RowsNum = Application.CountA(SourceSheet.Range(ColStr & ":" & ColStr))
        Tip1.Value = ColStr & "列 " & Format(RowsNum / 10000, "0.00") & " 万"
    End If
End Sub
Private Sub Col2_Change()
    If Col2 <> "" Then
        Set SourceSheet = Excel.Workbooks(ComboBox1.Value).Sheets(ComboBox2.Value)
        ColStr = StrtoW(Col2.Value)
        RowsNum = Application.CountA(SourceSheet.Range(ColStr & ":" & ColStr))
        Tip2.Value = ColStr & "列 " & Format(RowsNum / 10000, "0.00") & " 万"
    End If
End Sub
Private Sub Col3_Change()
    If Col3 <> "" Then
        Set TargetSheet = Excel.Workbooks(ComboBox3.Value).Sheets(ComboBox4.Value)
        ColStr = StrtoW(Col3.Value)
        RowsNum = Application.CountA(TargetSheet.Range(ColStr & ":" & ColStr))
        Tip3.Value = ColStr & "列 " & Format(RowsNum / 10000, "0.00") & " 万"
    End If
End Sub
Private Sub Col4_Change()
    If Col4 <> "" Then
        Set TargetSheet = Excel.Workbooks(ComboBox3.Value).Sheets(ComboBox4.Value)
        ColStr = StrtoW(Col4.Value)
        RowsNum = Application.CountA(TargetSheet.Range(ColStr & ":" & ColStr))
        Tip4.Value = ColStr & "列 " & Format(RowsNum / 10000, "0.00") & " 万"
    End If
End Sub

Private Sub ComboBox1_Change() '工作簿改动时添加工作表名称
    If ComboBox1.Value <> "" Then
        Set SourceBook = Excel.Workbooks(ComboBox1.Value)
        Me.Controls("ComboBox2").Clear
        For Index = 1 To SourceBook.Sheets.Count
            Me.Controls("ComboBox2").AddItem SourceBook.Sheets(Index).Name
        Next
        ComboBox2.ListIndex = 0
    End If
End Sub

Private Sub ComboBox2_Change()
    If ComboBox2.Value <> "" Then
        Me.Controls("Col1").Clear
        Me.Controls("Col2").Clear
        Set SourceSheet = Excel.Workbooks(ComboBox1.Value).Sheets(ComboBox2.Value)
        ColNum = SourceSheet.UsedRange.Columns.Count
        For n = 1 To ColNum
            If SourceSheet.Cells(1, n) <> "" Then
                Me.Controls("Col1").AddItem n & "." & SourceSheet.Cells(1, n)
                Me.Controls("Col2").AddItem n & "." & SourceSheet.Cells(1, n)
            Else
                Me.Controls("Col1").AddItem n & ".NULL"
                Me.Controls("Col2").AddItem n & ".NULL"
            End If
        Next n
    End If
End Sub
Private Sub ComboBox4_Change()
    If ComboBox4.Value <> "" Then
        Me.Controls("Col3").Clear
        Me.Controls("Col4").Clear
        Set TargetSheet = Excel.Workbooks(ComboBox3.Value).Sheets(ComboBox4.Value)
        ColNum = TargetSheet.UsedRange.Columns.Count
        For n = 1 To ColNum
            If TargetSheet.Cells(1, n) <> "" Then
            
                Me.Controls("Col3").AddItem n & "." & TargetSheet.Cells(1, n)
                Me.Controls("Col4").AddItem n & "." & TargetSheet.Cells(1, n)
            Else
                Me.Controls("Col3").AddItem n & ".NULL"
                Me.Controls("Col4").AddItem n & ".NULL"
            End If
        Next n
        Me.Controls("Col4").AddItem n & ".空白尾列"
    End If
End Sub

Private Sub ComboBox3_Change() '工作簿改动时添加工作表名称
    If ComboBox3.Value <> "" Then
        Set SourceBook = Excel.Workbooks(ComboBox3.Value)
        Me.Controls("ComboBox4").Clear
        For Index = 1 To SourceBook.Sheets.Count
            Me.Controls("ComboBox4").AddItem SourceBook.Sheets(Index).Name
        Next
        ComboBox4.ListIndex = 0
    End If
End Sub

Private Sub CommandButton2_Click()
    UserForm3.Show
End Sub

Private Sub Image1_Click()
    Me.Controls("ComboBox1").Clear
    BooksCount = 0
    For Index = 1 To Excel.Workbooks.Count
        If Excel.Workbooks(Index).Name <> ThisWorkbook.Name Then
            Me.Controls("ComboBox1").AddItem Excel.Workbooks(Index).Name
'            Me.Controls("ComboBox3").AddItem Excel.Workbooks(Index).Name
            BooksCount = BooksCount + 1
        End If
    Next
    If BooksCount > 0 Then
        ComboBox1.ListIndex = 0
'        ComboBox3.ListIndex = 0
    End If
End Sub
Private Sub Image2_Click()
    Me.Controls("ComboBox3").Clear
    BooksCount = 0
    For Index = 1 To Excel.Workbooks.Count
        If Excel.Workbooks(Index).Name <> ThisWorkbook.Name Then
'            Me.Controls("ComboBox1").AddItem Excel.Workbooks(Index).Name
            Me.Controls("ComboBox3").AddItem Excel.Workbooks(Index).Name
            BooksCount = BooksCount + 1
        End If
    Next
    If BooksCount > 0 Then
'        ComboBox1.ListIndex = 0
        ComboBox3.ListIndex = 0
    End If
End Sub

Private Sub UserForm_Initialize() '窗体默认启动时添加工作簿名称
    BooksCount = 0
    For Index = 1 To Excel.Workbooks.Count
        If Excel.Workbooks(Index).Name <> ThisWorkbook.Name Then
            Me.Controls("ComboBox1").AddItem Excel.Workbooks(Index).Name
            Me.Controls("ComboBox3").AddItem Excel.Workbooks(Index).Name
            BooksCount = BooksCount + 1
        End If
    Next
    If BooksCount > 0 Then
        ComboBox1.ListIndex = 0
        ComboBox3.ListIndex = 0
    End If
    CheckBox1.Value = True
    TextBox5.Value = 2
    TextBox6.Value = "#N/A"
    TextBox7.Value = "&"
End Sub
'主函数（按键开始）入口
Private Sub CommandButton1_Click()
 
    Dim SourceKeyCol, SourceValueCol As String, ThisKeyCol As String, ThisValueCol As String, WriteNum As Integer, This_rows As Long, Source_rows As Long
    t0 = Timer
    SourceKeyCol = StrtoW(Col1.Value)
    SourceValueCol = StrtoW(Col2.Value)
    ThisKeyCol = StrtoW(Col3.Value)
    ThisValueCol = StrtoW(Col4.Value)
    WriteNum = TextBox5.Value
    Set SourceSheet = Excel.Workbooks(ComboBox1.Value).Sheets(ComboBox2.Value)
    Set ThisSheet = Excel.Workbooks(ComboBox3.Value).Sheets(ComboBox4.Value)
    Set MapDict = CreateObject("Scripting.Dictionary")
    CheckNum = Application.CountA(ThisSheet.Range(ThisValueCol & ":" & ThisValueCol))  '统计有效行数
    If CheckNum >= WriteNum Then
        UserForm2.Show
        If Is_Exit = True Then
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
            If Not MapDict.Exists(CStr(S_Key(n, 1))) Then
                MapDict.Add CStr(S_Key(n, 1)), S_Value(n, 1)
            End If
        Next n
    Else
        For n = 1 To Source_rows
            If Not MapDict.Exists(CStr(S_Key(n, 1))) Then
                MapDict.Add CStr(S_Key(n, 1)), S_Value(n, 1)
            Else
                MapDict(CStr(S_Key(n, 1))) = MapDict(CStr(S_Key(n, 1))) & TextBox7.Value & S_Value(n, 1)
            End If
        Next n
    End If
    Dim Result()
    ReDim Result(WriteNum To This_rows, 1 To 1)
    If CheckBox1.Value = False Then '索引结果写入到数组
        For n = WriteNum To This_rows
            Result(n, 1) = MapDict(CStr(T_Key(n, 1)))
        Next
    Else
        For n = WriteNum To This_rows
            If Not MapDict.Exists(CStr(T_Key(n, 1))) Then
                Result(n, 1) = TextBox6.Value
            Else
                Result(n, 1) = MapDict(CStr(T_Key(n, 1)))
            End If
        Next
    End If
    ThisSheet.Range(ThisValueCol & WriteNum).Resize(This_rows - WriteNum + 1, 1) = Result
    MsgBox "共查找 " & This_rows - WriteNum + 1 & " 条记录，耗时" & Format(Timer - t0, "0.00") & "秒。"
    Set MapDict = Nothing
End Sub

