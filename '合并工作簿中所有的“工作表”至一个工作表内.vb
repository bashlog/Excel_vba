'合并工作簿中所有的“工作表”至一个工作表内
Sub WorkSheetsMerge()
    Application.ScreenUpdating = False
    Cells.ClearContents '清空当前表格数据
    Cells.Clear '清空当前表格格式
    Range("A3") = "来源工作表名称"
    Range("B2") = " " '占位
    Tempelete = "WorkSheets Merge Tool"
    nTitleRow = Val(InputBox("请输入标题的行数，默认标题行数为1" & vbCrLf & "如无标题行则行数填写 0", Tempelete, 1))
    If nTitleRow < 0 Then MsgBox "标题行数不能为负数。", 64, "警告": Exit Sub
    For i = 1 To Sheets.Count
        If Sheets(i).Name <> ActiveSheet.Name Then
            rowused = Cells(Rows.Count, 2).End(xlUp).Row + 1
            nShtCount = nShtCount + 1 '汇总工作表的数量
            nStartRow = IIf(nTitleRow = 1, 1, 0) '判断遍历数据源是否应该扣掉标题行
            lastrow = rowused
            If nShtCount = 1 Then
                Sheets(i).UsedRange.Offset(0).Copy Cells(rowused, 2)
                rowused = Cells(Rows.Count, 2).End(xlUp).Row
                ActiveSheet.Range(Cells(lastrow + 1, 1), Cells(rowused, 1)) = Sheets(i).Name
            Else
                Sheets(i).UsedRange.Offset(nStartRow).Copy Cells(rowused, 2)
                rowused = Cells(Rows.Count, 2).End(xlUp).Row
                ActiveSheet.Range(Cells(lastrow, 1), Cells(rowused, 1)) = Sheets(i).Name
            End If
        End If
    Next
    Cells.Select
    Cells.EntireColumn.AutoFit
    Application.ScreenUpdating = True
    Range("A3").Select
    MsgBox "当前工作簿下的全部工作表已经合并完毕！" & vbCrLf & "一共汇总完成 " & nShtCount & "个工作表！", vbInformation, Tempelete
End Sub