'删除某一列中空白单元格所在的“行”
sub my()
dim i as long
for i = 1 to [a65536].end(xlup).row
    if cells(i, 2) = "小计" then                              '这里的2就是你的列数，可以自己替换
    rows(i & ":" & i).delete shift:=xlup
end if
next
end sub