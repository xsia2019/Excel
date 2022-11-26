### Excel 文件自动记录最后保存时间

在Sheet1的A1单元格中记录文档最后保存时间

```vba
Sub latestSaveTime()
Sheets("Sheet1").Range("A1") = Format(Now, "yyyy.mm.dd hh:mm")
End Sub
```

自动更新最后保存时间

打开开发工具，双击VBA窗口[Microsoft Excel对象]下的[ThisWorkBook]，
代码编辑窗口选择[Workbook]的[AfterSave]，
把代码放在Private Sub Workbook_AfterSave(ByVal Success As Boolean)
下、End Sub之前。