'数据库功能扩展宏
'适用于Excel及WPS环境下的原石计划MOD开发

'作者：小怡片刻的曲奇脆丝
'邮箱：2813443253@qq.com

'工作表批量保存为CSV文件
Sub SaveAsCSV()
  Dim pth As String
  Dim wb As WorkBook
  Dim sht As WorkSheet

  path = DbConfig("path")
  wb = ThisWorkBook

  For Each sht In wb.Worksheets
    If Not IsDbMeta(sht.Name) Then
      sheet.SaveAs path & sht.Name & ".csv", xlCSV
    End If
  Next sht
  wb.Activate
End Sub

Function DbConfig(data)
