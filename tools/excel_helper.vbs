' Excel 修改辅助脚本 — 由 Python 调用
' 参数: template_path output_path name_value wh_number

Dim args, template_path, output_path, name_value, wh_number
Set args = WScript.Arguments
template_path = args(0)
output_path = args(1)
name_value = args(2)
wh_number = args(3)

Dim excel, wb, ws, sh, used, r, c, v, cv, old_ch, new_ch, rr, cc
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

Set wb = excel.Workbooks.Open(template_path)

' 改封面
Set ws = wb.Worksheets("封面")
ws.Cells(12, 2).Value = name_value
ws.Cells(15, 8).Value = wh_number

' 提取旧桩号字符串
old_b12 = ws.Cells(12, 2).Value
Set re = New RegExp
re.Pattern = "[ZY]?K\d+\+[\d.]+[～][ZY]?K\d+\+[\d.]+"
re.IgnoreCase = True
Set matches = re.Execute(old_b12)
If matches.Count > 0 Then
    old_ch = matches(0).Value
    ' 解析新旧数值
    Set re2 = New RegExp
    re2.Pattern = "(\d+)\+([\d.]+)"
    Set parts = re2.Execute(old_ch)
    old_s = CLng(parts(0).SubMatches(0)) * 1000 + CDbl(parts(0).SubMatches(1))
    old_e = CLng(parts(1).SubMatches(0)) * 1000 + CDbl(parts(1).SubMatches(1))

    ' 新桩号字符串
    Set re3 = New RegExp
    re3.Pattern = "\d+\+[\d.]+"
    Set new_parts = re3.Execute(name_value)
    new_s = CLng(new_parts(0).SubMatches(0)) * 1000 + CDbl(new_parts(0).SubMatches(1))
    new_e = CLng(new_parts(1).SubMatches(0)) * 1000 + CDbl(new_parts(1).SubMatches(1))

    ' 遍历所有 sheet（除封面）
    For Each sh In wb.Worksheets
        If sh.Name <> "封面" Then
            ' 找"起点桩号""终点桩号"列
            start_col = 0
            end_col = 0
            For rr = 1 To 100
                For cc = 1 To 50
                    On Error Resume Next
                    cv = sh.Cells(rr, cc).Value
                    On Error GoTo 0
                    If Not IsEmpty(cv) And VarType(cv) = 8 Then ' string
                        If InStr(cv, "起点桩号") > 0 Then start_col = cc
                        If InStr(cv, "终点桩号") > 0 Then end_col = cc
                    End If
                Next
            Next

            ' 替换桩号列所有数值（跳过表头行）
            For rr = 1 To 500
                On Error Resume Next
                If start_col > 0 Then
                    cv = sh.Cells(rr, start_col).Value
                    If Not IsEmpty(cv) And VarType(cv) >= 5 And VarType(cv) <= 6 Then
                        sh.Cells(rr, start_col).Value = new_s
                    End If
                End If
                If end_col > 0 Then
                    cv = sh.Cells(rr, end_col).Value
                    If Not IsEmpty(cv) And VarType(cv) >= 5 And VarType(cv) <= 6 Then
                        sh.Cells(rr, end_col).Value = new_e
                    End If
                End If
                ' 替换字符串中的旧桩号
                cv = sh.Cells(rr, cc).Value
                If Not IsEmpty(cv) And VarType(cv) = 8 Then
                    If InStr(cv, old_ch) > 0 Then
                        sh.Cells(rr, cc).Value = Replace(cv, old_ch, name_value)
                    End If
                End If
                On Error GoTo 0
            Next
        End If
    Next
End If

wb.SaveAs output_path
wb.Close
excel.Quit
