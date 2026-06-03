' Excel 修改脚本 - 由 Python _gen_one 调用
' 参数: template_path output_path new_start new_end wh_number name_value
' 逻辑: 遍历所有工作表，每行第1个>=50000的数值→新起点，第2个→新终点

Dim args, tpl, out, ns, ne, wh, nv
Set args = WScript.Arguments
tpl = args(0) : out = args(1)
ns = CDbl(args(2)) : ne = CDbl(args(3))
wh = args(4) : nv = args(5)

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile tpl, out, True

Dim excel, wb, sh, rr, cc, cv, cnt
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Open(out)

' 改封面
wb.Worksheets("封面").Cells(12, 2).Value = nv
wb.Worksheets("封面").Cells(15, 8).Value = wh

' 遍历工作表
For Each sh In wb.Worksheets
    If sh.Name <> "封面" Then
        For rr = 1 To 500
            cnt = 0
            For cc = 1 To 100
                On Error Resume Next
                cv = sh.Cells(rr, cc).Value
                If Not IsEmpty(cv) And IsNumeric(cv) Then
                    If CDbl(cv) >= 50000 Then
                        cnt = cnt + 1
                        If cnt = 1 Then sh.Cells(rr, cc).Value = ns
                        If cnt = 2 Then sh.Cells(rr, cc).Value = ne
                    End If
                End If
                On Error GoTo 0
            Next
        Next
    End If
Next

wb.Save
wb.Close
excel.Quit
