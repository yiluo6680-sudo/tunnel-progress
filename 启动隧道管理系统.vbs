' 隧道工程管理系统 — 静默启动（无命令窗口）
' 双击此文件即可启动，浏览器打开后可最小化到托盘

Dim shell
Set shell = CreateObject("WScript.Shell")

' 切换到项目目录
shell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' 静默运行（0=隐藏窗口）
shell.Run "streamlit run app.py", 0, False

' 等几秒让服务启动后打开浏览器
WScript.Sleep 5000
shell.Run "http://localhost:8501"
