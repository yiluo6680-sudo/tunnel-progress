' 隧道工程管理系统 — 双击启动，无命令窗口，自动打开浏览器
Dim shell, fs
Set shell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")

' 切换到脚本所在目录
shell.CurrentDirectory = fs.GetParentFolderName(WScript.ScriptFullName)

' 启动 Streamlit（隐藏窗口）
shell.Run "python -m streamlit run app.py", 0, False

' 等几秒让服务启动
WScript.Sleep 6000

' 打开浏览器
shell.Run "http://localhost:8501"
