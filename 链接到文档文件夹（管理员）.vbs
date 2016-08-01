scrPath=Wscript.ScriptFullName
scrDir=left(scrPath,instrrev(scrPath,"\"))
set uac=CreateObject("Shell.Application")
for each arg in wscript.arguments
    uac.ShellExecute "wscript.exe",""""+scrDir+"链接到文档文件夹.vbs"+""" """+arg+"""","","runas",1
next