scrPath=Wscript.ScriptFullName
scrDir=left(scrPath,instrrev(scrPath,"\"))
set uac=CreateObject("Shell.Application")
for each arg in wscript.arguments
    uac.ShellExecute "wscript.exe",""""+scrDir+"���ӵ��ĵ��ļ���.vbs"+""" """+arg+"""","","runas",1
next