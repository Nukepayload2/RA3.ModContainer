const IOException=52
const OperationCanceledException=5

Main

Sub Main()
	set fso=createobject("scripting.filesystemobject")
	set sh=CreateObject("Wscript.Shell")
	ra3=sh.SpecialFolders("MyDocuments")+"\Red Alert 3"
	if not fso.FolderExists(ra3) then
		Err.Raise IOException,,"未能找到原版红警3文件夹。请正确安装红警3，并且至少完成一局战斗。"
	end if
	mods=ra3+"\Mods"
	if not fso.FolderExists(mods) then
		if msgbox("之前似乎没有安装过任何Mod。要创建Mods文件夹吗？", vbYesNo or vbQuestion,"链接Mods")=vbYes then
			fso.CreateFolder(mods)
		else
			Err.Raise OperationCanceledException,,"操作被用户取消"
		end if
	end if
	modPath=wscript.arguments(0)
	modName=mid(modPath,instrrev(modPath,"\")+1)
	msgbox modName, , "正在处理以下mod:"
	set cmd=sh.Exec("cmd.exe")
	cmd.StdIn.WriteLine "cd /d """+mods+""""
	cmd.StdIn.WriteLine "mklink /d """ + modName + """ """ + modPath + """"
	msgbox "操作结束",vbSystemModal or vbInformation,"看看Mods文件夹有什么变化"
End Sub