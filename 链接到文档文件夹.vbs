const IOException=52
const OperationCanceledException=5

Main

Sub Main()
	set fso=createobject("scripting.filesystemobject")
	set sh=CreateObject("Wscript.Shell")
	ra3=sh.SpecialFolders("MyDocuments")+"\Red Alert 3"
	if not fso.FolderExists(ra3) then
		Err.Raise IOException,,"δ���ҵ�ԭ��쾯3�ļ��С�����ȷ��װ�쾯3�������������һ��ս����"
	end if
	mods=ra3+"\Mods"
	if not fso.FolderExists(mods) then
		if msgbox("֮ǰ�ƺ�û�а�װ���κ�Mod��Ҫ����Mods�ļ�����", vbYesNo or vbQuestion,"����Mods")=vbYes then
			fso.CreateFolder(mods)
		else
			Err.Raise OperationCanceledException,,"�������û�ȡ��"
		end if
	end if
	modPath=wscript.arguments(0)
	modName=mid(modPath,instrrev(modPath,"\")+1)
	msgbox modName, , "���ڴ�������mod:"
	set cmd=sh.Exec("cmd.exe")
	cmd.StdIn.WriteLine "cd /d """+mods+""""
	cmd.StdIn.WriteLine "mklink /d """ + modName + """ """ + modPath + """"
	msgbox "��������",vbSystemModal or vbInformation,"����Mods�ļ�����ʲô�仯"
End Sub