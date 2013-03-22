'����������, ����������ͻᱻ��
Const AccessLimit = 150

'IIS��Active Directory·��������������վ����IIS��·������IIS��������Բ鿴��ID����������վ��IDΪ1����Ϊ��IIS://LocalHost/W3SVC/1/ROOT
Const IISADString = "IIS://LocalHost/W3SVC/1/ROOT"

Dim StdIn, StdOut
Set StdIn = WScript.StdIn
Set StdOut = WScript.StdOut

Dim AllowList
AllowList=Array()
LoadAllowList

'���������б�Ϊ�գ���ִ����Ӧ�Ĵ���
if ( 2 > 1 ) Then
	Dim Input, Parts
	Dim AccessCount, AccessIp
	'��Stdin�л��IP�����б�
	While Not StdIn.AtEndOfStream
		Input = StdIn.ReadLine
		Parts = Split(Trim(Input), Chr(32))
		If UBound(Parts) = 2 Then
			AccessCount = Parts(0)
			AccessIp = Parts(1)
			AccessUserAccount = Parts(2)

			RunLog AccessIp,AccessCount,AccessUserAccount
			'StdOut.WriteLine AccessCount + " " + AccessIp + " " + AccessUserAccount
			If CInt(AccessCount) >= AccessLimit Then
				If InList(AllowList,AccessIp)=-1 Then	
					'StdOut.WriteLine AccessCount + " " + AccessIp
					DoDenyIP AccessIp,AccessCount
				End If
			End If
		End If
	Wend
End If
DisplayIpList

'ִ�з�IP�����ĺ�����ע��IIS6��IIS7������ͬ��������IIS7�Ľű���������Բο�΢��IIS�Ĺٷ���վ��
Sub DoDenyIP(AccessIp, AccessCount)
	Set adminManager = WScript.CreateObject("Microsoft.ApplicationHost.WritableAdminManager")
	adminManager.CommitPath = "MACHINE/WEBROOT/APPHOST"
	'������һ�е�MySites.web�滻��������վ��IIS�е����ơ�
	Set ipSecuritySection = adminManager.GetAdminSection("system.webServer/security/ipSecurity", "MACHINE/WEBROOT/APPHOST/MySites.web")
	Set ipSecurityCollection = ipSecuritySection.Collection

	Set addElement = ipSecurityCollection.CreateNewElement("add")
	addElement.Properties.Item("ipAddress").Value = AccessIp
	addElement.Properties.Item("allowed").Value = False
	ipSecurityCollection.AddElement(addElement)

	adminManager.CommitChanges()

	LogDenyIP AccessIp,AccessCount
End Sub

'��������LOG�������ҽ�LOG������D:\iapps\AutoBlockIP�ļ����¡�
Sub RunLog(AccessIp, AccessCount, AccessUserAccount)
	Const ForAppending = 8
	Dim fso, f
	Set fso = CreateObject( "Scripting.FileSystemObject" )
	Set f = fso.OpenTextFile( "D:\iapps\AutoBlockIP\Runing.log", ForAppending, True )
	f.WriteLine AccessCount + Chr(32) + AccessIp + Chr(32) + AccessUserAccount + Chr(32) + CStr(Now)
	f.Close() 
End Sub

'���汻��IP��LOG
Sub LogDenyIP(AccessIp, AccessCount)
	StdOut.writeline "Ip Added : " & AccessIp

	Const ForAppending = 8
	Dim fso, f
	Set fso = CreateObject( "Scripting.FileSystemObject" )
	Set f = fso.OpenTextFile( "D:\iapps\AutoBlockIP\DenyIPList.log", ForAppending, True )
	f.WriteLine AccessIp + Chr(32) + AccessCount + Chr(32) + CStr(Now)
	f.Close() 
End Sub

Function InList(IPList, AccessIp)
	InList = -1
	Dim i
	For i = 0 to UBound(IPList)
		If AccessIp = IpList(i) Then
			InList = i
		End If
	Next
End Function

Sub DisplayIpList()
	Set IIsWebVirtualDirObj = GetObject(IISADString) 
	Set IIsIPSecurityObj = IIsWebVirtualDirObj.IPSecurity 
	Dim IPList 
	IPList = Array() 

	If True = IIsIPSecurityObj.GrantByDefault Then
		IPList = IIsIPSecurityObj.IPDeny 
		StdOut.WriteLine "����б�"
		For i = 0 to UBound( IPList )
			StdOut.WriteLine IPList(i)
		Next
	Else
		StdOut.WriteLine "��ǰ��ֹ����IP���ʡ�"
	End If
End Sub

Sub LoadAllowList()
	Const ForReading = 1
	Dim fso, f
	Set fso = CreateObject( "Scripting.FileSystemObject" )
	Set f = fso.OpenTextFile( "D:\iapps\AutoBlockIP\AllowIpList.txt", ForReading )

	Dim Input
	While not f.AtEndOfStream
		Input = f.ReadLine
		ReDim Preserve AllowList(UBound(AllowList)+1)
		AllowList(UBound(AllowList)) = Input
	Wend
	f.Close()
End Sub
