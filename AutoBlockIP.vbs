'访问量上限, 超过这个量就会被封
Const AccessLimit = 150

'IIS的Active Directory路径，这里是您网站所在IIS的路径，在IIS属性里可以查看其ID，比如我网站的ID为1，即为：IIS://LocalHost/W3SVC/1/ROOT
Const IISADString = "IIS://LocalHost/W3SVC/1/ROOT"

Dim StdIn, StdOut
Set StdIn = WScript.StdIn
Set StdOut = WScript.StdOut

Dim AllowList
AllowList=Array()
LoadAllowList

'如果传入的列表不为空，则执行相应的代码
if ( 2 > 1 ) Then
	Dim Input, Parts
	Dim AccessCount, AccessIp
	'从Stdin中获得IP访问列表
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

'执行封IP操作的函数，注意IIS6和IIS7有所不同，这里是IIS7的脚本，具体可以参考微软IIS的官方网站。
Sub DoDenyIP(AccessIp, AccessCount)
	Set adminManager = WScript.CreateObject("Microsoft.ApplicationHost.WritableAdminManager")
	adminManager.CommitPath = "MACHINE/WEBROOT/APPHOST"
	'下面这一行的MySites.web替换成您的网站在IIS中的名称。
	Set ipSecuritySection = adminManager.GetAdminSection("system.webServer/security/ipSecurity", "MACHINE/WEBROOT/APPHOST/MySites.web")
	Set ipSecurityCollection = ipSecuritySection.Collection

	Set addElement = ipSecurityCollection.CreateNewElement("add")
	addElement.Properties.Item("ipAddress").Value = AccessIp
	addElement.Properties.Item("allowed").Value = False
	ipSecurityCollection.AddElement(addElement)

	adminManager.CommitChanges()

	LogDenyIP AccessIp,AccessCount
End Sub

'保存运行LOG，这里我将LOG保存在D:\iapps\AutoBlockIP文件夹下。
Sub RunLog(AccessIp, AccessCount, AccessUserAccount)
	Const ForAppending = 8
	Dim fso, f
	Set fso = CreateObject( "Scripting.FileSystemObject" )
	Set f = fso.OpenTextFile( "D:\iapps\AutoBlockIP\Runing.log", ForAppending, True )
	f.WriteLine AccessCount + Chr(32) + AccessIp + Chr(32) + AccessUserAccount + Chr(32) + CStr(Now)
	f.Close() 
End Sub

'保存被封IP的LOG
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
		StdOut.WriteLine "封禁列表："
		For i = 0 to UBound( IPList )
			StdOut.WriteLine IPList(i)
		Next
	Else
		StdOut.WriteLine "当前禁止所有IP访问。"
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
