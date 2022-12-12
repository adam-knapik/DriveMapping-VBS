'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Mapowanie udzialu oraz skrot na pulpicie
'Network drive mapping and dekstop shortcut
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Dysk udostępniony
Set objNetwork = createobject("WScript.Network")
 
objNetwork.MapNetworkDrive "S:", "\\192.168.0.2\shared"

	Set objShell = CreateObject("Wscript.Shell")
	strDesktopFolder = objShell.SpecialFolders("Desktop")
 
	Set objShortCut = objShell.CreateShortCut(strDesktopFolder &  "\shared.lnk")
	objShortCut.TargetPath="S:\"
	objShortCut.Description="shared"
	objShortCut.Save
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Mapowanie udzialu oraz skrot na pulpicie w zaleznosci od uzytkownika
'Network drive mapping and dekstop shortcut depending on the user
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

strUser = CreateObject("WScript.Network").UserName
'Wscript.Echo strUser 'print UserName

Select Case strUser
	Case "admin"
		'Wscript.Echo "admin test" 'print test
		Set objNetwork = createobject("WScript.Network")
		
		objNetwork.MapNetworkDrive "P:", "\\192.168.0.2\sharedAdmin"

			Set objShell = CreateObject("Wscript.Shell")
			strDesktopFolder = objShell.SpecialFolders("Desktop")
		 
			Set objShortCut = objShell.CreateShortCut(strDesktopFolder &  "\sharedAdmin.lnk")
			objShortCut.TargetPath="P:\"
			objShortCut.Description="sharedAdmin"
			objShortCut.Save
			
	Case "user"
		'Wscript.Echo "USER test" 'print test
		Set objNetwork = createobject("WScript.Network")
		
		objNetwork.MapNetworkDrive "R:", "\\192.168.0.2\sharedUser"

			Set objShell = CreateObject("Wscript.Shell")
			strDesktopFolder = objShell.SpecialFolders("Desktop")
		 
			Set objShortCut = objShell.CreateShortCut(strDesktopFolder &  "\sharedUser.lnk")
			objShortCut.TargetPath="R:\"
			objShortCut.Description="sharedUser"
			objShortCut.Save
End Select

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Mapowanie udzialu oraz skrot na pulpicie w zaleznosci od grupy
'Network drive mapping and dekstop shortcut depending on the user group
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set objSysInfo = createobject("ADSystemInfo")
strUserPath = "LDAP://" & objSysInfo.UserName
 
Set objUser = GetObject(strUserPath)
 
For Each strGroup in objUser.MemberOf
   strGroupPath = "LDAP://" & strGroup
    Set objGroup = GetObject(strGroupPath)
    strGroupName = objGroup.CN
 
'Wscript.Echo strGroupName

Select Case strGroupName
	Case "Administrators"
		'Wscript.Echo "Administrators test"
		
		Set objNetwork = CreateObject("Wscript.Network")
 
	    objNetwork.MapNetworkDrive "L:", "\\192.168.0.2\sharedAdmin"
			Set objShell = CreateObject("Wscript.Shell")
			strDesktopFolder = objShell.SpecialFolders("Desktop")
			Set objShortCut = objShell.CreateShortCut(strDesktopFolder &  "\sharedAdmin.lnk")
 
			objShortCut.TargetPath="L:\"
			objShortCut.Description="sharedAdmin"
			objShortCut.Save
		
 End Select
 
 
Select Case strGroupName

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Koniec
'End
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Select
Next
Wscript.Quit