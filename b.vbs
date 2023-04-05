Set WshShell = CreateObject("wscript.Shell")
UserProfilePath = WshShell.ExpandEnvironmentStrings("%UserProfile%")
PathFile = UserProfilePath & "\Drivers\SysUtils\Driven\Tree\booger.bat"
MsgBox qq(PathFile)
Return = WshShell.Run(qq(PathFile),1)

Function qq(strIn)
    qq = Chr(34) & strIn & Chr(34)
End Function 