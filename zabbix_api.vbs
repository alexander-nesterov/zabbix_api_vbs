Option Explicit

const ZABBIX_URL = "http://zabbix/api_jsonrpc.php"
const ZABBIX_USER = "Admin"
const ZABBIX_PASSWORD = "zabbix"

call ZabbixAuth
call ZabbixLogout

'--------------------------------------------------------------
Function ZabbixAuth

	Dim strJSON

	strJSON = "{""jsonrpc"": ""2.0"", ""method"": ""user.login"", ""params"":, ""user"":, ""password"":, ""id"":1}"

	'for debug
	PrintMessage strJSON
	
	PrintMessage ZabbixSend(strJSON)

End Function

'--------------------------------------------------------------
Function ZabbixSend(json)

	Dim objHTTP
	
	Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
	
	objHTTP.open "POST", ZABBIX_URL, FALSE
	
	objHTTP.setRequestHeader "Content-Type", "application/json"

	objHTTP.send json
	
	ZabbixSend = objHTTP.responseText
	
End Function

'--------------------------------------------------------------
Function ZabbixLogout

End Function

'--------------------------------------------------------------
Function PrintMessage(message)

	Dim WshShell

	Set WshShell = WScript.CreateObject("WScript.Shell")
	
	WScript.Echo message
	
End Function
