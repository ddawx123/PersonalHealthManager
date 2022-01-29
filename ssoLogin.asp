<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Dim ScriptAddress,Servername,qs,loginResult,allowUserName

'配置允许登录系统的用户名
allowUserName = "ding"

ScriptAddress = CStr(Request.ServerVariables("SCRIPT_NAME"))
Servername = CStr(Request.ServerVariables("Server_Name"))
GetUrlWithoutParam ="http://"& Servername & ScriptAddress
qs=Request.QueryString
If qs<>"" Then
	GetUrl ="http://"& Servername & ScriptAddress &"?"&qs
Else
	GetUrl ="http://"& Servername & ScriptAddress
End If
If Request.QueryString("ticket") = "" Then
	Response.Redirect("https://id.dscitech.com/cas/login?renew=1&service=" & Server.URLEncode(GetUrl))
Else
	loginResult = GetHttpPage("http://10.33.0.8/cas/serviceValidate?ticket="&Request.QueryString("ticket")&"&service="&Server.URLEncode(GetUrlWithoutParam),"UTF-8")
	Session.Timeout = 3
	Session("loginUrl") = GetUrlWithoutParam
	If Instr(loginResult, "<cas:user>" & allowUserName & "</cas:user>") > 0 Then
		' 设置登录态3分钟超时
		Session("authData") = loginResult
		Response.Redirect("./manager.asp")
		Response.End()
	Else
		Response.Redirect("./ssoRefuse.htm")
		Response.End()
	End If

End If

Function GetHttpPage(url, charset) 
	Dim http
	Set http = Server.createobject("Msxml2.ServerXMLHTTP")
	http.Open "GET", url, false
	http.Send()
	If http.readystate<>4 Then
	Exit Function 
    End If 
	GetHttpPage = BytesToStr(http.ResponseBody, charset)
	Set http = Nothing
End function


Function BytesToStr(body, charset)
	Dim objStream
    Set objStream = Server.CreateObject("Adodb.Stream")
	objStream.Type = 1
	objStream.Mode = 3
	objStream.Open
	objStream.Write body
	objStream.Position = 0
	objStream.Type = 2
	objStream.Charset = charset
	BytesToStr = objStream.ReadText
	objStream.Close
	Set objStream = Nothing
End Function
%>