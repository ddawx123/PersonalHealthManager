<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Dim loginUrl
sLoginUrl = Session("loginUrl")
Session.Abandon()
Response.Redirect("https://id.dscitech.com/cas/logout?service=" & Server.URLEncode(sLoginUrl))
Response.End()
%>
