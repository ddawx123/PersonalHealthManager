<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("authData") = "" Then
Response.Redirect("./ssoLogin.asp")
Response.End()
End If
%>
<%
If Request.Form("id") = "" Then
	Response.Write("不支持此请求")
	Response.End()
End If
%>
<%
dim conn, connstr, db, rs, rt
db = "xieob.accdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath(db)
conn.Open connstr
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "select * from record where ID="&CLng(Request.Form("id")), conn, 1, 2
If Not rs.eof Then
	rs.Close()
	sqlcode = "delete from record where ID="&CLng(Request.Form("id"))
	conn.Execute sqlcode, rt
	if rt > 0 Then 
		Response.Write("<script type='text/javascript'>alert('删除成功且已同步数据至私有云平台');location.replace('manager.asp');</script>")
	Else
		Response.Write("<script type='text/javascript'>alert('因数据写访问异常导致本次操作失败');location.replace('manager.asp');</script>")
	End If
Else
	Response.Write("<script type='text/javascript'>alert('没有找到此编号，该记录可能已经删除。');location.replace('manager.asp');</script>")
	rs.Close()
	conn.Close()
	Response.End()
End If
%>