<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("authData") = "" Then
Response.Redirect("./ssoLogin.asp")
Response.End()
End If
%>
<%
If Request.Form("id") = "" Then
	Response.Write("��֧�ִ�����")
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
		Response.Write("<script type='text/javascript'>alert('ɾ���ɹ�����ͬ��������˽����ƽ̨');location.replace('manager.asp');</script>")
	Else
		Response.Write("<script type='text/javascript'>alert('������д�����쳣���±��β���ʧ��');location.replace('manager.asp');</script>")
	End If
Else
	Response.Write("<script type='text/javascript'>alert('û���ҵ��˱�ţ��ü�¼�����Ѿ�ɾ����');location.replace('manager.asp');</script>")
	rs.Close()
	conn.Close()
	Response.End()
End If
%>