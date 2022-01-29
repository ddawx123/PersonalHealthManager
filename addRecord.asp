<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("authData") = "" Then
Response.Redirect("./ssoLogin.asp")
Response.End()
End If
%>
<%
If Request.Form("oDate") = "" Or Request.Form("oTime") = "" Or Request.Form("oLog") = "" Or Request.Form("oEmotion") = "" Then
	Response.Write("<script type='text/javascript'>alert('表单参数初筛后反馈字段不完整');location.replace('manager.asp');</script>")
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
rs.open "select * from record where rDate='"&Request.Form("oDate")&"' and rTime='"&Request.Form("oTime")&"'", conn, 1, 2
If Not rs.eof Then
	Response.Write("<script type='text/javascript'>alert('已有同日期时间的记录，如需修改请直接在原记录上进行。');location.replace('manager.asp');</script>")
	rs.Close()
	conn.Close()
	Response.End()
Else
	rs.Close()
	sqlcode = "insert into record (rDate,rTime,rEmotion,rLog) values ('"&Request.Form("oDate")&"','"&Request.Form("oTime")&"','"&Request.Form("oEmotion")&"','"&Request.Form("oLog")&"')"
	conn.Execute sqlcode, rt
	If rt > 0 Then
		Response.Write("<script type='text/javascript'>alert('打卡成功且已上传数据至私有云平台');location.replace('manager.asp');</script>")
	Else
		Response.Write("<script type='text/javascript'>alert('因数据写访问异常导致本次打卡失败');location.replace('manager.asp');</script>")
	End If
	conn.Close()
End If
%>