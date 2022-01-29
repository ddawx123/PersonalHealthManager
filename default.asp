<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
dim conn, connstr, db, rs
db = "xieob.accdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath(db)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=0.7, maximum-scale=2.0, user-scalable=no">
<title>小丁健康日记</title>
<style type="text/css">
	.nav {
		padding-bottom: 9px;
		text-align: right;
	}
	table {
		border-collapse: collapse;
		border: 1px solid #000000;
	}
	table tr th, table tr td {
		border: 1px solid #000000;
	}
	.footer {
		bottom: 0px;
		position: fixed;
	}
</style>
</head>

<body>
	<h3 style="text-align: center">小丁健康日记 1.0 自用版</h3>
    <div class="nav">
    	<a href="analysis.asp" target="_blank">查询数据统计</a> | 
    	<a href="manager.asp" target="_blank">进入信息核录</a>
    </div>
	<table style="text-align: center; width: 100%">
        <tr>
        	<th>#</th>
        	<th>日期</th>
            <th>时间</th>
            <th>情绪打卡</th>
            <th>日志记录</th>
        </tr>
        <%
		conn.Open connstr
		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open "select top 2 * from record", conn, 1, 2
		Do while not rs.eof
		%>
        <tr>
        	<td><% Response.Write(rs("ID")) %></td>
        	<td><% Response.Write(rs("rDate")) %></td>
            <td><% Response.Write(rs("rTime")) %></td>
            <td><% Response.Write(rs("rEmotion")) %></td>
            <td><a href="javascript:;" onclick="alert('<% Response.Write(rs("rLog")) %>')" target="_self">查看</a></td>
        </tr>
        <% rs.movenext %>
        <%
		Loop
		rs.close
		conn.close
        %>
        <h5>更多数据已作隐藏，需要通过企业微信侧应用或网页登陆后访问。</h5>
    </table>
    <div class="footer">&copy; 2012-2019 DingStudio Technology All Rights Reserved</div>
</body>
</html>
