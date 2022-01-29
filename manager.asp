<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("authData") = "" Then
Response.Redirect("./ssoLogin.asp")
Response.End()
End If
%>
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
<title>后台管理 - 小丁健康日记</title>
<style type="text/css">
	.nav {
		padding-bottom: 9px;
		text-align: right;
	}
	.add {
		padding-bottom: 9px;
		text-align: center;
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
	<h3 style="text-align: center">后台管理 - 小丁健康日记 1.0 自用版</h3>
    <div class="nav">
    	<a href="analysis.asp" target="_blank">查询数据统计</a> | 
    	<a href="ssoLogout.asp" target="_self">退出</a>
    </div>
<div class="add">
    	<form action="addRecord.asp" method="post">
	        <p>
        		<label for="oDate">记录日期：</label>
        		<input id="oDate" name="oDate" type="text" placeholder="填一下日期" />
			</p>
            <p>
            <label for="oTime">记录时间：</label>
            <input id="oTime" name="oTime" type="text" placeholder="填一下时间" />
            </p>
            <p>
            <label for="oLog">随便说说：</label>
            <input id="oLog" name="oLog" type="text" placeholder="有什么想留档的话" />
            </p>
            <p>
            
            </p>
            <p>
            	<label for="oEmotion">情绪卡：</label>
            	<select id="oEmotion" name="oEmotion">
            		<option value="happy">开心</option>
                	<option value="normal">一般</option>
            		<option value="sad">丧</option>
            	</select>
            	<input name="submit" type="submit" value="登记" />
                <input name="reset" type="reset" value="重置" />
                <input name="refresh" type="button" value="刷新" onclick="putSysTimeIntoBox()" />
			</p>
        </form>
    </div>
	<table style="text-align: center; width: 100%">
        <tr>
        	<th>#</th>
        	<th>日期</th>
            <th>时间</th>
            <th>情绪打卡</th>
            <th>日志记录</th>
            <th>管理操作</th>
        </tr>
        <%
		conn.Open connstr
		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.open "select * from record", conn, 1, 2
		Do while not rs.eof
		%>
        <tr>
        	<td><%=rs("ID") %></td>
        	<td><%=rs("rDate") %></td>
            <td><%=rs("rTime") %></td>
            <td><%=rs("rEmotion") %></td>
            <td><a href="javascript:;" onclick="alert('<%=rs("rLog") %>')" target="_self">查看</a></td>
            <td><a href="modify.asp?id=<%=rs("ID") %>" target="_self">查看</a>|<a href="javascript:;" onclick="removeItemByConfirm(<%=rs("ID") %>)" target="_self">删除</a></td>
        </tr>
        <% rs.movenext %>
        <%
		Loop
		rs.Close
		conn.Close
        %>
    </table>
    <div class="footer">&copy; 2012-2019 DingStudio Technology All Rights Reserved</div>
    <script type="text/javascript">
		function putSysTimeIntoBox() {
	    	document.getElementById("oDate").value = new Date().getFullYear() + "-" + parseInt(new Date().getMonth()+1) + "-" + new Date().getDate();
			document.getElementById("oTime").value = (new Date().getHours()<10?"0"+new Date().getHours():new Date().getHours()) + ":" + (new Date().getMinutes()<10?"0"+new Date().getMinutes():new Date().getMinutes()) + ":" + (new Date().getSeconds()<10?"0"+new Date().getSeconds():new Date().getSeconds());
		}
		function removeItemByConfirm(itemId) {
			if (confirm("确认删除编号为" + itemId + "的打卡记录吗？此操作无法撤消，并将在企业微信中同步反馈结果。")) {
				var frmSbt = document.createElement("form");
				frmSbt.action = "remove.asp?actionUTC=" + new Date().getTime();
				frmSbt.method = "post";
				var dataEle = document.createElement("input");
				dataEle.name = "id";
				dataEle.type = "hidden";
				dataEle.value = itemId;
				frmSbt.appendChild(dataEle);
				document.body.appendChild(frmSbt);
				frmSbt.submit();
			}
		}
		putSysTimeIntoBox();
    </script>
</body>
</html>
