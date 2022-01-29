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
<title>��̨���� - С�������ռ�</title>
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
	<h3 style="text-align: center">��̨���� - С�������ռ� 1.0 ���ð�</h3>
    <div class="nav">
    	<a href="analysis.asp" target="_blank">��ѯ����ͳ��</a> | 
    	<a href="ssoLogout.asp" target="_self">�˳�</a>
    </div>
<div class="add">
    	<form action="addRecord.asp" method="post">
	        <p>
        		<label for="oDate">��¼���ڣ�</label>
        		<input id="oDate" name="oDate" type="text" placeholder="��һ������" />
			</p>
            <p>
            <label for="oTime">��¼ʱ�䣺</label>
            <input id="oTime" name="oTime" type="text" placeholder="��һ��ʱ��" />
            </p>
            <p>
            <label for="oLog">���˵˵��</label>
            <input id="oLog" name="oLog" type="text" placeholder="��ʲô�������Ļ�" />
            </p>
            <p>
            
            </p>
            <p>
            	<label for="oEmotion">��������</label>
            	<select id="oEmotion" name="oEmotion">
            		<option value="happy">����</option>
                	<option value="normal">һ��</option>
            		<option value="sad">ɥ</option>
            	</select>
            	<input name="submit" type="submit" value="�Ǽ�" />
                <input name="reset" type="reset" value="����" />
                <input name="refresh" type="button" value="ˢ��" onclick="putSysTimeIntoBox()" />
			</p>
        </form>
    </div>
	<table style="text-align: center; width: 100%">
        <tr>
        	<th>#</th>
        	<th>����</th>
            <th>ʱ��</th>
            <th>������</th>
            <th>��־��¼</th>
            <th>�������</th>
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
            <td><a href="javascript:;" onclick="alert('<%=rs("rLog") %>')" target="_self">�鿴</a></td>
            <td><a href="modify.asp?id=<%=rs("ID") %>" target="_self">�鿴</a>|<a href="javascript:;" onclick="removeItemByConfirm(<%=rs("ID") %>)" target="_self">ɾ��</a></td>
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
			if (confirm("ȷ��ɾ�����Ϊ" + itemId + "�Ĵ򿨼�¼�𣿴˲����޷���������������ҵ΢����ͬ�����������")) {
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
