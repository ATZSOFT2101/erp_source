<!--#include Virtual = "/INC/CONFIG.ASP"                -->
<html>

<head>
<title>주소록 : 추가하기</title>
<meta name="generator" content="Namo WebEditor v4.0">
<link rel="stylesheet" type="text/css" href="../inc/css/DA_mainstyle.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function returnvalue(formName) {
    opener.zipReturn(formName.zipcode.value,formName.zipcode1.value,formName.addr.value);
	window.close();
}

//-->
</SCRIPT>
<%
dong = Request.QueryString("dong")

strSQL="select * from zipcode where dong like '%"& dong &"%'"
set rs = syscon.execute(strSQL)

%>
<style type="text/css">
.JJ 	{padding:0px;overflow-x:hidden;overflow-y:auto;vertical-align:top;width:100%;}
</style>
<body class=body_04 >

<table cellpadding="0" cellspacing=0 border=0 width="100%" bgcolor=ffffff align=center>
<tr><Td  background="<%=img_path%>cal_top_bg.gif" style="padding-left:5px;" height="33">
	<Table border=0 cellpadding=0 cellspacing=0 width="100%">
	<Tr><Td style="padding-left:13px;" class=black_bold> 주소록 추가</td>
		<TD align=right style="padding-right:10px;"><a href="javascript:close();"><img src="<%=img_path%>window_close.gif" border=0></a></td>
	</tr>
	</table>
</td></tr>
</table>

<div class="JJ" style="width:100%;height:92%;">
<table cellpadding="0" cellspacing=0 border=0 width="100%" bgcolor=ffffff align=center>
<Tr><Td align=center valign=top>

	<table border="0" width="100%" cellspacing="0" cellpadding="2" > 
		<Tr><Td height=30 colspan=5 class=9FB4DC_normal>* 일치되는 우편번호를 선택하시면 자동으로 입력됩니다.</td></tr>
		<tr><td colspan=5 bgcolor="cccccc" height="1"></td></tr>
		<TR bgcolor=f2f2f2><td align=center>우편번호</td>
			<td align=center>지역</td>
			<td align=center>구</td>
			<td align=center>동</td>
			<td align=center>번지</td>
		</tr>
		<tr><td bgcolor="d7d7d7" height="1" colspan=5></td></tr>
		<%
		i=0
		do while not rs.EOF
		i=i+1
		formName="form"&Cstr(i)
		addr=rs(1)&" "&rs(2)&" "&rs(3)&" "&rs(4)
		%>
		<form method=post name="<%=formName%>" onsubmit="returnvalue(this)">
		<tr>
			<input type=hidden name=zipcode value="<%=left(rs(0),3)%>">
			<input type=hidden name=zipcode1 value="<%=right(rs(0),3)%>">    
			<input type=hidden name=addr value="<%=addr%>">
			<td align=center><a href="javascript:returnvalue(<%=formName%>)"><span class=blue_bold><%=rs(0)%></span></a></td>
			<td align=center><%=rs(1)%></td>
			<td align=center><%=rs(2)%></td>    
			<td align=center><%=rs(3)%></td>
			<%
			   if not rs(4)="" then 
			   Response.Write "<td align=center>"&rs(4)&"</td>"
			   else
			   Response.Write "<td>&nbsp;</td>"
			   end if
			%>
		</tr>
		<tr><td bgcolor="d7d7d7" height="1" colspan=5></td></tr>
		</form>
		<%
		rs.movenext
		loop
		rs.close
		set rs=nothing
		set syscon=nothing
		%>
	</table>
</td></tr>
<Tr><Td height=20></td></tr>
<tr><td height="1" bgcolor="d7d7d7"></td></tr>
<Tr><Td bgcolor=f1f1f1 height=40>

	<table align="center" border="0" width="95%" height=40 >
		<tr><td align=right><a href="javascript:history.back()"><img src="<%=img_path%>move_51x23.gif" border="0"></a></td>
		</tr>
	</table>


</td></tr>
</table>
</div>

</BODY>
</HTML>
