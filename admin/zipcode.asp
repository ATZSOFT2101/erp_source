
<!--#include Virtual = "/INC/CONFIG.ASP"                -->
<html>

<head>
<title>주소록 : 추가하기</title>
<meta name="generator" content="Namo WebEditor v4.0">
<link rel="stylesheet" type="text/css" href="../inc/css/DA_mainstyle.css">
<script language="javascript">
function send(form) {
	if (form.dong.value=="") {
		alert("\n 동을 입력하세요.");
		form.dong.focus();
	    return;
	}
	form.submit();
}
</script>
</HEAD>
<BODY onload="form.dong.focus();" class=body_04>


<form method=get action="zip_search.asp" name=form>
<table cellpadding="0" cellspacing=0 border=0 width="100%" bgcolor=ffffff align=center>
<tr><Td  background="<%=img_path%>cal_top_bg.gif" style="padding-left:5px;" height="33">
	<Table border=0 cellpadding=0 cellspacing=0 width="100%">
	<Tr><Td style="padding-left:13px;" class=black_bold> 주소찾기 : 우편번호 검색</td>
		<TD align=right style="padding-right:10px;"><a href="javascript:close();"><img src="<%=img_path%>window_close.gif" border=0></a></td>
	</tr>
	</table>
</td></tr>
<Tr><Td height=10></td></tr>
<tr><td align=center>동을 입력하여주시기바랍니다.</td></tr>
<Tr><Td align=center>
	<Table>
	<Tr><Td align=right><INPUT type="text" name=dong class=input_box /></td>
	<Td ><Input type=button onclick='javascrtip:send(this.form)' value="검색" class=btn></td></tr>
	</table>
</tr>
<tr><td align=center>(예) "<FONT color=blue>전북 군산시 나운동</FONT>" 일경우 "<span class=ff0000_normal>나운동</span>"</td></form></tr>
</table>

</BODY>
</HTML>
