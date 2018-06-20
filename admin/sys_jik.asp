<!--#include file = "../inc/config.asp"-->
<!--#include file = "../inc/function/function.asp"-->
<%
mode   = Request("mode")
jikcd = Request("jikcd")
jiknm = Request("jiknm")

select case mode
case "save"
   Set rsCnt = syscon.Execute( "SELECT COUNT(*) cnt FROM ADMIN_JIK_GUBUN WHERE jik_cd = '" & jikcd & "'" )
   if rsCnt("cnt") = 0 then  bIsNew = True else bIsNew = False
   rsCnt.Close
   Set rsCnt = Nothing
   
   sql = "INSERT INTO ADMIN_JIK_GUBUN VALUES ( '" & jikcd & "', '" & jiknm & "' )"
   syscon.Execute( sql )
case "modify"
   sql = "UPDATE ADMIN_JIK_GUBUN " & _
         "   SET jik_nm = '" & jiknm & "' " & _
         " WHERE jik_cd = '" & jikcd & "'"
   syscon.Execute( sql )
case "del"
   sql = "DELETE FROM ADMIN_JIK_GUBUN WHERE jik_cd = '" & jikcd & "'"
   syscon.Execute( sql )
end select

Set rsgod = syscon.Execute( "SELECT * FROM ADMIN_JIK_GUBUN ORDER BY jik_cd ASC" )
%>
<html>
<head>
<title>관리자 : 직급구분 설정</title>
<link rel="stylesheet" type="text/css" href="../inc/css/DA_mainstyle.css">
<script language="Javascript" src="../inc/js/DA_script.js"></script>

</head>
<script language="javascript">
<!--
function savegod( frm )
{
   if ( '' == frm.jikcd.value )
   {
      alert( '분류코드를 입력하십시오.' );
      frm.jikcd.focus();
      return;
   }
   
   if ( '' == frm.jiknm.value )
   {
      alert( '직급을 입력하십시오.' );
      frm.jiknm.focus();
      return;
   }   

   if ( frm.jikcd.value == frm.oldcd.value )
   {
      frm.mode.value = 'modify';
   }
   else
   {
      frm.mode.value = 'save';
   }
         
   frm.submit();
}

function deletegod( frm )
{
   if ( '' == frm.jikcd.value || '' == frm.jiknm.value )
   {
      alert( '삭제할 직급을 선택하십시오.' );
      return;
   }
   
   frm.mode.value = 'del';
   frm.submit();
}

function selectgod( jikcd, jiknm )
{
   document.input_form.jikcd.value = jikcd;
   document.input_form.jiknm.value = jiknm;
   document.input_form.oldcd.value  = jikcd;
   document.input_form.oldnm.value  = jiknm;   
}

function ongodLoad()
{
   document.input_form.jikcd.focus();
   document.input_form.jikcd.select();
}
//-->
</script>

<Fieldset align=center style="width:95%;"><LEGEND style="font-weight:bold; font-size:9pt;" > 직급구분 설정 </LEGEND><div style="padding-top:10px;"></div>

<tablE  border="0" cellpadding="1" cellspacing="1">
   <form name="input_form" method="post" action="sys_jik.asp">
   <input type="hidden" name="mode">
   <input type="hidden" name="oldcd" value="">
   <input type="hidden" name="oldnm" value="">
   <tr><tD WIDTH=20></TD>
      <td align="right">분류코드&nbsp;</td>
      <td WIDTH=70><input type="text" name="jikcd" size="5" maxlength="5" class=input_box></td>
      <td align="right">직급분류&nbsp;</td>
      <td><input type="text" name="jiknm" size="20" maxlength="50" class=input_box></td>
      <td><input type="button" value=" 저장 " onClick="savegod(document.input_form);" class=btn>&nbsp;<input type="button" value=" 삭제 " onClick="deletegod(document.input_form);" class=btn></td>
   </tr></form>
</table>
<p>

<table width="95%"  border="0" cellpadding="4" cellspacing="1" align=center BGCOLOR=666666>
  <tr align=center> 
    <td bgcolor="#629AD9" CLASS=FFFFFF_BOLD>분류코드</td>
    <td bgcolor="#629AD9" CLASS=FFFFFF_BOLD>직급분류명</td>
    <td bgcolor="#629AD9" CLASS=FFFFFF_BOLD>분류코드</td>
    <td bgcolor="#629AD9" CLASS=FFFFFF_BOLD>직급분류명</td>
  </tr>
  <%
if rsgod.Bof and rsgod.Eof then
%> 
  <tr><td colspan="4" bgcolor="#F5F5F5" height="100" align="center">입력된 분류가 없습니다.</td></tr>
  <%
else
   do while not rsgod.Eof 
%> 
  <tr> 
    <td bgcolor="#dddddd" align="center"><%=rsgod("jik_cd")%></td>
    <td bgcolor="#ffffff"><a href="javascript:selectgod('<%=rsgod("jik_cd")%>', '<%=rsgod("jik_nm")%>');"><%=rsgod("jik_nm")%></a></td>
    <%      
      rsgod.MoveNext
      if rsgod.Eof then
%> 
    <td bgcolor="#dddddd" align="center">&nbsp;</td>
    <td bgcolor="#ffffff">&nbsp;</td>
    <%      
      else
%> 
    <td bgcolor="#dddddd" align="center"><%=rsgod("jik_cd")%></td>
    <td bgcolor="#ffffff"><a href="javascript:selectgod('<%=rsgod("jik_cd")%>', '<%=rsgod("jik_nm")%>');"><%=rsgod("jik_nm")%></a></td>
    <%      
         rsgod.MoveNext
      end if 
%> </tr>
  <%   
   loop
end if

rsgod.Close
Set rsgod=Nothing
syscon.Close
Set syscon=Nothing

%> 
</table>
<div style="padding-top:20px;"></div>
</Fieldset>
<div style="padding-top:20px;"></div>



</body>
</html>