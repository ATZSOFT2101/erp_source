    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
<%
Set rsgod = syscon.Execute( "SELECT * FROM ADMIN_DEPT_GUBUN ORDER BY dept_cd ASC" )
%>

<table width="100%"  border="0" cellpadding="4" cellspacing="1" align=center BGCOLOR=666666>
  <%
if rsgod.Bof and rsgod.Eof then
%> 
  <tr><td  colspan="4" bgcolor="#F5F5F5" height="100" align="center">입력된 분류가 없습니다.</td></tr>
  <%
else
   do while not rsgod.Eof 
%> 
  <tr> 
    <td width="20%" bgcolor="#dddddd" align="center"><%=rsgod("dept_cd")%></td>
    <td width="30%" bgcolor="#ffffff"><a href="javascript:selectgod('<%=rsgod("dept_cd")%>', '<%=rsgod("dept_nm")%>');"><%=rsgod("dept_nm")%></a></td>
    <%      
      rsgod.MoveNext
      if rsgod.Eof then
%> 
    <td bgcolor="#dddddd" align="center">&nbsp;</td>
    <td bgcolor="#ffffff">&nbsp;</td>
    <%      
      else
%> 
    <td width="20%" bgcolor="#dddddd" align="center"><%=rsgod("dept_cd")%></td>
    <td width="30%" bgcolor="#ffffff"><a href="javascript:selectgod('<%=rsgod("dept_cd")%>', '<%=rsgod("dept_nm")%>');"><%=rsgod("dept_nm")%></a></td>
    <%      
         rsgod.MoveNext
      end if 
%> </tr>
  <%   
   loop
end if


%> 
</table>
