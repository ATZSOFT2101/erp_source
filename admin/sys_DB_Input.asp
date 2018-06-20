    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : DB Query analyzer                              					--> 
    <!-- 내용: DB Query analyzer												    --> 
    <!-- 작 성 자 : 문성원(050407)                                                    --> 
    <!-- 작성일자 : 2006년 6월 19일                                                   --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->

	<%
	Dim chrObjectType
	Dim strObjectName
	Dim strDataBase

	chrObjectType = Request("ObjectType")	' u : 테이블, p : 스토어드 프로시저
	strObjectName = Request("ObjectName")	' 오브젝트의 이름
	strDataBase = Request("DataBase")		' 데이터베이스의 이름

	if ( chrObjectType = "" ) then	chrObjectType = "u"		' 값이 없으면 테이블로 설정
	if ( strDataBase = "" ) then	strDataBase = "syerp"	' 데이터베이스가 입력되지 않았으면
															' pubs 디비로 설정
	%>
    <BODY CLASS="BODY_01" SCROLL="no">


	<SCRIPT LANGUAGE="JavaScript">
	<!--
		function ChangeObject( s )
		{
			location.href = 'sys_db_Input.asp?DataBase=<%=strDataBase%>&ObjectType=' + s.options[s.selectedIndex].value;
		}

		function SetSelectStatement()
		{
			var f = document.QueryForm;
			f.QueryStatement.value = "Select top 100 * From " + f.ObjectName.options[f.ObjectName.selectedIndex].value;
		}

		function SetInsertStatement()
		{
			var f = document.QueryForm;
			f.QueryStatement.value = "Insert Into " + f.ObjectName.options[f.ObjectName.selectedIndex].value + "(  )\nvalues (  )";
		}

		function SetUpdateStatement()
		{
			var f = document.QueryForm;
			f.QueryStatement.value = "Update " + f.ObjectName.options[f.ObjectName.selectedIndex].value + "\nSet Column1 = Value1 \nWhere WhereStatement";
		}

		function SetDeleteStatement()
		{
			var f = document.QueryForm;
			f.QueryStatement.value = "Delete From " + f.ObjectName.options[f.ObjectName.selectedIndex].value + "\nWhere WhereStatement";
		}

		function SetHelpStatement()
		{
			var f = document.QueryForm;
			f.QueryStatement.value = "sp_Help " + f.ObjectName.options[f.ObjectName.selectedIndex].value;
		}
	//-->
	</SCRIPT>

    <!-- HEAD 부분 시작 --> 
    <TABLE CLASS=Sch_Header id=d2>
	  <TR><TD>
	<table width = "100%" cellpadding = 0 cellspacing =0 border=0>
		<tr><form name = QueryForm method="POST" action="sys_Db_OutPut.asp" target=sub_c1>
			<td width=120>현재DB : <B><%=strDataBase%></B></td>
			<td><select name="ObjectType" OnChange="javascript:ChangeObject( this );">
			<option value = p <% if ( chrObjectType = "p" ) then Response.Write "Selected" end if %>>S_Procedure</option>
			<option value = u <% if ( chrObjectType = "u" ) then Response.Write "Selected" end if %>>Table</option>
			</select>
			<select name="ObjectName">
			<%


			' 시스템 테이블에서 원하는 오브젝트의 이름을 찾아온다.
			strSQL = "Select Name From " & strDataBase & "..SysObjects Where Type = '" & chrObjectType & "'Order By Name"

			Set objRs = Conn.Execute( strSQL )

			if not ( ( objRs.BOF = True ) or ( objRs.EOF = True ) ) then
				Do Until ( objRs.EOF = True )
			%>
			<option value = <%=objRs(0)%>><%=objRs(0)%></option>
			<%
			objRs.MoveNext
			Loop
			end if

			objRs.Close
			Set objRs = Nothing

			%>
			</select></td>
			<td>

			<%	if ( chrObjectType = "u" ) then	%>
			<input type="button" value="Select" OnClick = "javascript: SetSelectStatement();" class=btn>
			<input type="button" value="Insert" OnClick = "javascript: SetInsertStatement();" class=btn>
			<input type="button" value="Update" OnClick = "javascript: SetUpdateStatement();" class=btn>
			<input type="button" value="Delete" OnClick = "javascript: SetDeleteStatement();" class=btn>
			<%	end if	%>
			<input type="button" value="Info View"  OnClick = "javascript: SetHelpStatement();" class=btn>
			<input type="submit" value="Query Execute" class=btn></td>
		</tr><input type = hidden name = DataBase value = "<%=strDataBase%>">
	</td></tr>
	</table>

	</td></tr>
	</table>


    <!-- CONTENT 부분 시작 -->
    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%"  style="TABLE-LAYOUT: fixed;" >
		<tr><td><textarea rows="6" name="QueryStatement" cols="72" style="width:100%;height:260px;" class=input></textarea></td></form></tr>
	</table>
    <!-- CONTENT 부분 끝 -->

    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->

