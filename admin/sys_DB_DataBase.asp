    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : DB Query analyzer                              					--> 
    <!-- 내용: DB Query analyzer														--> 
    <!-- 작 성 자 : 문성원(050407)                                                    --> 
    <!-- 작성일자 : 2006년 6월 19일                                                   --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->

    <BODY CLASS="BODY_01" SCROLL="no" onload="resize_scroll(50);" onresize="resize_scroll(50);">
    
    <!-- CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="140"  style="TABLE-LAYOUT: fixed;" >
        <TR>
            <TD ALIGN="CENTER" class="TBL_RW_00" WIDTH="30%">SEQ</td>
            <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="70%">DB명</td>
		</tr>
		</table>
		<DIV class="scr_01" id="scr_01">
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="140"  style="TABLE-LAYOUT: fixed;" >
        <%

			' 시스템 테이블에서 데이터 베이스의 목록을 가져온다.
			strSQL = "Select Name From master.sys.SysDataBases Order By Name"

			Set objRs = Conn.Execute( strSQL )

			' 데이터베이스의 이름을 브라우저에 출력한다.
			if not ( ( objRs.BOF = True ) or ( objRs.EOF = True ) ) then
				
				Dim i,bgcolor
				i=0
				Do Until ( objRs.EOF = True )


			if objRs(0) <> "master" and  objRs(0) <> "Pubs" and objRs(0) <> "Northwind" and objRs(0) <> "tempdb" and objRs(0) <> "msdb" and objRs(0) <> "model" then
		%>

		<TR><Td align=center CLASS="TBL_DRW_00" width="30%"><%=i%></td>
			<Td align=left style="padding-left:3px;" CLASS="TBL_DRW_01" width="70%"><a href = "sys_Db_Input.asp?DataBase=<%=objRs(0)%>" target=sub_c><%=objRs(0)%></a></td>
		</tR>
		<%
		end if
		i=i+1
		objRs.MoveNext
		Loop
		end if

		objRs.Close
		Set objRs = Nothing

		%>
	</table>
	</div>
	</div>
        <!-- CONTENT 부분 끝 -->

    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->
