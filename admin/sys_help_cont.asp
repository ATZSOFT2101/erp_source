    <!-- ========================================================================== -->
    <!-- 프로그램명 : 도움말 정보관리                              			        --> 
    <!-- 내용: 도움말 정보관리                                                      --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 20일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <%
    PRG_id			= REQUEST("PRG_id")
    PRG_name		= REQUEST("PRG_name")
    
    IF PRG_id <> "" THEN
       subSQL=" WHERE help_prg_id="& PRG_id 
	   Sql = "SELECT * FROM SYS_HELP "& subsQL
	ELSE
	   Sql = "SELECT * FROM SYS_HELP WHERE HELP_ID = '000000' "
	END IF
    
    Set Rs = syscon.Execute(Sql)
    %>
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <BODY CLASS="BODY_01" SCROLL="no" onload="resize_scroll(40);" onresize="resize_scroll(40);">


    <!-- HEAD 부분 시작 --> 
    <TABLE CLASS=Sch_Header id=d2>
	  <TR><TD><form name="check" method="post" action="sys_menu_cont.asp?slt=삭제&pid=<%=Request("id")%>">
		      <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="700">
		      <TR><TD><B>현재 프로그램명 :</B> <%=PRG_name%></TD>
				  <TD ALIGN="right">
				  <% IF PRG_ID <> "" THEN %>
				  <input TYpe="button" value="등록" onclick="javascript:centerWindow('sys_help_wrt.asp?help_prg_id=<%=prg_id%>&PRG_name=<%=prg_name%>','sys_help_wrt','700','700','n');" class="btn">
				  <% ELSE %>
				  <input TYpe="button" value="등록" onclick="alert('왼쪽 단위업무를 먼저 선택해 주십시오');" class="btn">
				  <% END IF %>
				  </TD></TR>
			  </TABLE>
	  </TD></TR>
    </TABLE>
    <!--  HEAD 부분 끝 -->     

    <!-- CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="700"  style="TABLE-LAYOUT: fixed;" >
        <tr><TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">SEQ</td>
			<TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">등록자</td>
			<TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">PRGID</td>
			<TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="70%">도움말제목</td>
		</tr>
        </TABLE>
        <DIV CLASS="scr_01" id=scr_01>
           <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="700" style="TABLE-LAYOUT: fixed;">		
		    <%
		    While Not Rs.EOF

		    nVcnt		  =  nVcnt  +  1
		    help_idx	  =  Rs("help_idx") 
		    help_prg_id	  =  Rs("help_prg_id") 
		    help_title	  =  Rs("help_title") 
		    help_id	      =  Rs("help_id") 
		    help_content  =  Rs("help_content") 	

		    %>
		    <TR><td ALIGN="CENTER" WIDTH="10%"  CLASS="TBL_DRW_00"><%=(help_idx)%></td>
		        <td ALIGN="CENTER" WIDTH="10%"  CLASS="TBL_DRW_01"><%=(help_id)%>&nbsp;</td>
		        <td ALIGN="CENTER" WIDTH="10%"  CLASS="TBL_DRW_03"><a href="javascript:centerWindow('sys_help_Wrt.asp?help_idx=<%=help_idx%>&help_prg_id=<%=help_prg_id%>&slt=조회&PRG_name=<%=prg_name%>&mode=view','sys_search_win','700','700');"><span class=Link_class><%=(help_prg_id)%></span></a>&nbsp;</td></td>
		        <td ALIGN="LEFT" WIDTH="70%"  CLASS="TBL_DRW_01"><%=(help_title)%>&nbsp;</td>
	       </tr>
	       <%
	       Rs.MoveNext
	       Wend
	       Rs.Close
	       Set Rs = Nothing
	       %>  
	       </table>
	   </DIV>
    </DIV>	   
    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->    
