    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 도움말 정보관리                              			        --> 
    <!-- 내용: Mlevel 구성 관리                                                     --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 20일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
	PRG_GUBUN1		= trim(REQUEST("PRG_GUBUN1"))
	SEARCHSTRING	= trim(REQUEST("SEARCHSTRING"))

	IF PRG_GUBUN1 = "WA1" or PRG_GUBUN1 = "" THEN
	   subSQL = " "
	ELSE
	   subSQL = " AND PRG_GUBUN1 ='"& PRG_GUBUN1 &"' "
	END IF

	IF SEARCHSTRING = "" THEN
	   srcSQL = " "
	ELSE
	   srcSQL = " AND PRG_NAME like '%"& SEARCHSTRING &"%' "
	END IF

    SQL = " SELECT PRG_ID, PRG_NAME, GUBUN_NAME, PRG_UID "
	SQL = SQL & " FROM PRGID "
	SQL = SQL & " Left JOIN GUBUN1 On Left(PRG_Gubun1,2)=GUBUN_CODE and right(PRG_Gubun1,2)=GUBUN_seq "
	SQL = SQL & " WHERE PRG_AUTH <> 'N' "& subSQL &" "& srcSQL
	SET RS = SYSCON.EXECUTE(SQL)
    %>
    <BODY CLASS="BODY_01" SCROLL="no" onload="resize_scroll(60);" onresize="resize_scroll(60);">
    
	<!-- HEAD 부분 시작 -->
    <TABLE CLASS="Sch_Header" id="TABLE1">
	  <TR><TD><FORM NAME="form" ACTION="sys_help_left.asp" method="post" >
		      <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" style="TABLE-LAYOUT: fixed;">
		      <TR><Td><% CALL FAC_slt1("","PRG_GUBUN1","WA",""& PRG_GUBUN1 &"") %></td>   
				  <Td width="100"><input type="text" name="searchstring" class="input" size=20 value="<%=searchstring%>"></td>
				  <TD><INPUT TYPE="submit" VALUE="검색" CLASS="btn" ></TD></FORM> 
              </TR>
              </TABLE>
	  </TD></TR>
    </TABLE>
    <!--  HEAD 부분 끝 -->  
   
	   
	
    <!-- CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="270"  style="TABLE-LAYOUT: fixed;" >
        <tr><TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">SEQ</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="25%">구분</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="65%">프로그램명</td>
        </tr>
        </TABLE>
        <DIV CLASS="scr_01" id=scr_01>
           <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="270" style="TABLE-LAYOUT: fixed;">

	        <%
	        VCNT=1
	        While Not Rs.EOF

	        PRG_id		=  Rs("PRG_id") 
	        PRG_name	=  Rs("PRG_name") 
	        GUBUN_NAME	=  Rs("GUBUN_NAME")	: IF isNULL(GUBUN_NAME) THEN GUBUN_NAME = "&nbsp;" END IF
        	
	        %>
            <TR><td ALIGN="CENTER" WIDTH="10%"  CLASS="TBL_DRW_00"><%=(Vcnt)%></td>
                <td ALIGN="CENTER" WIDTH="25%"  CLASS="TBL_DRW_01"><%=(GUBUN_NAME)%></td>
                <td ALIGN="LEFT" WIDTH="65%"  CLASS="TBL_DRW_03"><A href="sys_help_cont.asp?PRG_id=<%=(PRG_id)%>&PRG_name=<%=(PRG_name)%>"  target="sub_c"><span class=blue_normal><%=(PRG_name)%></span></a></td>
           </tr>
           <%
           Rs.MoveNext
           Vcnt=Vcnt+1
           Wend

           Rs.Close
           Set Rs = Nothing
            %>        
            </table> 
       </DIV>
    </DIV>   
    
    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->    
         
