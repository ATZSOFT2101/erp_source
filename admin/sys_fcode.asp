    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 계열사 정보 관리                                   				--> 
    <!-- 내용: 계열사 정보 관리														--> 
    <!-- 작 성 자 : 문성원(050407)													--> 
    <!-- 작성일자 : 2006년 6월 19일													--> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
	<%
	Sql = "select * from Admin_fac_gubun  order by fac_Cd asc"
	Set Rs  =  syscon.execute(sql)
	%>
	<BODY CLASS="BODY_01" SCROLL="no">

    <!-- HEAD 부분 시작 -->
    <TABLE CLASS=Sch_Header>
	  <TR><TD><FORM NAME="form" ACTION="dast001.asp" method="post" >	  

		    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="99%">
		    <TR><TD></TD> 
				<TD ALIGN="right">
				
					<INPUT TYPE="button" CLASS=btn value="부서" Onclick="javascript:centerWindow('sys_dept.asp','sys_dept','600','500','no');">
					<INPUT TYPE="button" CLASS=btn value="직급" Onclick="javascript:centerWindow('sys_jik.asp','sys_jik','600','500','no');">

			</TD></TR>
			</TABLE>


	  </TD></FORM></TR>
    </TABLE>
    <!-- HEAD 부분 끝 -->

    <!-- CONTENT 부분 시작 -->
    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" style="TABLE-LAYOUT: fixed;">
    <TR >
		<TD ALIGN="CENTER" CLASS="TBL_RW_07" WIDTH="5%">SEQ</TD>
		<TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="10%">계열사명</TD>
		<TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="10%">회사코드</TD>
        <TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="7%">공장코드</td>
		<TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="8%">구분명</TD>
		<TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="10%">대표자명</TD>
		<TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="10%">사업자등록번호</TD>
		<TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="10%">전화번호</TD>
		<TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="10%">팩스번호</TD>
		
        <TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="10%">CODEPAGE</td>
        <TD ALIGN="CENTER" CLASS="TBL_RW_06" WIDTH="10%">등록일자</td>
    </TR>
    <%
    nVcnt = 1
    While Not Rs.EOF  
    
      fac_ucd           = RS("fac_cd")
      fac_unm           = Trim(RS("fac_nm"))
      fac_captin        = Trim(RS("fac_captin"))            : if IsNULL(fac_captin) THEN fac_captin="&nbsp; " END IF
      fac_bunho         = Trim(RS("fac_bunho"))             : if IsNULL(fac_bunho) THEN fac_bunho="&nbsp; " END IF
      fac_bupin         = Trim(RS("fac_bupin"))             : if IsNULL(fac_bupin) THEN fac_bupin="&nbsp; " END IF
      fac_post          = RS("fac_post")
      fac_address       = Trim(RS("fac_address"))
      fac_tel           = Trim(RS("fac_tel"))               : if IsNULL(fac_tel) THEN fac_tel="&nbsp; " END IF
      fac_fax           = Trim(RS("fac_fax"))               : if IsNULL(fac_fax) THEN fac_fax="&nbsp; " END IF
      fac_uptae         = Trim(RS("fac_uptae"))
      fac_jongmok       = Trim(RS("fac_jongmok"))
      fac_codepage      = Trim(RS("fac_codepage"))          
      fac_lang          = Trim(RS("fac_lang"))
      fac_gw_domain     = Trim(RS("fac_gw_domain"))
      fac_mail_domain   = Trim(RS("fac_mail_domain"))
      fac_wr_date       = RS("fac_wr_date")
      fac_ed_date       = RS("fac_ed_date")
      fac_sabun         = RS("fac_sabun")
      fac_fgubun        = RS("fac_fgubun")                  : if IsNULL(fac_fgubun) THEN fac_fgubun="&nbsp; " END IF
      fac_fgubun_nm     = RS("fac_fgubun_nm")               : if IsNULL(fac_fgubun_nm) THEN fac_fgubun_nm="&nbsp; " END IF

    %>
    <TR><!-- PROFILE 정보 -->
		<TD ALIGN="CENTER" WIDTH="5%" CLASS="TBL_DRW_00"><%=nVcnt%></td>
		<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><a href="javascript:centerWindow('sys_fcode_Wrt.asp?fac_ucd=<%=(fac_ucd)%>&fac_fgubun=<%=(fac_fgubun)%>&slt=조회','userwin','500','400','no');"><span class=blue_normal><%=(fac_unm)%></span></a></td>
		<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_03"><%=fac_ucd%></td>
		<TD ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01"><%=fac_fgubun%></td>
		<TD ALIGN="CENTER" WIDTH="8%" CLASS="TBL_DRW_01"><%=fac_fgubun_nm%></td>
		<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><%=fac_captin%></td>
		<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_03"><%=fac_bunho%></td>
		<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><%=fac_tel%></td>
		<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_03"><%=fac_fax%></td>
		<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><%=fac_lang%> (<%=fac_codepage%>)</td>
		<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><%=Left(fac_wr_date,11)%></td>
    </TR>       
    <%
    nVcnt = nVcnt + 1
    Rs.MoveNext
    Wend

    Rs.Close
    Set Rs = Nothing
    %>
    </TABLE>
    <!-- CONTENT 부분 끝 -->
    
    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->

