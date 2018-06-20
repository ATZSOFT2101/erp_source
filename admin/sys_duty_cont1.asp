    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 사용자 권한 관리                                   			--> 
    <!-- 내용: 사용자별 권한 구성 관리                                              --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 15일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    css1     = " style=""font-family:arial;font-size:12px;"" "
    
	'초기화 작업 (불필요 데이터 삭제처리)
	SYSCON.EXECUTE ("DELETE MENU_DUTY  WHERE MR_READ='0' and MR_ADD='0' and MR_DELETE='0' and MR_PRINT='0' and MR_EXCEL='0' and MR_EDIT='0' and MR_APPRO='0'")

    SLT	          = REQUEST("SLT")    
    PID	          = REQUEST("PID")
    PGID	      = REQUEST("PGID")
    PRG_NAME	  = REQUEST("PRG_NAME")
    FORMFAC       = REQUEST("FORMFAC")
    REID          = REQUEST("REID")
    FIELD         = REQUEST("FIELD")
    VALUE         = REQUEST("VALUE")
    
    IF FORMFAC <> "" THEN 
       FAC_SQL =" AND FAC='"& FORMFAC &"' "
    ELSE
       FORMFAC ="F0001"
       FAC_SQL =" AND FAC='F0001' "     '공장구분값이 없을경우 기본값으로 본사로 한다.
    END IF
    
    FORMDEPT       = REQUEST("FORMDEPT")
    IF FORMDEPT <> "" THEN 
       DEPT_SQL =" AND DEPT='"& FORMDEPT &"' "
    END IF

    IF  SLT = "process" THEN
        
        IF PGID <> "" and REID <> "" THEN
    	    
           IF FIELD="ALL" THEN
			  
			  SYSCON.EXECUTE("DELETE MENU_Duty WHERE MR_MLVL_ID="& PGID &" AND MR_PRO_ID='"& REID &"'")
			  
			  IF TRIM(VALUE)="1" THEN
			  SYSCON.EXECUTE("INSERT INTO MENU_DUTY (MR_MLVL_ID, MR_PRO_ID, MR_READ, MR_ADD ,MR_DELETE, MR_PRINT, MR_EXCEL, MR_EDIT, MR_APPRO) Values ("& PGID &",'"& REID &"','"& TRIM(VALUE) &"','"& TRIM(VALUE) &"','"& TRIM(VALUE) &"' ,'"& TRIM(VALUE) &"' ,'"& TRIM(VALUE) &"' ,'"& TRIM(VALUE) &"','"& TRIM(VALUE) &"' )")
			  END IF
           ELSE
		      SQL =" SELECT * FROM MENU_DUTY WHERE MR_MLVL_ID="& PGID &" AND MR_PRO_ID='"& REID &"'"
			  SET RS=SYSCON.EXECUTE(SQL)

	          IF Rs.EOF THEN

		         SYSCON.EXECUTE("INSERT INTO MENU_DUTY (MR_MLVL_ID, MR_PRO_ID,MR_READ, MR_ADD ,MR_DELETE, MR_PRINT, MR_EXCEL, MR_EDIT, MR_APPRO) Values ("& PGID &",'"& REID &"','0','0','0' ,'0' ,'0' ,'0','0' )")
		         SYSCON.EXECUTE("UPDATE MENU_DUTY SET "& TRIM(FIELD) &"='"& TRIM(VALUE) &"' WHERE MR_MLVL_ID="& PGID &" AND MR_PRO_ID='"& REID &"'")

	          ELSE
		         SYSCON.EXECUTE("UPDATE MENU_DUTY SET "& TRIM(FIELD) &"='"& TRIM(VALUE) &"' WHERE MR_MLVL_ID="& PGID &" AND MR_PRO_ID='"& REID &"'")
	          END IF

			  RS.Close : SET RS=NOTHING
           END IF
	        
	    ELSE
	        Alert_Message "PRGID 값이 넘어오지 않았습니다","1"	    
	    END IF
	    
	END IF



    IF  PGID <> "" THEN
        
        SQL = " SELECT SUNG,ID,"
        SQL = SQL &" (SELECT top 1 MR_READ   FROM MENU_DUTY WHERE MR_PRO_ID=ID AND  MR_MLVL_ID= "& PGID &") AS MR_READ,   "
        SQL = SQL &" (SELECT top 1 MR_ADD    FROM MENU_DUTY WHERE MR_PRO_ID=ID AND  MR_MLVL_ID= "& PGID &") AS MR_ADD,    " 
        SQL = SQL &" (SELECT top 1 MR_DELETE FROM MENU_DUTY WHERE MR_PRO_ID=ID AND  MR_MLVL_ID= "& PGID &") AS MR_DELETE, "
        SQL = SQL &" (SELECT top 1 MR_PRINT  FROM MENU_DUTY WHERE MR_PRO_ID=ID AND  MR_MLVL_ID= "& PGID &") AS MR_PRINT,  "
        SQL = SQL &" (SELECT top 1 MR_EXCEL  FROM MENU_DUTY WHERE MR_PRO_ID=ID AND  MR_MLVL_ID= "& PGID &") AS MR_EXCEL,  "
        SQL = SQL &" (SELECT top 1 MR_EDIT   FROM MENU_DUTY WHERE MR_PRO_ID=ID AND  MR_MLVL_ID= "& PGID &") AS MR_EDIT,   "
        SQL = SQL &" (SELECT top 1 MR_APPRO  FROM MENU_DUTY WHERE MR_PRO_ID=ID AND  MR_MLVL_ID= "& PGID &") AS MR_APPRO   "
        SQL = SQL &" FROM PROFILE WHERE IDUSE <> 'N' "& FAC_SQL &" "& DEPT_SQL &" order by sung "
        'RESPONSE.WRITE SQL
        SET RS = SYSCON.EXECUTE(SQL)
    


    %>
    <BODY CLASS="BODY_01" SCROLL="no" onload="resize_scroll(60);" onresize="resize_scroll(60);" style="background-color:#C3DAF9;margin:0;border:0;">

    <!-- HEAD 부분 시작 --> 
    <TABLE CLASS=Sch_Header id=d2>
	  <TR><TD><form name=form action="sys_duty_cont1.asp" method="post">
			  <Input type=hidden name="PGID" value="<%=PGID%>">
			  <Input type=hidden name="PRG_name" value="<%=PRG_name%>">  
		      
		      <TABLE BORDER='0' CELLPADDING='0' CELLSPACING='0' WIDTH="99%">
		      <TR><TD><B>프로그램명</B> :</span> <%=PRG_name%></TD>
			      <TD align=right>공장 
				    <select  name="formfac">
				    <option value=""></option>
				    <%
				    sqlgroup=" Select distinct fac_cd,fac_nm From ADMIN_FAC_GUBUN order by fac_Cd asc"
				    Set Rsgroup = syscon.Execute(sqlgroup)

				    Do until Rsgroup.EOF
    				  
				    IF trim(formfac) = trim(Rsgroup("fac_cd")) then
				       Response.write   "<option value='"& Rsgroup("fac_cd") &"' selected>"& Rsgroup("fac_nm") &"</option>"
				    ELSE
				       Response.write   "<option value='"& Rsgroup("fac_cd") &"'>"& Rsgroup("fac_nm") &"</option>"
				    END IF
    						  
				    Rsgroup.Movenext
				    Loop

				    Rsgroup.close
				    Set Rsgroup = Nothing
				    %>
				    </select> 
				    부서 
				    <select  name="formdept">
				    <option value=""></option>
				    <%
				    sqlgroup=" Select * From ADMIN_DEPT_GUBUN order by dept_Cd asc"
				    Set Rsgroup = syscon.Execute(sqlgroup)

				    Do until Rsgroup.EOF
    				  
				    IF trim(formdept) = trim(Rsgroup("dept_cd")) then
				       Response.write   "<option value='"& Rsgroup("dept_cd") &"'' selected>"& Rsgroup("dept_nm") &"</option>"
				    ELSE
				       Response.write   "<option value='"& Rsgroup("dept_cd") &"'>"& Rsgroup("dept_nm") &"</option>"
				    END IF
    						  
				    Rsgroup.Movenext
				    Loop

				    Rsgroup.close
				    Set Rsgroup = Nothing
				    %>
				    </select></td>
			      <TD width=40><input type=submit  value="검색" class=btn></td></form>
			  </TR>
			  </TABLE>
	  </TD></TR>
    </TABLE>
    <!--  HEAD 부분 끝 --> 
   
    <!-- CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER='0' CELLPADDING='0' CELLSPACING='0' WIDTH="640"  style="TABLE-LAYOUT: fixed;" >     
        <tr><td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="5%">SEQ</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">사번</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="29%">성명</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">전체</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">읽기</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">수정</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">등록</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">삭제</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">인쇄</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">엑셀</td> 
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">결재</td>
        </TR>
        </TABLE>

        <DIV CLASS="scr_01" id=scr_01>
        <TABLE BORDER='0' CELLPADDING='0' CELLSPACING='0' WIDTH="640"  style="TABLE-LAYOUT: fixed;" >
	    <%
	    nVcnt=1
	    While Not Rs.EOF

	    reid            =  Rs("id")
	    sung            =  Rs("sung")
	    mr_read			=  Rs("mr_read") 
	    mr_edit			=  Rs("mr_edit") 
	    mr_add			=  Rs("mr_add") 
	    mr_delete		=  Rs("mr_delete") 
	    mr_print		=  Rs("mr_print") 
	    mr_excel		=  Rs("mr_excel") 
	    mr_appro		=  Rs("mr_appro")
    	
	    %>
        <tr>  
           <td ALIGN="CENTER" WIDTH="05%" CLASS="TBL_DRW_00" <%=css1%> style="border-left:none;"><%=(Nvcnt)%></td>
           <td ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01" <%=css1%>><%=(reid)%></td>
           <td ALIGN="CENTER" WIDTH="29%" CLASS="TBL_DRW_03" <%=css1%>><span class=Link_class><%=(sung)%></a></td>
           <td ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01" <%=css1%>>
	       <%
	       IF mr_read="1" then
	          Response.write "<input type=button value='전체' class=btn onclick=""URL_move('sys_duty_cont1.asp?FIELD=ALL&value=0&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">"
	       else
	          Response.write "<input type=button value='전체' class=btn onclick=""URL_move('sys_duty_cont1.asp?FIELD=ALL&value=1&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">"
	       end if
	       %></td>
           <td ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01">
	       <%
	       IF mr_read = "1" THEN
	          Response.write "<INPUT TYPE=checkbox name=mr_read value='1' checked onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_Read&value=0&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">"
		   ELSE
	          Response.write "<INPUT TYPE=checkbox name=mr_read value='0' onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_Read&value=1&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">" 
		   END IF
	       %>
	       </td>       
	       <td ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01">
	       <%
	       IF mr_edit = "1" THEN
	          Response.write "<INPUT TYPE=checkbox name=mr_edit value='1' checked onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_edit&value=0&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">"
		   ELSE
	          Response.write "<INPUT TYPE=checkbox name=mr_edit value='0' onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_edit&value=1&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">" 
		   END IF
	       %>
	       </td>
           <td ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01">
	       <%
	       IF mr_add = "1" THEN
	          Response.write "<INPUT TYPE=checkbox name=mr_add value='1' checked onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_add&value=0&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">"
		   ELSE
	          Response.write "<INPUT TYPE=checkbox name=mr_add value='0' onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_add&value=1&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">"
		   END IF
	       %>
	       </td>
           <td ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01">
	       <%
	       IF mr_delete = "1" THEN
	          Response.write "<INPUT TYPE=checkbox name=mr_delete value='1' checked onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_delete&value=0&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">"
		   ELSE
	          Response.write "<INPUT TYPE=checkbox name=mr_delete value='0' onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_delete&value=1&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">"
		   END IF
	       %></td>
           <td ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01">
	       <%
	       IF mr_print = "1" THEN
	          Response.write "<INPUT TYPE=checkbox name=mr_print value='1' checked onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_print&value=0&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">"
	       ELSE
	          Response.write "<INPUT TYPE=checkbox name=mr_print value='0' onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_print&value=1&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">" 
		   END IF
	       %></td>
           <td ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01">
	       <%
	       IF mr_excel = "1" THEN
	          Response.write "<INPUT TYPE=checkbox name=mr_excel value='1' checked onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_excel&value=0&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"&PRG_name="& PRG_name &"');"">"
		   ELSE
	          Response.write "<INPUT TYPE=checkbox name=mr_excel value='0' onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_excel&value=1&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">" 
		   END IF
	       %></td>
           <td ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01">
	       <%
	       IF mr_appro = "1" THEN
	          Response.write "<INPUT TYPE=checkbox name=mr_appro value='1' checked onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_appro&value=0&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"&PRG_name="& PRG_name &"');"">"
		   ELSE
	          Response.write "<INPUT TYPE=checkbox name=mr_appro value='0' onclick=""URL_move('sys_duty_cont1.asp?FIELD=mr_appro&value=1&PGID="& PGID &"&REID="& reid &"&SLT=process&formfac="& formfac &"&formdept="& formdept &"');"">" 
		   END IF
	       %></td>

       </tr>
       <%
       Nvcnt=Nvcnt+1
       Rs.MoveNext
       Wend

       Rs.Close
       Set Rs = Nothing
       %>        
       </TABLE>
       </DIV>
       </DIV>
       <!-- CONTENT 부분 끝 -->
       <%
       ELSE       
       %>
       <BODY CLASS="BODY_01" SCROLL="no" >

       
       <!-- HEAD 부분 시작 --> 
        <TABLE CLASS=Sch_Header id=TABLE1>
	      <TR><TD>
		          <TABLE BORDER='0' CELLPADDING='0' CELLSPACING='0' WIDTH="99%">
		          <TR><TD align=center class=paging_white_normal >왼쪽 트리에서 먼저 선택해 주십시오.</td></TR>
			      </TABLE>
	      </TD></TR>
        </TABLE>
        <!--  HEAD 부분 끝 -->        
       <%
       
       END IF
       %> 
    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->