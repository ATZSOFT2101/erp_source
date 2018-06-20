    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 사용자 권한 관리                                   			--> 
    <!-- 내용: 프로그램별 사용자 권한 구성 관리                                     --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 14일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    SLT         =   TRIM(REQUEST("SLT"))
    FIELD       =   TRIM(REQUEST("FIELD"))
    VALUE       =   REQUEST("VALUE")
    REID        =   REQUEST("REID")
    PGID        =   REQUEST("PGID")
    PID         =   REQUEST("PID")

    IF  SLT = "process" THEN
        
        IF PGID <> "" and REID <> "" THEN
    
           IF FIELD="ALL" THEN
			  SYSCON.EXECUTE("DELETE MENU_Duty WHERE MR_MLVL_ID="& PGID &" AND MR_PRO_ID='"& REID &"'")
			  SYSCON.EXECUTE("INSERT INTO MENU_DUTY (MR_MLVL_ID, MR_PRO_ID, MR_READ, MR_ADD ,MR_DELETE, MR_PRINT, MR_EXCEL, MR_EDIT, MR_APPRO) Values ("& PGID &",'"& REID &"','"& TRIM(VALUE) &"','"& TRIM(VALUE) &"','"& TRIM(VALUE) &"' ,'"& TRIM(VALUE) &"' ,'"& TRIM(VALUE) &"' ,'"& TRIM(VALUE) &"','"& TRIM(VALUE) &"' )")
		   ELSE 
		      SQL =" SELECT * FROM MENU_DUTY WHERE MR_MLVL_ID="& PGID &" AND MR_PRO_ID='"& REID &"'"
			  SET RS=SYSCON.EXECUTE(SQL)

	          IF Rs.EOF THEN
		         SYSCON.EXECUTE("INSERT INTO MENU_DUTY (MR_MLVL_ID, MR_PRO_ID, MR_READ, MR_ADD ,MR_DELETE, MR_PRINT, MR_EXCEL, MR_EDIT, MR_APPRO) Values ("& PGID &",'"& REID &"','0','0','0' ,'0' ,'0' ,'0','0' ) ")
		         SYSCON.EXECUTE("UPDATE MENU_DUTY SET "& TRIM(FIELD) &"='"& TRIM(VALUE) &"' WHERE MR_MLVL_ID="& PGID &" AND MR_PRO_ID='"& REID &"'")
	          ELSE
		         SYSCON.EXECUTE("UPDATE MENU_DUTY SET "& TRIM(FIELD) &"='"& TRIM(VALUE) &"' WHERE MR_MLVL_ID="& PGID &" AND MR_PRO_ID='"& REID &"'")
	          END IF
		   END IF
	        
        ELSE
	        Alert_Message "PRGID 값이 넘어오지 않았습니다","1"	    
	    END IF
	    
	END IF
    
    IF PID = "" THEN  PID=324   END IF
    IF PGID = "" THEN  PGID=PID   END IF
    
	SQL = " SELECT ID,PGID,NAME FROM MLEVEL "
    SQL = SQL& " WHERE PID="& PID &" and  PGID <> 9999 Group by id, pgid, name "

    Set Rs = syscon.Execute(Sql)       
    %>
    <BODY CLASS="BODY_01" SCROLL="no" onload="resize_scroll(25);" onresize="resize_scroll(25);" >
    
    <!-- CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="640"  style="TABLE-LAYOUT: fixed;" >     
        <tr><td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="5%">SEQ</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">PRGID</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="29%">프로그래명</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">전체</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%" onclick="URL_move('sys_duty_cont.asp?SLT=process&field=Read_ALL&value=0&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">읽기</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">수정</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">등록</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">삭제</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">인쇄</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">엑셀</td> 
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="7%">결재</td>
        </TR>
        </TABLE>

        <DIV CLASS="scr_01" id=scr_01>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="640"  style="TABLE-LAYOUT: fixed;" >
	    <%
	    
	    NVCNT=1
	    WHILE NOT RS.EOF 
	    
	        ID			    =  RS("ID") 
	        PGID			=  RS("PGID") 
	        NAME			=  RS("NAME") 
        
			subSQL = " SELECT * FROM menu_duty WHERE mr_mlvl_id="& PGID &" and  MR_PRO_ID='"& REID &"'"

			Set subRs = syscon.Execute(subSQL)       
			
			IF Not subRS.EOF THEN
				MR_READ		=  subRs("MR_READ")       : IF ISNULL(MR_READ)    THEN MR_READ    ="0"      END IF
				MR_EDIT		=  subRs("MR_EDIT")       : IF ISNULL(MR_EDIT)    THEN MR_EDIT    ="0"      END IF
				MR_ADD		=  subRs("MR_ADD")        : IF ISNULL(MR_ADD)     THEN MR_ADD     ="0"      END IF
				MR_DELETE	=  subRs("MR_DELETE")     : IF ISNULL(MR_DELETE)  THEN MR_DELETE  ="0"      END IF
				MR_PRINT	=  subRs("MR_PRINT")      : IF ISNULL(MR_PRINT)   THEN MR_PRINT   ="0"      END IF
				MR_EXCEL	=  subRs("MR_EXCEL")      : IF ISNULL(MR_EXCEL)   THEN MR_EXCEL   ="0"      END IF
				MR_APPRO	=  subRs("MR_APPRO")      : IF ISNULL(MR_APPRO)   THEN MR_APPRO   ="0"      END IF
			ELSE
				MR_READ    ="0" 
				MR_EDIT    ="0" 
				MR_ADD     ="0"  
				MR_DELETE  ="0" 
				MR_PRINT   ="0" 
				MR_EXCEL   ="0" 
				MR_APPRO   ="0" 
			END IF
	    %>
        <tr >  
           <td ALIGN="CENTER" WIDTH="5%" CLASS="TBL_DRW_00"><%=(Nvcnt)%></td>
           <td ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><%=(pgid)%></td>
           <td ALIGN="LEFT" WIDTH="29%" CLASS="TBL_DRW_03"><span class=Link_class><a href="sys_duty_cont1.asp?PID=<%=PID%>&PGID=<%=(pgid)%>&PRG_name=<%=NAME%>" target=sub_d><%=(NAME)%></a></td>
           <td ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01">
	       <%
	       IF mr_Read="1" then
	       %>
	       <input type=button value='전체' class=btn onclick="URL_move('sys_duty_cont.asp?SLT=process&field=ALL&value=0&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
	       else
	       %>
	       <input type=button value='전체' class=btn onclick="URL_move('sys_duty_cont.asp?SLT=process&field=ALL&value=1&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
	       end if
	       %></td>
           <td ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01">
	       <%
	        SELECT CASE trim(mr_read)
		    CASE "1"
	       %>
	       <INPUT TYPE=checkbox name=mr_read value="1" checked onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_Read&value=0&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
		    CASE "0" 
	       %>
	       <INPUT TYPE=checkbox name=mr_read value="0" onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_Read&value=1&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');"> 
	       <%
		    end select
	       %>
	       </td>       
	       <td ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01">
	       <%
	        SELECT CASE trim(mr_edit)
		    CASE "1"
	       %>
	       <INPUT TYPE=checkbox name=mr_edit value="1" checked onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_edit&value=0&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
		    CASE "0"
	       %>
	       <INPUT TYPE=checkbox name=mr_edit value="0" onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_edit&value=1&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');"> 
	       <%
		    end select
	       %>
	       </td>

           <td ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01">
	       <%
	        SELECT CASE trim(mr_add)
		    CASE "1"
	       %>
	       <INPUT TYPE=checkbox name=mr_add value="1" checked onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_add&value=0&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
		    CASE "0"
	       %>
	       <INPUT TYPE=checkbox name=mr_add value="0" onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_add&value=1&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
		    end select
	       %>
	       </td>
           <td ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01">
	       <%
	        SELECT CASE trim(mr_delete)
		    CASE "1"
	       %>
	       <INPUT TYPE=checkbox name=mr_delete value="1" checked onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_delete&value=0&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
		    CASE "0"
	       %>
	       <INPUT TYPE=checkbox name=mr_delete value="0" onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_delete&value=1&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
		    end select
	       %></td>
           <td ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01">
	       <%
	        SELECT CASE trim(mr_print)
		    CASE "1"
	       %>
	       <INPUT TYPE=checkbox name=mr_print value="1" checked onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_print&value=0&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
		    CASE "0"
	       %>
	       <INPUT TYPE=checkbox name=mr_print value="0" onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_print&value=1&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');"> <%
		    end select
	       %></td>
           <td ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01">
	       <%
	        SELECT CASE trim(mr_excel)
		    CASE "1"
	       %>
	       <INPUT TYPE=checkbox name=mr_excel value="1" checked onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_excel&value=0&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
		    CASE "0"
	       %>
	       <INPUT TYPE=checkbox name=mr_excel value="0" onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_excel&value=1&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');"> <%
		    end select
	       %></td>
           <td ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01">
	       <%
	        SELECT CASE trim(mr_appro)
		    CASE "1"
	       %>
	       <INPUT TYPE=checkbox name=mr_appro value="1" checked onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_appro&value=0&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');">
	       <%
		    CASE "0"
	       %>
	       <INPUT TYPE=checkbox name=mr_appro value="0" onclick="URL_move('sys_duty_cont.asp?SLT=process&field=mr_appro&value=1&PGID=<%=pgid%>&REID=<%=REID %>&PID=<%=PID%>');"> <%
		    end select
	       %></td>

       </tr>
       <%

       Nvcnt=Nvcnt+1
	   PGID_PREV = PGID
       Rs.MoveNext
       Wend

       Rs.Close
       Set Rs = Nothing
       %>        
        </TABLE>
        </DIV>
        </DIV>
        <!-- CONTENT 부분 끝 -->
        
    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->


