    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 전자결재 문서코드관리                              		    --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 08일                                                 --> 
    <!-- 내    용 : 전자결재 문서코드관                                             --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    SQL = "select ID, dcode, dname, dpart, pname, pname1, doc_fac, doc_dept, doc_type1, doc_dbname,usechk from docmst where usechk='Y'" 	
    Set Rs  =  Dbcon.execute(SQL)
    %>

    <BODY CLASS="BODY_01" SCROLL="no"  onload="resize_scroll(30);" onresize="resize_scroll(30);">
    
    <!-- HEAD 부분 시작 
    <TABLE CLASS=Sch_Header id=d2>
      <TR><TD>
      
      </TD></TR>
    </TABLE>-->


    <!-- CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
        <TR><td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="5%">ID</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="15%">문서코드</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="30%">문서명</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">공장</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">부서</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">TYPE</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">참조DB</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">사용여부</td>
     </tr>
     </TABLE>
     <DIV CLASS="scr_01" id=scr_01>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980" style="TABLE-LAYOUT: fixed;">
	    <%
        While Not Rs.EOF  
            
        nVcnt       =  nVcnt  +  1
	    ID		    =  RS("ID") 
	    dcode	    =  RS("dcode") 
	    dname	    =  RS("dname") 
	    dpart	    =  RS("dpart") 
	    doc_fac	    =  RS("doc_fac") 
	    doc_dept    =  RS("doc_dept") 
	    doc_type1   =  RS("doc_type1") 
	    doc_dbname  =  RS("doc_dbname") 

	    usechk	 =  RS("usechk")
    	
	    Select case usechk
	    case "Y"
	    usechk_txt="가능"
	    case "N"
	    usechk_txt="불가"
	    end select

	    %>
        <TR><TD ALIGN="CENTER" WIDTH="5%"  CLASS="TBL_DRW_00"><%=ID%></td>
            <TD ALIGN="CENTER" WIDTH="15%"  CLASS="TBL_DRW_01"><%=dcode%></td>
            <TD ALIGN="LEFT" WIDTH="30%"  CLASS="TBL_DRW_03"><A  onclick="centerWindow('sys_appro_code_wrt.asp?id=<%=id%>&slt=조회','apdmin_win','500','520','no');" style="cursor:hand;"><span class=blue_normal><%=dname%></span></a></td> 
            <TD ALIGN="CENTER" WIDTH="10%"  CLASS="TBL_DRW_01"><%=doc_fac%></td>
            <TD ALIGN="CENTER" WIDTH="10%"  CLASS="TBL_DRW_01"><%=doc_dept%></td>
            <TD ALIGN="CENTER" WIDTH="10%"  CLASS="TBL_DRW_01"><%=doc_type1%></td>
            <TD ALIGN="CENTER" WIDTH="10%"  CLASS="TBL_DRW_01"><%=doc_dbname%></td>
            <TD ALIGN="CENTER" WIDTH="10%"  CLASS="TBL_DRW_01"><%=usechk_txt%></td>
        </tr>
	    <%
        Rs.MoveNext
        Wend

        Rs.Close
        Set Rs = Nothing
	    %>        
    </table>
    </div>
    </DIV>
    
    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
    

