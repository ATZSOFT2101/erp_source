    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 그룹웨어 결재목록 관리                              			--> 
    <!-- 내용: 그룹웨어 결재목록 관리                                               --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 19일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    css1 = " style=""font-family:arial;font-size:12px;"" "
    
    pagesize    =   25
    searchtext  =   trim(Request("searchtext"))
    key			=   trim(Request("key"))
    MDB         =   trim(Request("MDB"))
    chk         =   trim(Request("chk"))
    GotoPage    =   Request("GotoPage")
    
    if GotoPage = "" then GotoPage = 1
  
    select case chk
    case "A" 
       ssql = " and a.stat = 10 "
    case "B"
       ssql = " and a.stat <> 10 " 
    case "C"
       ssql = " and a.stat = 12 " 
    end select


    if  searchtext  =  ""   Then
	    SQL = "select count(idx) as recCount from "& MDB &"..docmt a where makedate > '2005-01-01' "& ssql  
    else
	    SQL = "select count(idx) as recCount from "& MDB &"..docmt a where makedate > '2005-01-01' "& ssql &" "
	    sql = sql &" and "& key &" like '%" & searchtext & "%' and makedate > '2005-01-01'"
    end if
    Set Rs = Dbcon.Execute(SQL)
    
    recordCount = Rs(0)
    pagecount = int((recordCount-1)/pagesize) +1
    

    if  searchtext  =  ""   Then
        SQL = "SELECT  top "& pagesize &" a.idx, dcode, dname, makeuser,b.sung as makename, a.stat, formindex, "
        sql = sql &" docdate, makedate,apmax, apuser,c.sung as apname, subject "
        sql = sql &" FROM "& MAIL_DB &"..docmt a "
        sql = sql &" left join profile b on makeuser = b.id "
        sql = sql &" left join profile c on apuser = c.id "
        SQL = SQL &" WHERE makedate > '2005-01-01' "& ssql &" "
        sql = sql &" and a.idx not in (select top " & ((GotoPage - 1) * pagesize) & " idx from "& MAIL_DB &"..docmt a WHERE makedate > '2005-01-01' "& ssql &" "
        sql = sql &" ORDER BY idx DESC) order by a.idx desc "	
    else
        SQL = "SELECT top "& pagesize &" a.idx, dcode, dname, makeuser,b.sung as makename, a.stat, formindex, "
        sql = sql &" docdate, makedate,apmax, apuser,c.sung as apname, subject "
        sql = sql &" FROM "& MAIL_DB &"..docmt a "
        sql = sql &" left join profile b on makeuser = b.id "
        sql = sql &" left join profile c on apuser = c.id "
        SQL = sql &" WHERE makedate > '2005-01-01' "& ssql &" and  "& key &" like '%" & searchtext & "%' "
        sql = sql &" and a.idx not in (select top " & ((GotoPage - 1) * pagesize) & " idx from "& MAIL_DB &"..docmt a WHERE makedate > '2005-01-01' "& ssql &" "
        sql = sql &" and  "& key &" like '%" & searchtext & "%' ORDER BY idx DESC) order by a.idx desc "

    end if    

    Set Rs  =  Server.CreateObject("ADODB.Recordset")
    Rs.Open SQL, Dbcon, 1, 1

    %>
    <BODY CLASS="BODY_01" SCROLL="no" style="margin:0;border:0;" onload="resize_scroll(90);" onresize="resize_scroll(90);">
    
    <!-- HEAD 부분 시작 --> 
    <TABLE CLASS=Sch_Header id=d2>
	  <TR><TD>
			  
		<table cellspacing="1" cellpadding=0 width="100%">
		<tr><form name=form action="sys_appro.asp" method="post" >
			<Td width=150><select name=key>
				<option value="subject" <% if key="subject" then %>selected <% end if %>>문서제목</option>	
				<option value="dname"   <% if key="dname" then %>selected <% end if %>>문서명</option>	
				<option value="dcode"   <% if key="dcode" then %>selected <% end if %>>문서코드</option>
				<option value="makeuser" <% if key="makeuser" then %>selected <% end if %>>기안자(사번)</option>
				<option value="makedate" <% if key="makedate" then %>selected <% end if %>>작성일(ex.2005-12-01)</option>
				<option value="apuser" <% if key="apuser" then %>selected <% end if %>>현결재자</option>
				</select>
				</td>
            <Td width=100><SELECT name="MDB"  style="width:130px;">
                            <OPTION VALUE="syerp"    <% IF UCASE(MDB) = "MAIL" THEN %>SELECTED<% END IF %>>삼영화성</OPTION>
                            </SELECT></Td>	
            <Td width="60"><select name="chk">
            <option value="A" <% if chk="A" or chk="" then %>selected <% end if %>>완료</option>
            <option value="B" <% if chk="B" then %>selected <% end if %>>진행중</option>
            <option value="C" <% if chk="C" then %>selected <% end if %>>반려</option></select></Td>                            			
			<td width="140" align="center"><input type="input" name="searchtext"  size="20"  class="input" value="<%=searchtext%>"></td>
			<td><input type="submit"  value=" 검 색 " class="btn"></td></form>

		</tr>
		</table>
	  </TD></TR>
    </TABLE>
    <!--  HEAD 부분 끝 --> 

    <!-- CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
        <TR><FORM NAME="check" METHOD="post">
            <TD ALIGN="CENTER" class="TBL_RW_00" WIDTH="6%">SEQ</td>
            <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="7%">CODE</td>
		    <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="15%">문서명</td>
		    <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="34%">문서제목</td>
		    <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="5%">기안자</td>
            <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="5%">상태</td>
            <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="8%">작성일</td>
            <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="5%">라인</td>
            <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="5%">현결</td>
		    <TD ALIGN="CENTER" class="TBL_RW_01" WIDTH="10%">관리</td>
        </tr>
        </table>
        <DIV CLASS="scr_01" id=scr_01>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
	    <%
        While Not Rs.EOF  
            
        nVcnt    =  nVcnt  +  1
        
	    makename    =   Rs("makename")
	    apname      =   Rs("apname")

	    stat = Rs("stat")

	    IF stat=10 then
	    stat="완료"
	    elseif stat=12 then
	    stat="반려"
	    else
	    stat="진행"
	    end if

	    %>
        <TR><TD <%=css1%> ALIGN="CENTER" WIDTH="6%" CLASS="TBL_DRW_00" style="border-left:none;"><%=Rs("idx")%></td>
            <TD <%=css1%> ALIGN="CENTER" WIDTH="7%" CLASS="TBL_DRW_01"><%=Rs("dcode")%></td>
            <TD <%=css1%> ALIGN="CENTER" WIDTH="15%" CLASS="TBL_DRW_01"><%=Rs("dname")%>&nbsp;</td>
            <TD <%=css1%> ALIGN="LEFT" WIDTH="34%" CLASS="TBL_DRW_03" onclick="centerWindow('../gware/approval/ap_win.asp?dcode=<%=Rs("dcode")%>&docm_idx=<%=Rs("idx")%>','name','800','600','yes');" style="cursor:hand;"><%=Rs("subject")%>&nbsp;</td>
            <TD <%=css1%> ALIGN="CENTER" WIDTH="5%" CLASS="TBL_DRW_01"><%=makename%>&nbsp;</td>
            <TD <%=css1%> ALIGN="CENTER" WIDTH="5%" CLASS="TBL_DRW_01"><%=stat%>&nbsp;</td>
            <TD <%=css1%> ALIGN="CENTER" WIDTH="8%" CLASS="TBL_DRW_01"><%=Rs("makedate")%>&nbsp;</td>
            <TD <%=css1%> ALIGN="CENTER" WIDTH="5%" CLASS="TBL_DRW_01"><%=Rs("apmax")%>&nbsp;</td>
            <TD <%=css1%> ALIGN="CENTER" WIDTH="5%" CLASS="TBL_DRW_01"><%=apname%>&nbsp;</td>
		    <TD <%=css1%> ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><Input type="button" value="라인" class="btn" onclick="centerWindow('sys_appro_wrt.asp?dcode=<%=Rs("dcode")%>&docm_idx=<%=Rs("idx")%>&index=<%=Rs("idx")%>&stat=<%=Rs("stat")%>&mail_db=<%=mail_db %>','sys_appro_wrt','500','400','no');"> <Input type="button" value="삭제" class="btn" onclick="javascript:appro_delete('<%=Rs("idx")%>','<%=mail_db%>');"></td>
        </tr>
	    <%
        Rs.MoveNext
        Wend

        Rs.Close
        Set Rs = Nothing
	    %>        
        </TABLE>
        </DIV>
        </DIV>
        <!-- CONTENT 부분 끝 -->
        
        
	<!-- PAGING OR 합계,집계 부분 시작 -->    
	<div class="layer_03" style="border-top:none;">
	<table border="0" cellpadding="0" cellspacing="0" width="980" style="table-layout:fixed;">
		<tr><td <%=css1%> class="tbl_rw_01" width="62%" style="padding-left:10;text-align:left;"><%call gotoPageHTML(GotoPage, Pagecount)%></td>
			<td <%=css1%> class="tbl_rw_01" width="38%" style="text-align:right;padding-right:10;" >레코드 건수 : 총 <%=nvcnt%>건</td>
		</tr>
	</table>
	</div>
    <!-- PAGING OR 합계,집계 부분 끝 -->         
        

    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->
    <script>
    ////////////////////////////////////////////////////////////////////////////////////////
    //							전자결재목록 삭제 스크립트(관리자)  		              //
    ////////////////////////////////////////////////////////////////////////////////////////

    function appro_delete(idx){

		    if (confirm("해당 결재문서를 정말로 삭제하시겠습니까?\n\nDOCM,FORMDATA,APLINE,DISTLINE,ADDFILE 모두 삭제됩니다\n\n삭제시 복구되지 않습니다.")) {
		    window.open('sys_appro_prc.asp?method=delete&docm_idx='+idx,'prc_win','width=1,height=1,left=3000,top=3000,scrolling=no');
				    }
    }
    </script>
    
    
    
    <%
    
    Sub gotoPageHTML(page, Pagecount)
    
    Dim blockpage, i
    blockpage=Int((page-1)/10)*10+1

    
    url = "&key="& key &"&MDB="& mdb &"&chk="& chk &"&searchtext="& searchtext

    response.Write "<table border=0 cellpadding=1 cellspacing=0><tr>"

    '********** 이전 10 개 구문 시작 **********
    if blockPage = 1 Then
        Response.Write "<td width=15 align=center><img src='"& img_path &"arrow_allprev_off.gif' border=0 title='이전10개'></td>"
    Else
        Response.Write"<td width=15 align=center><a href='?gotopage="& blockPage-10 & url &"' title='이전10개'>"
        response.Write "<img src='"& img_path &"arrow_allprev.gif' border=0></a></td>"
    End If
    '********** 이전 10 개 구문 끝**********

    i=1
    Do Until i > 10 or blockpage > Pagecount
    If  blockpage=int(page) Then
        response.write "<td width=15 align=center>"
        response.Write "<div style='border:1px solid #6593CF;background-color:B5D2F7;width:18;height:18;font-weight:bold;font-family:arial;'>"
        response.Write blockpage & "</div></td>"
    Else 
        response.write "<td width=15 align=center ><a href='?gotopage="& blockpage & url &"'>"
        response.Write "<span style='font-family:arial;'>" & blockpage & "</span></a></td>" 
    End If

    blockpage=blockpage+1
    i = i + 1
    Loop

    '********** 다음 10 개 구문 시작**********
    if blockpage > Pagecount Then
        response.Write "<td width=15 align=center><img src='"& img_path &"arrow_allnext_off.gif' border=0 title='다음 10개'></a></td>"
    Else
        response.write "<td width=15 align=center>"
        response.Write "<a href='?gotopage=" & blockpage & url &"' title='다음 10개'><img src='"& img_path &"arrow_allnext.gif' border=0></a></td>"
        End If
    '********** 다음 10 개 구문 끝**********
    
    response.Write "</tr></table>"
    
    End Sub

    
    %>