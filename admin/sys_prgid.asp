	<!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램 : 단위업무관리                                			        --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 5월 23일                                                 --> 
    <!-- 내    용 : 프로그램 명세서, 단위 프로그램의 등록/조회/인쇄/엑셀/결재 주소  --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    //파일명 분리 함수 추가 
    Function filename_conv(str)
       IF InStr(str, "/") > 0 THEN
	   filename_conv = Mid(str, InStrRev(str, "/") + 1)
	   ELSE
	   filename_conv = "&nbsp;"  
	   END IF         
    END Function
    
    PRG_TYPE		= request("PRG_TYPE")
	PRG_GUBUN1		= request("PRG_GUBUN1")
    PRG_GUBUN		= request("PRG_GUBUN")
    searchstring	= request("searchstring")
    searchkey		= request("searchkey")
    If PRG_GUBUN    = "" then PRG_GUBUN="F0001"
    IF PRG_GUBUN1   = "" then PRG_GUBUN1 = "WA01"
    
    SELECT CASE searchkey 
    CASE "PRG_filename" 
       Sub_SQL = " PRG_address  Like '%"& searchstring &"%' or PRG_address1 Like '%"& searchstring &"%' or PRG_print Like '%"& searchstring &"%' or PRG_excel  Like '%"& searchstring &"%' and "
    CASE "PRG_name"
       Sub_SQL = " PRG_name Like N'%"& searchstring &"%' and "
    CASE ""
       Sub_SQL = ""
    END SELECT 
	
	IF PRG_GUBUN1 = "WA1" or PRG_GUBUN1 = "" THEN
	   Sub_SQL1 = " "
	ELSE
	   Sub_SQL1 = " AND PRG_GUBUN1 ='"& PRG_GUBUN1 &"' "
	END IF


	IF PRG_TYPE <> "" THEN
	   Sub_SQL3 = " AND PRG_TYPE ='"& PRG_TYPE &"' "
	ELSE
	   Sub_SQL3 = ""
	END IF

	SQL = " SELECT PRG_id, PRG_name, PRG_TableID, PRG_exp, PRG_AUTH, IsNULL(PRG_address,' ') AS PRG_address, IsNULL(PRG_address1,' ') AS PRG_address1,PRG_Type,PRG_GUBUN1,Gubun_name, IsNULL(PRG_print,' ') AS PRG_print, IsNULL(PRG_excel,' ') AS PRG_excel, IsNULL(PRG_appro,' ') AS PRG_appro"
	SQL = SQL & " FROM PRGID  "
	SQL = SQL & " Left JOIN GUBUN1 On Left(PRG_Gubun1,2)=GUBUN_CODE and right(PRG_Gubun1,2)=GUBUN_seq "
	SQL = SQL & " WHERE " & Sub_SQL &" PRG_GUBUN='"& PRG_GUBUN &"' "& Sub_SQL1 &" "& Sub_SQL3 &" order by PRG_name"
    Set Rs = syscon.Execute(Sql)
    
    'response.Write SQL
    %>
    <script language='javascript' src='/INC/JS/rowspan.js'></script>
    <script language='javascript'>
	    function init()
	    {
		    cellMergeChk(document.all.dataList, 0, 2);		//첫번째 td 처리
    //		cellMergeChk(document.all.dataList, 1, 1);		//두번째 td 처리
	    }
    </script>    
    <BODY CLASS="BODY_01" SCROLL="no"  style="margin:0;" onload="init();resize_scroll(90);" onresize="resize_scroll(90);">	

   
    <!-- HEAD 부분 시작 -->
    <TABLE CLASS=Sch_Header id=d2>
      <TR><TD>
          <table cellpadding=2 cellspacing=0 border=0>
          <tr><form name=form action="sys_prgid.asp" method="post"><input type=hidden name=prg_gubun value=<%=prg_gubun%>>
              <td><select onChange="if(this.selectedIndex!=0) self.location=this.options[this.selectedIndex].value">
				  <option >--------------</option>
					<%
					Set Rsgroup = syscon.Execute("Select distinct fac_cd,fac_nm From  ADMIN_FAC_GUBUN order by fac_nm")

					Do until Rsgroup.EOF
					if trim(PRG_gubun) = trim(Rsgroup("fac_cd")) then
						%>
						<option value="sys_prgid.asp?prg_gubun=<%=Rsgroup("fac_cd")%>" selected style="color:ffffff;background-color:9FB4DC;"><%=Rsgroup("fac_nm")%></option>
						<% 
					else
						%>
						<option value="sys_prgid.asp?prg_gubun=<%=Rsgroup("fac_cd")%>"><%=Rsgroup("fac_nm")%></option>
						<%
					end if
					Rsgroup.Movenext
					Loop

					Rsgroup.close
					Set Rsgroup = Nothing
					%>
					</select></TD>
				<Td><select name="PRG_Type">
					<option value="0" <% IF PRG_Type="0" THEN%> SELECTED <% END IF %>>조회</option>
					<option value="1" <% IF PRG_Type="1" THEN%> SELECTED <% END IF %>>등록</option>
					<option value="2" <% IF PRG_Type="2" THEN%> SELECTED <% END IF %>>산출</option>
					<option value="3" <% IF PRG_Type="3" THEN%> SELECTED <% END IF %>>출력</option>
					<option value="4" <% IF PRG_Type="4" THEN%> SELECTED <% END IF %>>엑셀</option>
					<option value="5" <% IF PRG_Type="5" THEN%> SELECTED <% END IF %>>결재</option>
					</select></td>
				<Td><% CALL FAC_slt1("","PRG_GUBUN1","WA",""& PRG_GUBUN1 &"") %></td>
				<td><input type=radio name=searchkey value=PRG_name <% IF searchkey="PRG_name" or searchkey="" THEN %> checked <% END IF %> > 프로그램명 <input type=radio name=searchkey value=PRG_filename <% IF searchkey="PRG_filename" THEN %> checked <% END IF %>> 파일명</td>
				<Td><input type=input name=searchstring  size=17  class=input value="<%=searchstring%>"></Td>
				<td width=40 style="padding-left:5px;"> <input type=submit  value="search" class=btn></td></form>
				
				<td><input TYPE=button CLASS=btn name="Submit" value="삭제" onclick="javascript:ck()" onMouseOver="this.style.cursor = 'hand'"></td>
			</tr>
			</table>
				
      </TD></TR>
    </TABLE>
    <!-- HEAD 부분 끝 -->

    <!-- CONTENT 부분 시작 -->
		<DIV class="layer_01" align=left>
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
			<TR>	
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="05%">SEQ</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="05%">PID</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="05%">구분</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="20%">프로그램명</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="05%">USE</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">조회</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">등록</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">출력</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">엑셀</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">결재</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="05%">HIT</td>
			</TR>
			</TABLE>
			<DIV CLASS="scr_01" id=scr_01>
				<TABLE BORDER="0" id='dataList' CELLPADDING="0" CELLSPACING="0" WIDTH="980" style="TABLE-LAYOUT: fixed;">
				<form name="check" method="post" action="sys_prg_Del.asp">	
				<%
				VCNT=0
				While Not Rs.EOF
				
				PRG_id		=  Rs("PRG_id") 
				PRG_name	=  Rs("PRG_name") 
				PRG_TableID =  Rs("PRG_TableID") 
				PRG_exp		=  Rs("PRG_exp") 
				PRG_AUTH	=  Rs("PRG_AUTH") 
			
				PRG_address =  filename_conv(Rs("PRG_address"))	: IF Rs("PRG_address") <> ""  THEN sPRG_Address  = replace(Rs("PRG_address"),"/","\")   END IF 
				PRG_address1=  filename_conv(Rs("PRG_address1") )
				
				PRG_print	=  filename_conv(Rs("PRG_print"))
				PRG_excel   =  filename_conv(Rs("PRG_excel"))
				PRG_appro   =  filename_conv(Rs("PRG_appro"))
				PRG_Type    =  Rs("PRG_Type")
				PRG_GUBUN1  =  trim(Rs("PRG_GUBUN1"))
				Gubun_name  =  trim(Rs("Gubun_name")) : If isNULL(Gubun_name) THEN Gubun_name="&nbsp; " END IF
				

				'subSQL = "select count(*) AS hitcnt from sys_TRlog where tr_prgid="& PRG_id
				'Set subRS = syscon.execute(SubSQL)

				'IF Not subRS.EOF then
				'   hitcnt = subRS("hitcnt")
				'END IF
				
				'subRS.close : Set subRS = nothing

				%>            
				<TR><TD ALIGN="CENTER" WIDTH="05%"  CLASS="TBL_DRW_00"><input type="checkbox" name="del" value = "<%=PRG_id%>"></TD>
					<TD ALIGN="CENTER" WIDTH="05%"  CLASS="TBL_DRW_01"><a href="fso_view/vc.asp?FileName=\<%=sPRG_Address%>" target="_blank"><%=(PRG_id)%></TD>	
					<TD ALIGN="CENTER" WIDTH="05%"  CLASS="TBL_DRW_01"><%=Gubun_name%></TD>	   	   	
					<TD ALIGN="LEFT"   WIDTH="20%"  CLASS="TBL_DRW_03" onclick="javascript:centerWindow('sys_prgid_wrt.asp?prg_id=<%=prg_id%>&sLT=조회','sys_prg_id','700','700','n');" style="cursor:hand;"><span class=blue_normal><%=Server.HTMLEncode(PRG_name)%></span></TD>	   
					<td ALIGN="CENTER" WIDTH="5%" CLASS="TBL_DRW_01"><%=PRG_AUTH%></td>	            
					<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01" title="<%=RS("PRG_address")%>" ONCLICK="javascript:centerWindow('sys_ver_cont1.asp?prg_id=<%=prg_id%>&ps_prg_gubun=0&ADDR=PRG_address','sys_prg_win','600','500','no');"  style="cursor:hand;"><span class=Link_class><%=(PRG_address)%></span></a></TD>	            
					<td ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01" title="<%=RS("PRG_address1")%>" ONCLICK="javascript:centerWindow('sys_ver_cont1.asp?prg_id=<%=prg_id%>&ps_prg_gubun=1&ADDR=PRG_address1','sys_prg_win','600','500','no');"  style="cursor:hand;"><%=(PRG_address1)%></td>
					<td ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01" title="<%=RS("PRG_print")%>" ONCLICK="javascript:centerWindow('sys_ver_cont1.asp?prg_id=<%=prg_id%>&ps_prg_gubun=3&ADDR=PRG_print','sys_prg_win','600','500','no');"  style="cursor:hand;"><%=(PRG_print)%></td>
					<td ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01" title="<%=RS("PRG_excel")%>" ONCLICK="javascript:centerWindow('sys_ver_cont1.asp?prg_id=<%=prg_id%>&ps_prg_gubun=4&ADDR=PRG_excel','sys_prg_win','600','500','no');"  style="cursor:hand;"><%=(PRG_excel)%></td>
					<td ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01" title="<%=RS("PRG_appro")%>" ONCLICK="javascript:centerWindow('sys_ver_cont1.asp?prg_id=<%=prg_id%>&ps_prg_gubun=5&ADDR=PRG_appro','sys_prg_win','600','500','no');"  style="cursor:hand;"><%=(PRG_appro)%></td> 
					<TD ALIGN="CENTER" WIDTH="05%"  CLASS="TBL_DRW_01"><%=hitcnt%></TD>	   	                   
				</TR>
				<%
				Rs.MoveNext
				Vcnt=Vcnt+1
				Wend

				Rs.Close
				Set Rs = Nothing
				%> </form>           
				</TABLE>
			</DIV>
		</DIV> 

    <!-- CONTENT 부분 끝 -->

    <!-- PAGING OR 합계,집계 부분 시작 -->    
    <DIV class="layer_03">
      <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
      <TR><TD CLASS="TBL_RW_00" style="padding-left:10px;text-align:left;">레코드건수 : <%=VCNT%> 건</TD></TR>
    </TABLE>
    </DIV>
    <!-- PAGING OR 합계,집계 부분 끝 -->  
	
    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->

    <script language="javascript">
    
    function ck() 
	{
    if (confirm("선택하신 데이타를 삭제하시겠습니까?\n\n사용자메뉴구성(Mlevel), 사용자별 권한(menu_duty), 단위업무(PRgID)   \n\n모두가 삭제됩니다. ")) {
    //		location = 'order_list.asp?mode=del_all&countn='+nu;
        document.check.submit();
        return;
	    }
    else {
	    return;
       }   
    }
    </script>
