	<!--[if !mso]>
	<style>
	v\:* {behavior:url(#default#VML);}
	o\:* {behavior:url(#default#VML);}
	.shape {behavior:url(#default#VML);}
	</style>
	<![endif]-->    
	<!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 금형 Tryout Report 등록                           			--> 
    <!-- 작성자: 문성원(050407)                                                     --> 
    <!-- 작성일자 : 2006년 7월 05일                                                 --> 
    <!-- 내용: 금형 Tryout Report 등록                                              --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->

    <!--#include Virtual = "/INC/CHART/DA_BarChart.asp"  -->
    <%
    stype = REQUEST("stype")
	IF stype = "" THEN stype = "A" END IF

	SELECT CASE stype
	CASE "A" : S1 = "SELECTED"
	CASE "B" : S2 = "SELECTED"
	CASE "C" : S3 = "SELECTED"
	CASE "D" : S4 = "SELECTED"
	CASE "E" : S5 = "SELECTED"
	CASE "F" : S6 = "SELECTED"
	CASE "G" : S7 = "SELECTED"
	END SELECT
    %>
    <script type="text/javascript" src="/inc/js/tabpane.js"></script>
    <link type="text/css" rel="StyleSheet" href="/inc/css/tab.css">

	<% IF stype = "A" THEN  '전체로그별 %>
	<BODY CLASS="BODY_01" SCROLL="no" onload="resize_scroll(90);" onresize="resize_scroll(90);">
	
		<TABLE CLASS=Sch_Header id=d2>
		<TR><TD>
			<table cellpadding=2 cellspacing=0 border=0>
			<TR><form name="form" method="post" action="sys_trlog.asp">
				<input type="hidden" name="stype" value="A" >
				<TD>통계유형</TD>
				<TD><select onChange="if(this.selectedIndex!=0) self.location=this.options[this.selectedIndex].value">
					<option >-------</option>
					<option value="sys_trlog.asp?stype=A" <%=S1%>>로그검색</option>
					<option value="sys_trlog.asp?stype=B" <%=S2%>>사용자별통계</option>
					<option value="sys_trlog.asp?stype=C" <%=S3%>>월별통계</option>
					<option value="sys_trlog.asp?stype=D" <%=S4%>>일별통계</option>
					<option value="sys_trlog.asp?stype=E" <%=S5%>>IP별통계</option>
					<option value="sys_trlog.asp?stype=F" <%=S6%>>OS별통계</option>
					<option value="sys_trlog.asp?stype=G" <%=S7%>>Browser별통계</option>
					</select>
				</TD>
				<TD><% CALL date_frm("B","-","기간","tr_sdate","tr_edate",""& Request("tr_sdate") &"",""& Request("tr_edate") &"") %></TD>
				<TD><INPUT TYPE="Submit" VALUE="검색" CLASS="btn"></TD></form>
			</TR>
			</TABLE>
		</TD></TR>
		</TABLE>

		<DIV class="layer_01" >
		<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
		<TR><TD CLASS="TBL_RW_00" WIDTH="05%">SEQ</TD>
			<TD CLASS="TBL_RW_01" WIDTH="20%">Program Name</TD>
			<TD CLASS="TBL_RW_01" WIDTH="10%">User</TD>
			<TD CLASS="TBL_RW_01" WIDTH="25%">Date</TD>
			<TD CLASS="TBL_RW_01" WIDTH="10%">OS</TD>
			<TD CLASS="TBL_RW_01" WIDTH="15%">Browser</TD>
			<TD CLASS="TBL_RW_01" WIDTH="15%">Ip Address</TD>
		</TR>
		</TABLE>
		
		<DIV CLASS="scr_01" id="scr_01">
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980" BGCOLOR=666666 style="TABLE-LAYOUT: fixed;">
			<%
			IF Request("tr_sdate") <> "" THEN	
				tr_sdate  = Request("tr_sdate")	  
			ELSE 
				IF month(now()) < 10	THEN m = "0"&month(now())	 ELSE m =  month(now())	  END IF
				IF Day(now()) < 10	THEN d = "0"&Day(now())	 ELSE d =  day(now())	  END IF
				tr_sdate  = Year(now())&"-"&m&"-"&d
			END IF

			IF Request("tr_edate") <> "" THEN 
				tr_edate  = Request("tr_edate")	  
			ELSE 
				IF month(now()) < 10	THEN m = "0"&month(now())	 ELSE m =  month(now())	  END IF
				IF Day(now()) < 10	THEN d = "0"&Day(now())	 ELSE d =  day(now())	  END IF
				tr_edate  = Year(now())&"-"&m&"-"&d
			END IF

			SQL = "SELECT *,PRG_NAME,Sung FROM SYS_TRLOG "
			SQL = SQL &" LEFT JOIN PRGID ON PRG_ID = TR_PRGID "
			SQL = SQL &" LEFT JOIN PROFILE ON id = TR_sabun WHERE iduse <> 'N' AND "
			SQL = SQL &" TR_DATE between '"& tr_sdate &" 00:00:00.000' and '"& tr_edate &" 23:59:00.000' ORDER BY TR_IDX DESC"
			SET RS = Syscon.execute(SQL)

			Nvcnt = 1
			WHILE NOT RS.EOF
				TR_idx		= RS("TR_idx") 
				TR_prgid	= RS("TR_prgid")  
				TR_sabun	= RS("TR_sabun")  
				TR_date		= RS("TR_date") 
				TR_ip		= Trim(RS("TR_ip"))
				IF  TR_ip = "210.219.199.10" THEN 
					TR_IP ="<font color='blue'>내부망(210.219.199.10)</font>" 
				ELSEIF Left(TR_IP,10) = "211.240.19" THEN
					TR_IP ="<font color='#339900'>삼성화학 ("& Trim(RS("TR_ip")) &")</font>" 
				ELSEIF left(TR_IP,7) = "218.151" THEN
					TR_IP ="<font color='#FF00CC'>KNX망 ("& Trim(RS("TR_ip")) &")</font>" 
				ELSEIF left(TR_IP,10) = "211.112.53" THEN
					TR_IP ="<font color='#FF0066'>진례 ("& Trim(RS("TR_ip")) &")</font>" 
				ELSE
				END IF
				TR_os		= RS("TR_os")  
				TR_browser	= RS("TR_browser") 
				PRG_NAME	= RS("PRG_NAME") 
				Sung	= RS("Sung") 
			%> 
			<TR><TD CLASS="TBL_DRW_01" WIDTH="05%" ALIGN="center"><%=Nvcnt%></TD>
				<TD CLASS="TBL_DRW_03" WIDTH="20%"><%=PRG_NAME%></TD>
				<TD CLASS="TBL_DRW_01" WIDTH="10%" ALIGN="center"><%=Sung%></TD>
				<TD CLASS="TBL_DRW_01" WIDTH="25%"><%=TR_date%></TD>
				<TD CLASS="TBL_DRW_01" WIDTH="10%" ALIGN="center"><%=TR_os%></TD>
				<TD CLASS="TBL_DRW_01" WIDTH="15%" ALIGN="center"><%=TR_browser%></TD>
				<TD CLASS="TBL_DRW_01" WIDTH="15%"><%=TR_ip%></TD>
			</TR>
			<%
			Nvcnt = Nvcnt + 1
			RS.Movenext
			WEND
			%>
			</TABLE>
		</DIV>
		</DIV>

		<!-- PAGING OR 합계,집계 부분 시작 -->    
		<DIV class="layer_03">
		  <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%"  style="TABLE-LAYOUT: fixed;" >
		  <TR><TD CLASS="TBL_RW_00" style="padding-left:10px;text-align:left;">레코드건수 : <%=Nvcnt-1%> 건</TD></TR>
		</TABLE>
		</DIV>
		<!-- PAGING OR 합계,집계 부분 끝 --> 

	<% ELSEIF stype = "B" THEN  '사용자별%>

		<BODY CLASS="BODY_01" SCROLL="no" >	
		<script language="JavaScript" src="/inc/js/DecChart.js"></script>

		<TABLE CLASS=Sch_Header id=d2>
		<TR><TD>
			<table cellpadding=2 cellspacing=0 border=0>
			<TR><form name="form" method="post" action="sys_trlog.asp">
				<input type="hidden" name="stype" value="B" >
				<TD>통계유형</TD>
				<TD><select onChange="if(this.selectedIndex!=0) self.location=this.options[this.selectedIndex].value">
					<option >-------</option>
					<option value="sys_trlog.asp?stype=A" <%=S1%>>로그검색</option>
					<option value="sys_trlog.asp?stype=B" <%=S2%>>사용자별통계</option>
					<option value="sys_trlog.asp?stype=C" <%=S3%>>월별통계</option>
					<option value="sys_trlog.asp?stype=D" <%=S4%>>일별통계</option>
					<option value="sys_trlog.asp?stype=E" <%=S5%>>IP별통계</option>
					<option value="sys_trlog.asp?stype=F" <%=S6%>>OS별통계</option>
					<option value="sys_trlog.asp?stype=G" <%=S7%>>Browser별통계</option>
					</select>
				</TD>
				<TD><% CALL date_frm("B","-","기간","tr_sdate","tr_edate",""& Request("tr_sdate") &"",""& Request("tr_edate") &"") %></TD>
				<TD><INPUT TYPE="Submit" VALUE="검색" CLASS="btn"></TD></form>
			</TR>
			</TABLE>
		</TD></TR>
		</TABLE>

		<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" style="TABLE-LAYOUT: fixed;">
		<Tr><TD>
			<%
			IF Request("tr_sdate") <> "" THEN	
				tr_sdate  = Request("tr_sdate")	  
			ELSE 
				IF month(now()) < 10	THEN m = "0"&month(now())	 ELSE m =  month(now())	  END IF
				IF Day(now()) < 10	THEN d = "0"&Day(now())	 ELSE d =  day(now())	  END IF
				tr_sdate  = Year(now())&"-"&m&"-"&d
			END IF

			IF Request("tr_edate") <> "" THEN 
				tr_edate  = Request("tr_edate")	  
			ELSE 
				IF month(now()) < 10	THEN m = "0"&month(now())	 ELSE m =  month(now())	  END IF
				IF Day(now()) < 10	THEN d = "0"&Day(now())	 ELSE d =  day(now())	  END IF
				tr_edate  = Year(now())&"-"&m&"-"&d
			END IF

			SQL = "select top 30 sung,count(*) AS cnt from sys_trlog "
			SQL = SQL &" LEFT JOIN PROFILE ON id = TR_sabun WHERE "
			SQL = SQL &" TR_DATE between '"& tr_sdate &" 00:00:00.000' and '"& tr_edate &" 23:59:00.000' group by sung  ORDER BY cnt DESC"
			SET RS = Syscon.execute(SQL)

			Redim PlotData(0,29)
			Redim PlotLable(30)

			WHILE NOT RS.EOF

				PlotData(0,0)	= "aa"
				PlotData(0,i)	= RS("cnt")
				PlotLable(i)	= RS("sung") 

			i = i + 1
			RS.Movenext
			WEND
			
			Call BarChart(PlotData,980,500,PlotLable,"Ylab","<p align=center>상위 30 사용자별 접속횟수</a>")
			%>
			</td></tr>
		</TABLE>

	<% ELSEIF stype = "C" THEN  '월별통계%>

		<BODY CLASS="BODY_01" SCROLL="no" >	
		<script language="JavaScript" src="/inc/js/DecChart.js"></script>

		<TABLE CLASS=Sch_Header id=d2>
		<TR><TD>
			<table cellpadding=2 cellspacing=0 border=0>
			<TR><form name="form" method="post" action="sys_trlog.asp">
				<input type="hidden" name="stype" value="C" >
				<TD>통계유형</TD>
				<TD><select onChange="if(this.selectedIndex!=0) self.location=this.options[this.selectedIndex].value">
					<option >-------</option>
					<option value="sys_trlog.asp?stype=A" <%=S1%>>로그검색</option>
					<option value="sys_trlog.asp?stype=B" <%=S2%>>사용자별통계</option>
					<option value="sys_trlog.asp?stype=C" <%=S3%>>월별통계</option>
					<option value="sys_trlog.asp?stype=D" <%=S4%>>일별통계</option>
					<option value="sys_trlog.asp?stype=E" <%=S5%>>IP별통계</option>
					<option value="sys_trlog.asp?stype=F" <%=S6%>>OS별통계</option>
					<option value="sys_trlog.asp?stype=G" <%=S7%>>Browser별통계</option>
					</select>
				</TD>
				<TD><INPUT TYPE="Submit" VALUE="검색" CLASS="btn"></TD></form>
			</TR>
			</TABLE>
		</TD></TR>
		</TABLE>

		<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" style="TABLE-LAYOUT: fixed;">
		<Tr><TD>
			<%
			function SetDetail(theNum)

				if arrData(theNum,0) = maxCount and maxCount <> 0 then
					arrData(theNum,0) = "<b>" & arrData(theNum,0) & "</b>"
				end if

			end function


			SQL = "select top 30 sung,count(*) AS cnt from sys_trlog "
			SQL = SQL &" LEFT JOIN PROFILE ON id = TR_sabun WHERE "
			SQL = SQL &" TR_DATE between '"& tr_sdate &" 00:00:00.000' and '"& tr_edate &" 23:59:00.000' group by sung  ORDER BY cnt DESC"
			SET RS = Syscon.execute(SQL)

			Redim PlotData(0,29)
			Redim PlotLable(30)

			WHILE NOT RS.EOF

				PlotData(0,0)	= "aa"
				PlotData(0,i)	= RS("cnt")
				PlotLable(i)	= RS("sung") 

			i = i + 1
			RS.Movenext
			WEND
			
			Call BarChart(PlotData,980,500,PlotLable,"Ylab","<p align=center>상위 30 사용자별 접속횟수</a>")
			%>
			</td></tr>
		</TABLE>



	<% ELSEIF stype = "D" THEN  '일별통계%>

	<% ELSEIF stype = "E" THEN  'IP별통계%>

	<% ELSEIF stype = "F" THEN  'OS별통계%>

	<% ELSEIF stype = "G" THEN  '브라우저별통계%>

	<% END IF %>

    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->
