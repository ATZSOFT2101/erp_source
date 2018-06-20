    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : ERP 용어 관리                              			        --> 
    <!-- 작성자: 문성원(050407)                                                     --> 
    <!-- 작성일자 : 2006년 7월 20일                                                 --> 
    <!-- 내용: 언어별 ERP 용어 관리													--> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
	searchkey	 = REQUEST("searchkey")
	SEARCHSTRING = TRIM(REQUEST("SEARCHSTRING"))
	W_GUBUN		 = REQUEST("W_GUBUN")
	KEY			 = REQUEST("KEY")
	KEY_VALUE	 = REQUEST("KEY_VALUE")

	SELECT CASE KEY_VALUE
	CASE "ㄱ"   KEY_VALUE1 = "ㄴ" 
	CASE "ㄴ"   KEY_VALUE1 = "ㄷ" 
	CASE "ㄷ"   KEY_VALUE1 = "ㄹ" 
	CASE "ㄹ"   KEY_VALUE1 = "ㅁ" 
	CASE "ㅁ"   KEY_VALUE1 = "ㅂ" 
	CASE "ㅂ"   KEY_VALUE1 = "ㅅ" 
	CASE "ㅅ"   KEY_VALUE1 = "ㅇ" 
	CASE "ㅇ"   KEY_VALUE1 = "ㅈ" 
	CASE "ㅈ"   KEY_VALUE1 = "ㅊ" 
	CASE "ㅊ"   KEY_VALUE1 = "ㅋ" 
	CASE "ㅋ"   KEY_VALUE1 = "ㅌ" 
	CASE "ㅌ"   KEY_VALUE1 = "ㅍ" 
	CASE "ㅍ"   KEY_VALUE1 = "ㅎ" 
	CASE "ㅎ"   KEY_VALUE1 = "힣"
	CASE "A"   KEY_VALUE1 = "B" 
	CASE "B"   KEY_VALUE1 = "C" 
	CASE "C"   KEY_VALUE1 = "D" 
	CASE "D"   KEY_VALUE1 = "E" 
	CASE "E"   KEY_VALUE1 = "F" 
	CASE "F"   KEY_VALUE1 = "G" 
	CASE "G"   KEY_VALUE1 = "H" 
	CASE "H"   KEY_VALUE1 = "I" 
	CASE "I"   KEY_VALUE1 = "J" 
	CASE "J"   KEY_VALUE1 = "K" 
	CASE "K"   KEY_VALUE1 = "L" 
	CASE "L"   KEY_VALUE1 = "M" 
	CASE "M"   KEY_VALUE1 = "N" 
	CASE "N"   KEY_VALUE1 = "O"
	CASE "0"   KEY_VALUE1 = "P" 
	CASE "P"   KEY_VALUE1 = "Q" 
	CASE "Q"   KEY_VALUE1 = "R" 
	CASE "R"   KEY_VALUE1 = "S" 
	CASE "S"   KEY_VALUE1 = "T" 
	CASE "T"   KEY_VALUE1 = "U" 
	CASE "U"   KEY_VALUE1 = "V" 
	CASE "V"   KEY_VALUE1 = "W" 
	CASE "W"   KEY_VALUE1 = "X" 
	CASE "X"   KEY_VALUE1 = "Y" 
	CASE "Y"   KEY_VALUE1 = "Z" 
	CASE "Z"   KEY_VALUE1 = "ZZ" 
	CASE "0"   KEY_VALUE1 = "9" 
	CASE NULL  KEY_VALUE1 = "ALL" 
	END SELECT     


	IF W_GUBUN = "WA1" or W_GUBUN = "" THEN
	   Sub_SQL = " W_GUBUN <> '' "
	ELSE
	   Sub_SQL = " W_GUBUN ='"& W_GUBUN &"' "
	END IF

	IF SEARCHSTRING <> "" THEN
	   SUB_SQL1 = " AND "& SEARCHKEY &" LIKE N'%"& SEARCHSTRING &"%' "
    END IF 

	SELECT CASE KEY
	CASE "ALL"
	   Sub_SQL2 = " "
	CASE "KOR"
	   Sub_SQL2 = " AND W_KWORD >= N'"& KEY_VALUE &"' and W_KWORD < N'"& KEY_VALUE1 &"' "
	CASE "ENG"
	   Sub_SQL2 = " AND W_EWORD >= N'"& KEY_VALUE &"' and W_EWORD < N'"& KEY_VALUE1 &"' "
	CASE "CHN"
	   Sub_SQL2 = " AND W_CWORD >= N'"& KEY_VALUE &"' and W_CWORD < N'"& KEY_VALUE1 &"' "
	CASE "RUS"
	   Sub_SQL2 = " AND W_RWORD >= N'"& KEY_VALUE &"' and W_RWORD < N'"& KEY_VALUE1 &"' "
	END SELECT	

	SQL = " SELECT A.*,Gubun_name FROM SYS_WORD A LEFT JOIN GUBUN1 On gubun_code=Left(W_GUBUN,2) and gubun_seq=RIGHT(W_GUBUN,2) "
	SQL = SQL &" WHERE "& Sub_SQL &" "& SUB_SQL1 &" "& SUB_SQL2 &" order by w_code asc"
    Set Rs = syscon.Execute(Sql)
    
    'response.Write sql
    %>
    <BODY CLASS="BODY_01" SCROLL="no"  onload="resize_scroll(120);" onresize="resize_scroll(120);">
    
    <!-- HEAD 부분 시작 -->
    <TABLE CLASS=Sch_Header id=d2>
      <TR><TD>
          <table cellpadding=2 cellspacing=0 border=0>
          <tr><form name=form action="sys_word.asp" method="post">
			  <Input type="hidden" name="KEY" value="<%=KEY%>">
			  <Input type="hidden" name="KEY_VALUE" value="<%=KEY_VALUE%>">
              <td><% CALL FAC_slt1("","W_GUBUN","WA",""& W_GUBUN &"") %></TD>
			  <td>
			  <input type="radio" name="searchkey" value="W_KWORD" <% IF searchkey="W_KWORD" or searchkey="" THEN %> checked <% END IF %> > 한글 
			  <input type="radio" name="searchkey" value="W_EWORD" <% IF searchkey="W_EWORD" THEN %> checked <% END IF %>> 영문
			  <input type="radio" name="searchkey" value="W_CWORD" <% IF searchkey="W_CWORD" THEN %> checked <% END IF %>> 중문
			  <input type="radio" name="searchkey" value="W_RWORD" <% IF searchkey="W_RWORD" THEN %> checked <% END IF %>> 중문
			  <input type="radio" name="searchkey" value="W_CONTENT" <% IF searchkey="W_CONTENT" THEN %> checked <% END IF %>> 용어설명
			  </td>
			  <Td><input type="input" name="searchstring"  size=17  class=input value="<%=searchstring%>"></Td>
			  <td width=40 style="padding-left:5px;"> <input type="submit" value=" 검색 " class="btn"></td></form>
			  <td><input type="button" value=" 등 록 " onclick="javascript:centerWindow('sys_word_wrt.asp','aaa','500','400','no');" class="btn" />
			      <input type="button" value=" 생 성 " onclick="javascript:centerWindow('sys_word_ex.asp','aaa','10','10','no');" class="btn" />
			  </td>
		  </tr>
		  </table>
				
      </TD></TR>
    </TABLE>
    <!-- HEAD 부분 끝 -->
    <TABLE CELLPADDING="5" CELLSPACING="1" BGCOLOR="666666" WIDTH="980" style="TABLE-LAYOUT: fixed;">
      <TR BGCOLOR="#CCFFFF">
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ALL" OR KEY_VALUE="" THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %>  Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ALL&KEY_VALUE=ALL');">전</a></TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㄱ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㄱ');">ㄱ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㄴ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㄴ');">ㄴ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㄷ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㄷ');">ㄷ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㄹ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㄹ');">ㄹ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅁ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅁ');">ㅁ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅂ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅂ');">ㅂ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅅ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅅ');">ㅅ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅇ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅇ');">ㅇ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅈ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅈ');">ㅈ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅊ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅊ');">ㅊ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅋ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅋ');">ㅋ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅌ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅌ');">ㅌ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅍ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅍ');">ㅍ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="ㅎ"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=KOR&KEY_VALUE=ㅎ');">ㅎ</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="A"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=A');">A</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="B"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=B');">B</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="C"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=C');">C</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="D"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=D');">D</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="E"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=E');">E</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="F"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=F');">F</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="G"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=G');">G</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="H"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=H');">H</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="I"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=I');">I</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="J"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=J');">J</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="K"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=K');">K</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="L"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=L');">L</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="M"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=M');">M</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="N"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=N');">N</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="O"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=O');">O</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="P"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=P');">P</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="Q"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=Q');">Q</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="R"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=R');">R</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="S"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=S');">S</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="T"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=T');">T</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="U"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=U');">U</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="V"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=V');">V</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="W"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=W');">W</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="X"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=X');">X</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="Y"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=Y');">Y</TD>
		  <TD style="cursor:hand;" align="center" <% IF KEY_VALUE="Z"  THEN %> style="background-color:#629AD9;color:#ffffff;font-weight:bold;" <% END IF %> Onclick="javascript:URL_move('sys_word.asp?W_GUBUN=<%=W_GUBUN%>&searchkey=<%=searchkey%>&searchstring=<%=searchstring%>&KEY=ENG&KEY_VALUE=Z');">Z</TD>
	  </TR>
    </TABLE>
	
	 <Div class="div_separate_5"></div>
    <!-- CONTENT 부분 시작 -->
		<DIV class="layer_01" align=left>
			<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
			<TR>	
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="05%">SEQ</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="08%">CODE</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="08%">구분</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="13%">KOREAN</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="13%">ENGLISH</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="13%">CHINESE</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="13%">RUSSIA</td>
				<td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="27%">CONTENT</td>
			</TR>
			</TABLE>
			<DIV CLASS="scr_01" id=scr_01>
				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980" style="TABLE-LAYOUT: fixed;">
				<%
				VCNT=1
				While Not Rs.EOF
				
				W_IDX		=  Rs("W_IDX") 
				W_CODE		=  Rs("W_CODE") 
				W_KWORD		=  Rs("W_KWORD") 
				W_EWORD		=  Rs("W_EWORD") 
				W_CWORD		=  Rs("W_CWORD") 
				W_RWORD		=  Rs("W_RWORD") 
				W_GUBUN		=  Rs("W_GUBUN") 
				W_CONTENT	=  Rs("W_CONTENT") 
				W_WDATE		=  Rs("W_WDATE")  
				W_EDATE		=  Rs("W_EDATE") 
				W_SABUN		=  Rs("W_SABUN") 
				GUBUN_NAME	=  Rs("GUBUN_NAME")  
				%>            
				<TR><TD ALIGN="CENTER" WIDTH="05%"  CLASS="TBL_DRW_00" ><%=VCNT%></TD>
					<TD ALIGN="CENTER" WIDTH="08%"  CLASS="TBL_DRW_01" ><%=W_CODE%></TD>
					<TD ALIGN="CENTER" WIDTH="08%"  CLASS="TBL_DRW_01"><%=GUBUN_NAME%></TD>	   	   	
					<TD ALIGN="LEFT"   WIDTH="13%"  CLASS="TBL_DRW_03" onclick="javascript:centerWindow('sys_word_wrt.asp?W_IDX=<%=W_IDX%>&SLT=MOD','sys_word_wrt','500','540','n');" style="cursor:hand;"><span class=blue_normal><%=W_KWORD%></span></TD>	   
					<td ALIGN="LEFT"   WIDTH="13%"  CLASS="TBL_DRW_01">&nbsp;<%=W_EWORD%></td>	            
					<TD ALIGN="LEFT"   WIDTH="13%"  CLASS="TBL_DRW_01">&nbsp;<%=W_CWORD%></a></TD>	            
					<TD ALIGN="LEFT"   WIDTH="13%"  CLASS="TBL_DRW_01">&nbsp;<%=W_RWORD%></a></TD>	            
					<td ALIGN="LEFT"   WIDTH="30%"  CLASS="TBL_DRW_01">&nbsp;<%=W_CONTENT%></td>                
				</TR>
				<%
				Rs.MoveNext
				Vcnt=Vcnt+1
				Wend

				Rs.Close
				Set Rs = Nothing
				%>          
				</TABLE>
			</DIV>
		</DIV> 

    <!-- CONTENT 부분 끝 -->

    <!-- PAGING OR 합계,집계 부분 시작 -->    
    <DIV class="layer_03">
      <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
      <TR><TD CLASS="TBL_RW_00" class=paging_white_normal style="padding-left:10px;text-align:left;">총 <%=Vcnt-1%> EA</TD></TR>
    </TABLE>
    </DIV>
    <!-- PAGING OR 합계,집계 부분 끝 -->  

	<!--#include Virtual = "/INC/INC_FOOT.ASP" -->