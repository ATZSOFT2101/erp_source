    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램 : 단위업무관리 등록/수정/삭제                   			        --> 
    <!-- 작 성 자 : 오세문                                                          --> 
    <!-- 작성일자 : 2006년 7월 19일                                                 --> 
    <!-- 내    용 : 프로그램 명세서, 단위 프로그램의 등록/조회/인쇄/엑셀/결재 주소  --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
	'######## 공용 Parameter 값 받기 ###############
    PRG_id		=  Request("PRG_id")
    slt			=  trim(Request("slt"))
    method		=  trim(Request("method"))
    action		=  trim(Request("action"))
	param		=  Request("param")

	'######## PRGID 관련 Parameter 값 받기 ###############
    PRG_name	=  trim(Request("PRG_name")) 
    PRG_TableID =  trim(Request("PRG_TableID"))  
    PRG_exp		=  trim(Request("PRG_exp"))  
    PRG_AUTH	=  trim(Request("PRG_AUTH"))  
    PRG_address =  trim(Request("PRG_address"))  
    PRG_address1=  trim(Request("PRG_address1")) 
     
    PRG_print   =  trim(Request("PRG_print"))
    PRG_excel   =  trim(Request("PRG_excel"))
    PRG_appro   =  trim(Request("PRG_appro"))

    PRG_img		=  trim(Request("PRG_img"))  
    PRG_GUBUN	=  trim(Request("PRG_GUBUN")) 
	PRG_GUBUN1  =  trim(Request("PRG_GUBUN1"))
	PRG_Type	=  trim(Request("PRG_Type"))
    PRG_Uid		=  ID

	'######## CRUD 관련 Parameter 값 받기 ###############
	tableid		= Request("tableid")
	tablenm		= Request("tablenm")
	matC		= Request("matC")
	matR		= Request("matR")
	matU		= Request("matU")
	matD		= Request("matD")

	'######## DUTY 관련 Parameter 값 받기 ###############
	mr_pro_id	= Request("mr_pro_id")
	mr_pro_name = Request("mr_pro_name")

	MR_READ		= Request("MR_READ")
	MR_ADD		= Request("MR_ADD")	
	MR_DELETE	= Request("MR_DELETE")
	MR_PRINT	= Request("MR_PRINT")	
	MR_EXCEL	= Request("MR_EXCEL")	
	MR_EDIT		= Request("MR_EDIT")	
	MR_APPRO	= Request("MR_APPRO")  
	
	IF slt = "" THEN S2 = "checked" END IF

    SELECT CASE SLT
    CASE ""		 S1 = "CHECKED" : S2 ="disabled='true'" : S3 ="disabled='true'"
    CASE "등록"  S1 = "CHECKED" : S2 ="disabled='true'" : S3 ="disabled='true'"
    CASE "조회"  S2 = "CHECKED" : S1 ="disabled='true'" 
    CASE "수정"  S2 = "CHECKED" : S1 ="disabled='true'" 
    END SELECT

	'####################  PRGID 처리하는 루틴 시작 #######################
    IF  SLT  <>  ""   THEN
        SELECT CASE  SLT
        CASE  "등록"
			  SQL =" SELECT * From PRGID Where PRG_Address ='"& PRG_Address &"' "
			  Set RS=sysCon.execute(SQL)

			  IF Not RS.EOF THEN
				 Alert_Message "같은 파일명으로는 등록되지 않습니다.","1"
			  ELSE
				 SQL = "insert into PRGID(PRG_name, PRG_TableID, PRG_exp, PRG_AUTH, PRG_istdate, PRG_address, PRG_address1, PRG_img, PRG_Uid,PRG_GUBUN,PRG_print,PRG_excel,PRG_appro,PRG_Type,PRG_GUBUN1) " 
				 SQL = SQL & " VALUES (N'" & PRG_name & "','" & PRG_TableID & "','" & PRG_exp & "','" & PRG_AUTH & "',"
				 SQL = SQL & " getdate(),'"& PRG_address & "','"& PRG_address1 & "','"& PRG_img & "','"& PRG_Uid & "','"& PRG_GUBUN & "','"& PRG_print & "','"& PRG_excel & "','"& PRG_appro & "','"& PRG_Type & "','"& PRG_GUBUN1 &"')"        
	             'response.Write sql			              
				 syscon.Execute(SQL)
			  END IF
			  RS.Close : SET RS=Nothing
        CASE "수정"

		      SET RS = SYSCON.EXECUTE("SELECT * FROM PRGID WHERE PRG_ID= "& PRG_ID &" or PRG_Address ='"& PRG_Address &"'")

              IF  RS.EOF THEN
				  Alert_Message "수정할 수 없습니다.\n\n등록된 자료가 없거나 같은 파일명이 지정되었습니다.","1"
              ELSE
				  SQL = "update PRGID set PRG_name=N'"& PRG_name &"',PRG_TableID='" & PRG_TableID & "',PRG_exp='" & PRG_exp & "',"
                  SQL = SQL & " PRG_AUTH='" & PRG_AUTH & "',PRG_udtdate=getdate(),PRG_address='"& PRG_address &"',"
		          SQL = SQL & " PRG_address1='"& PRG_address1 &"',PRG_img='"& PRG_img &"',PRG_Uid='"& PRG_Uid &"',PRG_GUBUN='"& PRG_GUBUN &"',PRG_print='"& PRG_print &"',PRG_excel='"& PRG_excel &"',PRG_appro='"& PRG_appro &"',PRG_Type='"& PRG_Type &"',PRG_GUBUN1 ='"& PRG_GUBUN1 &"' "
                  SQL = SQL & " where PRG_id =" & PRG_id
                  syscon.Execute(SQL)
			  END IF
			  RS.Close : SET RS=Nothing
        CASE "삭제"
		      SET RS = syscon.Execute("SELECT * FROM PRGID WHERE PRG_ID= "& PRG_ID )

              IF  RS.EOF THEN
				  Alert_Message "삭제할 수 없습니다.\n\n등록된 자료가 없습니다.","1"
              ELSE
				  SYSCON.EXECUTE("DELETE FROM MLEVEL WHERE PGID =" & PRG_ID)
				  SYSCON.EXECUTE("DELETE FROM MENU_DUTY WHERE MR_MLVL_ID =" & PRG_ID)
				  SYSCON.EXECUTE("DELETE FROM PRGID WHERE PRG_ID =" & PRG_ID)
              END IF
			  RS.Close : SET RS=Nothing
        END SELECT
    
	'####################  PRGID 값이 아닌 CRUD나 DUTY 처리하는 루틴 시작 #######################
	ELSE
	
		IF METHOD="CRUD" THEN

		   SELECT CASE ACTION
		   CASE "INST"

				SQL="SELECT COUNT(*) FROM COMCRUD WHERE PGMID ='"& PRG_ID &"' AND  TABLEID='"& tableid &"' "
				SET RS=SYSCON.EXECUTE(SQL)
				IF RS(0)=0 THEN
				   SYSCON.EXECUTE("INSERT INTO COMCRUD (SYSGB, PGMID, TABLEID, MATC, MATR, MATU, MATD) VALUES ('"& PRG_GUBUN1 &"',"& PRG_ID &",'"& tableid &"','"& MATC &"','"& MATR &"','"& MATU &"','"& MATD &"')")
				ELSE
				   Alert_Message "이미 등록된 자료가 있습니다.","1"
			    END IF
			    RS.Close : SET RS=Nothing

		   CASE "DEL"

				SQL="SELECT COUNT(*) FROM COMCRUD WHERE PGMID ='"& PRG_ID &"' AND  TABLEID='"& param &"' "
				SET RS=SYSCON.EXECUTE(SQL)
				IF RS(0) <> 0 THEN
				   SYSCON.EXECUTE("DELETE COMCRUD WHERE PGMID ='"& PRG_ID &"' AND  TABLEID='"& param &"'")
			    ELSE
				   Alert_Message "삭제할 수 없습니다.\n\n등록된 자료가 없습니다.","1"
				END IF
			    RS.Close : SET RS=Nothing

		   END SELECT

		ELSEIF METHOD="DUTY" THEN

		   SELECT CASE ACTION
		   CASE "INST"

				SQL="SELECT COUNT(*) FROM MENU_DUTY WHERE MR_MLVL_ID ="& PRG_ID &" AND  MR_PRO_ID='"& mr_pro_id &"' "
				SET RS=SYSCON.EXECUTE(SQL)
				IF RS(0)=0 THEN
				   SYSCON.EXECUTE("INSERT INTO MENU_DUTY (MR_PRO_ID, MR_MLVL_ID, MR_MLVL_GUBUN, MR_READ, MR_ADD, MR_DELETE, MR_PRINT, MR_EXCEL, MR_EDIT, MR_APPRO) VALUES ('"& mr_pro_id &"',"& PRG_ID &",'1','"& MR_READ &"','"& MR_ADD &"','"& MR_DELETE &"','"& MR_PRINT &"','"& MR_EXCEL &"', '"& MR_EDIT &"','"& MR_APPRO &"')")
				ELSE
				   SYSCON.EXECUTE("DELETE MENU_DUTY WHERE MR_MLVL_ID ="& PRG_ID &" AND  MR_PRO_ID='"& mr_pro_id &"'")
				   SYSCON.EXECUTE("INSERT INTO MENU_DUTY (MR_PRO_ID, MR_MLVL_ID, MR_MLVL_GUBUN, MR_READ, MR_ADD, MR_DELETE, MR_PRINT, MR_EXCEL, MR_EDIT, MR_APPRO) VALUES ('"& mr_pro_id &"',"& PRG_ID &",'1','"& MR_READ &"','"& MR_ADD &"','"& MR_DELETE &"','"& MR_PRINT &"','"& MR_EXCEL &"', '"& MR_EDIT &"','"& MR_APPRO &"')")
			    END IF
			    RS.Close : SET RS=Nothing

		   CASE "DEL"

				SQL="SELECT COUNT(*) FROM MENU_DUTY WHERE MR_ID ="& param 
				SET RS=SYSCON.EXECUTE(SQL)
				IF RS(0) <> 0 THEN
				   SYSCON.EXECUTE("DELETE MENU_DUTY WHERE MR_ID ="& param )
			    ELSE
				   Alert_Message "삭제할 수 없습니다.\n\n등록된 자료가 없습니다.","1"
				END IF
			    RS.Close : SET RS=Nothing
		   CASE "DELALL"

				SQL="SELECT COUNT(*) FROM MENU_DUTY WHERE MR_MLVL_ID ="& PRG_ID 
				SET RS=SYSCON.EXECUTE(SQL)
				IF RS(0) <> 0 THEN
				   SYSCON.EXECUTE("DELETE MENU_DUTY WHERE MR_MLVL_ID ="& PRG_ID  )
			    ELSE
				   Alert_Message "삭제할 수 없습니다.\n\n등록된 자료가 없습니다.","1"
				END IF
			    RS.Close : SET RS=Nothing

		   END SELECT

		END IF

    END IF
	

	IF  PRG_ID <> "" THEN

		SET RS = SYSCON.EXECUTE("SELECT * FROM PRGID WHERE PRG_ID= "& PRG_ID )
		IF  NOT ( RS.EOF OR RS.BOF ) THEN
			PRG_id		=  Rs("PRG_id") 
			PRG_name	=  Rs("PRG_name") 
			PRG_TableID =  Rs("PRG_TableID") 
			PRG_exp		=  Rs("PRG_exp") 
			PRG_AUTH	=  Rs("PRG_AUTH")	: IF PRG_AUTH="Y" THEN PRG_AUTH_chk1="checked" ELSE PRG_AUTH_chk2="checked" END IF			
			PRG_istdate =  Rs("PRG_istdate") 
			PRG_udtdate =  Rs("PRG_udtdate") 
			PRG_address =  Rs("PRG_address") 
			PRG_address1=  Rs("PRG_address1") 
			PRG_img		=  Rs("PRG_img") 
			PRG_Uid		=  Rs("PRG_Uid") 
			PRG_GUBUN	=  Rs("PRG_GUBUN")  
			PRG_GUBUN1  =  trim(Rs("PRG_GUBUN1"))
			PRG_print   =  Rs("PRG_print")
			PRG_excel   =  Rs("PRG_excel")
			PRG_appro   =  Rs("PRG_appro")
			PRG_Type    =  trim(Rs("PRG_Type"))
		END IF
	ELSE
		PRG_AUTH_chk1="checked"
	END IF
    %>
    <BODY CLASS="POPUP_01" SCROLL="no" onload="document.frm_CRUD.tablenm.readOnly = false;">
    <% CALL Popup_generate("단위업무 등록/수정","정확하게 입력하셔야만 합니다")%>
	
	<!-- 프로그램 기본 정보 등록 시작 -->
	<CENTER>
    <FIELDSET STYLE="width:98%;">
    <LEGEND><form action="sys_prgid_wrt.asp" method="POST" Name="PRGID">
			<input type="hidden" name="PRG_img" value="image/page.gif">
			<Input type="hidden" name="PRG_id" value="<%=prg_id%>" >
			<Input type="hidden" name="PRG_GUBUN" value="F0001" >
			<input type="radio" <%=(S1)%> name="slt" value="등록" >등록
			<input type="radio" <%=(S2)%> name="slt" value="수정" >수정                
			<input type="radio" <%=(S3)%> name="slt" value="삭제" >삭제
    </LEGEND>

    <Div class="div_separate_10"></div>

    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" width="98%">
		<tr><td ID="TBL_Title">프로그램명</td> 
            <td ID="TBL_Data"><input type="text" size="35"  name="prg_name"  class="input" value="<%=prg_name%>"></td> 
            <td ID="TBL_Title">테이블명</td>
            <td ID="TBL_Data"><input type="text" size="35"  name="PRG_TableID" class=input value="<%=PRG_TableID%>"></td>  
	    </tr>
        <tr><td ID="TBL_Title">조회주소</td>
            <td ID="TBL_Data"><input type="text" size="35"  name="PRG_address" class=input value="<%=PRG_address%>"></td>  
            <td ID="TBL_Title">사용유무</td>
            <td ID="TBL_Data"><input type="radio" name="prg_auth" id=prg_auth1 value="Y" <%=PRG_AUTH_chk1%>> 사용
                              <input type="radio" name="prg_auth" id=prg_auth2 value="N" <%=PRG_AUTH_chk2%>> 불가 </td>  
	    </tr>
        <tr><td ID="TBL_Title">등록주소</td>
            <td ID="TBL_Data"><input type="text" size="35"  name="PRG_address1" class=input value="<%=PRG_address1%>"></td>  
            <td ID="TBL_Title">엑셀주소</td>
            <td ID="TBL_Data"><input type="text" size="35"  name="PRG_excel" class=input value="<%=PRG_excel%>"></td>  
	    </tr>
        <tr><td ID="TBL_Title">출력주소</td>
            <td ID="TBL_Data"><input type="text" size="35"  name="PRG_print" class=input value="<%=PRG_print%>"></td>  
            <td ID="TBL_Title">결재주소</td>
            <td ID="TBL_Data"><input type="text" size="35"  name="PRG_appro" class=input value="<%=PRG_appro%>"></td>  
	    </tr>			
        <tr><td ID="TBL_Title">프로그램 설명</td>
            <td ID="TBL_Data"><input type="text" size="35"  name="PRG_exp" class=input value="<%=PRG_exp%>"></td>  
            <td ID="TBL_Title">프로그램 구분</td> 
            <td ID="TBL_Data"><Table border=0 cellpadding=0 cellspacing=0>
								<Tr><Td><select name="PRG_Type">
									<option value="0" <% IF PRG_Type="0" THEN%> SELECTED <% END IF %>>조회</option>
									<option value="1" <% IF PRG_Type="1" THEN%> SELECTED <% END IF %>>등록</option>
									<option value="2" <% IF PRG_Type="2" THEN%> SELECTED <% END IF %>>산출</option>
									<option value="3" <% IF PRG_Type="3" THEN%> SELECTED <% END IF %>>출력</option>
									<option value="4" <% IF PRG_Type="4" THEN%> SELECTED <% END IF %>>엑셀</option>
									<option value="5" <% IF PRG_Type="5" THEN%> SELECTED <% END IF %>>결재</option>
									</select></td>
									<Td><% CALL FAC_slt1("","PRG_GUBUN1","WA",""& PRG_GUBUN1 &"") %></td>
									<td><input type="submit"  value=" 확 인 " class=btn></td>
								</tr>
								</table></td></form>	       
        </tr>    
    </TABLE> 
	
    </FIELDSET>
	<!-- 프로그램 기본 정보 등록 끝 -->

	<!---------------------------- CRUD Matrix 등록 시작 ----------------------------->
    <FIELDSET STYLE="width:98%;">
	<form action="sys_prgid_wrt.asp" method="POST" Name="frm_CRUD">
	<Input Type="hidden" name="PRG_ID" value="<%=PRG_ID%>">
	<Input Type="hidden" name="PRG_GUBUN1" value="<%=PRG_GUBUN1%>">
	<Input Type="hidden" name="method" value="CRUD">
	<Input Type="hidden" name="action" value="INST">

    <LEGEND>CRUD Matrix</LEGEND>
    <Div class="div_separate_5"></div>
    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" width="98%">
		<tr height=25>
			<Td width="20%"><% CALL search_frm("sys03","","tableid","tablenm","searchvalue1","","") %></td> 
			<Td width="70%">
				<input type="checkbox" name="matC" value="1"> C 
				<input type="checkbox" name="matR" value="1"> R
				<input type="checkbox" name="matU" value="1"> U
				<input type="checkbox" name="matD" value="1"> D
			</td>
			<td width="10%" align="right"><input type="submit"  value="추가" class=btn></td></form>
		</tr>
		<%
		prgSQL = " SELECT A.*,B.TABLENM FROM DBO.COMCRUD A "
		prgSQL = prgSQL & " LEFT JOIN SYS_TABLE B ON A.TABLEID= B.TABLEID "
		prgSQL = prgSQL & " WHERE PGMID='"& PRG_ID &"' "
		SET prgRS=syscon.EXECUTE(prgSQL)
		%>
		<tr><TD colspan="3">

				<DIV class="layer_01" style="width:100%;">
				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="550"  style="TABLE-LAYOUT: fixed;" >
				<TR><TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">SEQ</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="20%">테이블명</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="30%">설명</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">C</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">R</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">U</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">D</TD>
				</TR>
				</TABLE>
				  
					<DIV CLASS="scr_01" style="height:100px;">
					<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="550" BGCOLOR=666666 style="TABLE-LAYOUT: fixed;">
					<%
					nVcnt = 1
					Do until prgRS.EOF
							
					tableid = prgRS("tableid")
					tablenm = prgRS("tablenm")
					MatC	= prgRS("MatC") : IF MatC="1" THEN MatC_chk="checked" END IF
					MatR	= prgRS("MatR") : IF MatR="1" THEN MatR_chk="checked" END IF
					MatU	= prgRS("MatU") : IF MatU="1" THEN MatU_chk="checked" END IF
					MatD	= prgRS("MatD") : IF MatD="1" THEN MatD_chk="checked" END IF
					%>
					<TR><TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_00" onclick="javascript:deletedata('<%=prg_id%>','<%=tableid%>','CRUD','DEL');" style="cursor:hand;"><img src="<%=img_path%>Sch_Quit.gif" border=0 /></TD>
						<TD ALIGN="LEFT"   WIDTH="20%" CLASS="TBL_DRW_03"><%=tableid%></TD>	   
						<TD ALIGN="LEFT"   WIDTH="30%" CLASS="TBL_DRW_01"><%=tablenm%></TD>	    	   
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=MatC%>" <%=MatC_chk%>></TD>
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=MatR%>" <%=MatR_chk%>></TD>
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=Matu%>" <%=MatU_chk%>></TD>
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=Matd%>" <%=MatD_chk%>></TD>	    	   
					</TR>       
					<%
					MatC_chk = "" : MatR_chk = "" : MatU_chk = "" : MatD_chk = "" 
					nVcnt = nVcnt + 1
					prgRS.Movenext
					Loop

					prgRS.Close : Set prgRS=nothing
					%>
					</TABLE>
					</DIV>
				</DIV> 

		</TD></TR>
    </TABLE>
	</FIELDSET>
	<!---------------------------- CRUD Matrix 등록 끝 ----------------------------->

	<!---------------------------- DeFault 권한 등록 시작 -------------------------->
    <FIELDSET STYLE="width:98%;">
	<form action="sys_prgid_wrt.asp" method="POST" Name="frm_DUTY">
	<Input Type="hidden" name="PRG_ID" value="<%=PRG_ID%>">
	<Input Type="hidden" name="method" value="DUTY">
	<Input Type="hidden" name="action" value="INST">

    <LEGEND>기본 사용자별 권한</LEGEND>
    <Div class="div_separate_5"></div>
    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" width="98%">
		<tr height=25>
			<td width="20%"><% CALL search_frm("sys02","","mr_pro_id","mr_pro_name","searchvalue2","","") %></td> 
			<Td width="60%">
				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0">
					<tr><td><input type="checkbox" name="mr_read"   value="1" checked>조회</td> 
						<td><input type="checkbox" name="mr_add"    value="1" checked>등록</td> 
						<td><input type="checkbox" name="mr_delete" value="1" checked>삭제</td> 
						<td><input type="checkbox" name="mr_print"  value="1" checked>인쇄</td> 
						<td><input type="checkbox" name="mr_excel"  value="1" checked>엑셀</td>
						<td><input type="checkbox" name="mr_edit"   value="1" checked>수정</td>
						<td><input type="checkbox" name="mr_appro"  value="1" checked>결재</td> 
					</tr>
				</TABLE> 
			</td>
			<td width="20%" align="right"><input type="submit"  value="추가" class=btn><input type="button"  value="전체삭제" class=btn onclick="javascript:deleteall('<%=prg_id%>','DUTY','DELALL');"></td></form>
		</tr>
		<%
		usrSQL = " SELECT Mr_id,sung,mr_pro_id,MR_READ,MR_ADD,MR_DELETE,MR_PRINT,MR_EXCEL,MR_EDIT,MR_APPRO FROM PRGID "
		usrSQL = usrSQL &" inner join MENU_DUTY On Prg_id = mr_mlvl_id "
		usrSQL = usrSQL &" inner Join Profile On mr_pro_id = id "
		usrSQL = usrSQL &" WHERE prg_id='"& PRG_ID &"' "
		SET usrRS=syscon.EXECUTE(usrSQL)
		%>
		<tr><TD colspan="3">

				<DIV class="layer_01" style="width:100%;">
				<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="550"  style="TABLE-LAYOUT: fixed;" >
				<TR><TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="10%">SEQ</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="20%">성명</TD>

					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">조회</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">등록</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">삭제</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">출력</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">엑셀</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">수정</TD>
					<TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">결재</TD>
				</TR>
				</TABLE>
				  
					<DIV CLASS="scr_01" style="height:100px;">
					<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="550" BGCOLOR=666666 style="TABLE-LAYOUT: fixed;">
					<%
					nVcnt = 1
					Do until usrRS.EOF
					
					Mr_id		= usrRs("Mr_id")
					sung		= usrRS("sung")
					mr_pro_id	= usrRS("mr_pro_id")
					MR_READ		= usrRS("MR_READ")		: IF MR_READ	="1"	THEN MR_READ_chk="checked"		END IF
					MR_ADD		= usrRS("MR_ADD")		: IF MR_ADD		="1"	THEN MR_ADD_chk="checked"		END IF
					MR_DELETE	= usrRS("MR_DELETE")	: IF MR_DELETE	="1"	THEN MR_DELETE_chk="checked"	END IF
					MR_PRINT	= usrRS("MR_PRINT")		: IF MR_PRINT	="1"	THEN MR_PRINT_chk="checked"		END IF
					MR_EXCEL	= usrRS("MR_EXCEL")		: IF MR_EXCEL	="1"	THEN MR_EXCEL_chk="checked"		END IF
					MR_EDIT		= usrRS("MR_EDIT")		: IF MR_EDIT	="1"	THEN MR_EDIT_chk="checked"		END IF
					MR_APPRO	= usrRS("MR_APPRO")		: IF MR_APPRO	="1"	THEN MR_APPRO_chk="checked"		END IF
					%>
					<TR><TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_00" onclick="javascript:deletedata('<%=prg_id%>','<%=mr_id%>','DUTY','DEL');" style="cursor:hand;"><img src="<%=img_path%>Sch_Quit.gif" border=0 /></TD>
						<TD ALIGN="LEFT"   WIDTH="20%" CLASS="TBL_DRW_03"><%=sung%> (<%=mr_pro_id%>)</TD>
						
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=MR_READ%>"		<%=MR_READ_chk%>></TD>
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=MR_ADD%>"		<%=MR_ADD_chk%>></TD>
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=MR_DELETE%>"	<%=MR_DELETE_chk%>></TD>
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=MR_PRINT%>"	<%=MR_PRINT_chk%>></TD>
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=MR_EXCEL%>"	<%=MR_EXCEL_chk%>></TD>
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=MR_EDIT%>"		<%=MR_EDIT_chk%>></TD>
						<TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><input type=checkbox value="<%=MR_APPRO%>"	<%=MR_APPRO_chk%>></TD>
						
					</TR>       
					<%
					MR_READ_chk = ""	: MR_ADD_chk = ""	: MR_DELETE_chk = "" : MR_PRINT_chk = ""
					MR_EXCEL_chk = ""	: MR_EDIT_chk = ""	: MR_APPRO_chk = "" 
					nVcnt = nVcnt + 1
					usrRS.Movenext
					Loop

					usrRS.Close : Set usrRS=nothing
					%>
					</TABLE>
					</DIV>
				</DIV> 

		</TD></TR>
    </TABLE>
	</FIELDSET>
	<!---------------------------- DeFault 권한 등록 끝 -------------------------->

    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
    <SCRIPT language="JavaScript">
    <!--
    self.moveTo((screen.availWidth-800)/2,(screen.availHeight-700)/2)
    self.resizeTo(800,700)

    function deletedata(prg_id,param,method,action) {
		location.href="sys_prgid_wrt.asp?prg_id="+ prg_id +"&method="+ method +"&action="+ action +"&param="+ param  ;
    }
    function deleteall(prg_id,method,action) {
		location.href="sys_prgid_wrt.asp?prg_id="+ prg_id +"&method="+ method +"&action="+ action  ;
    }

    //-->
    </SCRIPT>

