    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 사용자 관리                                          --> 
    <!-- 내용: 그룹웨어 사용자 관리												--> 
    <!-- 작 성 자 : 문성원(050407)                                                --> 
    <!-- 작성일자 : 2006년 6월 14일                                         --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    '************ 설정값 **********************************
    ENT_SCPT = " onkeypress=""return handleEnter(this, event);"" class=""input"" style=""width:80%;font-family:arial;font-size:12px;"" "
    DBL_SCPT = " onkeypress=""return handleEnter(this, event);"" class=""input"" style=""width:73;font-family:arial;font-size:12px;text-align:center;"" "
    SLT_SCPT = " style=""width:80%;font-family:arial;font-size:12px;"" "
   
    
    '************ 검색관련 REQUEST 값 **********************
    nvcnt       =   0
    fac	        =   trim(request("fac"))
    iduse       =   trim(request("iduse"))
    sung        =   trim(request("sung"))
    
    if request("odr") = "" then odr = "sung"    else odr   =   trim(request("odr")) end if
    if request("sdr") = "" then sdr = "asc"     else sdr   =   trim(request("sdr")) end if
    
    '************ 검색관련 쿼리문 시작 ********************
    
    sql = "select a.*,dm_name,ur_id,ur_mailboxsize,ur_mailid,ur_pwd,ur_permitprotocol,ur_curr_mail_size,"
    sql = sql &"  case when csuserid is null then 0 else 1 end as msg_chk,csuserid,csuserpw,csname,"
    sql = sql &"  dept_nm,jik_nm,cstate,isystem_classsystem,idx_classpostion"
    sql = sql & " from profile a "
    sql = sql & " left join domain b on a.dm_id = b.dm_id "
    sql = sql & " left join mc_usersimple on id = csuserid "
    sql = sql & " left join mc_userdetail on id = csuserid_usersimple "
    sql = sql & " left join admin_dept_gubun on dept = dept_cd "
    sql = sql & " left join domainuser d on a.email = d.ur_mailid and a.dm_id = d.dm_id"
    sql = sql & " left join admin_jik_gubun on jik = jik_cd "
    sql = sql & " where iduse = '"& iduse &"' and fac = '"& fac &"' and sung like '%"& sung &"%' order by "& odr &" "& sdr &" " 
  ' response.Write sql
    set rs  =  syscon.execute(sql)
    %>
	<BODY CLASS="BODY_01" SCROLL="no"  style="margin:0;" Onload="resize_scroll(340);" Onresize="resize_scroll(340);" >
	    
    <!-- HEAD 부분 시작 --> 
    <TABLE CLASS="Sch_Header" id=d2>
	    <TR><TD><FORM NAME="search_form" ACTION="sys_user.asp" method="post" >
            <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0">
            <TR><TD>공장구분</TD> 
                <TD><SELECT  NAME="fac">
				<%
				SQLGROUP = " SELECT DISTINCT FAC_CD,FAC_NM FROM ADMIN_FAC_GUBUN ORDER BY FAC_CD ASC"
				SET RSGROUP = SYSCON.EXECUTE(SQLGROUP)
                
				WHILE NOT RSGROUP.EOF
				  
				IF TRIM(FAC) = TRIM(RSGROUP("FAC_CD")) THEN
				   RESPONSE.WRITE   "<OPTION VALUE='"& RSGROUP("FAC_CD") &"' SELECTED>"& RSGROUP("FAC_NM") &"</OPTION>"
				ELSE
				   RESPONSE.WRITE   "<OPTION VALUE='"& RSGROUP("FAC_CD") &"'>"& RSGROUP("FAC_NM") &"</OPTION>"
				END IF
						  
				RSGROUP.MOVENEXT
				WEND

				RSGROUP.CLOSE : SET RSGROUP = NOTHING
				%>
				</SELECT></TD>
				<TD WIDTH="70" align="RIGHT">사용구분</TD>
				<TD><SELECT  NAME="iduse">
				<OPTION VALUE="A" <% if iduse="A" then%>selected<%end if %>>사용</OPTION>
				<OPTION VALUE="N" <% if iduse="N" then%>selected<%end if %>>불가</OPTION>
				</SELECT></TD>
				<TD WIDTH="40" align="RIGHT">성명</TD>
				<TD><INPUT TYPE="TEXT" NAME="sung" SIZE="13" VALUE="<%=sung%>" CLASS=INPUT></TD>
				<TD><INPUT TYPE="SUBMIT" VALUE="검색" CLASS="BTN"></TD></FORM>
			</TR> 
		    </TABLE>
        </TD></TR>
    </TABLE>
    <!--  HEAD 부분 끝 --> 

    <!-- CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"   >
        <TR><FORM NAME="check" METHOD="post">
            <TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="05%" rowspan="2">SEQ</TD>
		    <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%" rowspan="2"><a href="sys_user.asp?odr=sung&sdr=asc&fac=<%=fac%>&iduse=<%=iduse%>&sung<%=sung%>">성명</a></TD>
		    <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="08%" rowspan="2">I.D</TD>
		    
		    <TD ALIGN="CENTER" CLASS="TBL_RW_03" WIDTH="30%" colspan="3">그룹웨어</TD>
		    <TD ALIGN="CENTER" CLASS="TBL_RW_03" WIDTH="30%" colspan="2">메일서버</TD>
		    <TD ALIGN="CENTER" CLASS="TBL_RW_03" WIDTH="15%" colspan="3">메신저</TD>
       </TR>
       <tr>     
            <!-- PROFILE 정보 -->
		    <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="08%"><a href="sys_user.asp?odr=dept&sdr=asc&fac=<%=fac%>&iduse=<%=iduse%>&sung<%=sung%>">부서</a></TD>
		    <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="07%">직급</TD>
		    <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="07%">P.W</TD>
		    
            <!-- MAILSERVER 정보 -->		    
		    <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="15%">Email</TD>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="25%"><a href="sys_user.asp?odr=ur_curr_mail_size&sdr=desc&fac=<%=fac%>&iduse=<%=iduse%>&sung<%=sung%>">메일사용량</a></td>
		    
            <!-- MESSENGER 정보 -->		    
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="04%">USE</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="07%">P.W</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="04%">상태</td>
        </TR>
        </TABLE>

        <DIV CLASS="scr_01" id="scr_01" >
        <TABLE BORDER="0" id="datalist" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
        <%
        nVcnt = 1
        While Not Rs.EOF  
            
	      IDX		        =  RS("IDX")
	      ID		        =  TRIM(RS("ID"))
	      PASSWD	        =  TRIM(RS("PASSWD"))
	      EMAIL		        =  TRIM(RS("EMAIL"))
	      FEMAIL		    =  TRIM(RS("FEMAIL"))
	      SUNG		        =  TRIM(RS("SUNG"))
	      TELE		        =  TRIM(RS("TELE"))
	      TELE1		        =  TRIM(RS("TELE1"))
	      DDATE		        =  TRIM(RS("DDATE"))
	      SEX		        =  TRIM(RS("SEX"))
	      FAC		        =  TRIM(RS("FAC"))
	      DEPT		        =  TRIM(RS("DEPT"))
	      DEPT_nm	        =  TRIM(RS("DEPT_nm"))
	      JIK		        =  TRIM(RS("JIK"))
	      jik_nm            =  TRIM(RS("jik_nm"))
	      GUBUN		        =  TRIM(RS("GUBUN"))
	      WS_GUBUN		    =  TRIM(RS("WS_GUBUN"))
	      IDUSE		        =  TRIM(RS("IDUSE"))
	      COP		        =  TRIM(RS("COP"))
	      MAIL_DB		    =  TRIM(RS("MAIL_DB"))
	      ERP_DB		    =  TRIM(RS("ERP_DB"))
	      FAC_GUBUN		    =  TRIM(RS("FAC_GUBUN"))
	      CHARSET		    =  TRIM(RS("CHARSET"))
	      
	      DM_ID		        =  RS("DM_ID")
	      MS_ID		        =  RS("MS_ID")
	      UR_ID		        =  RS("UR_ID")
	   '   UR_MAILBOXSIZE    =  cdbl(RS("UR_MAILBOXSIZE"))*1000*1000
	      UR_MAILID		    =  RS("UR_MAILID")
	      UR_PWD		    =  trim(RS("UR_PWD"))
	      UR_PERMITPROTOCOL =  RS("UR_PERMITPROTOCOL")
	   '   UR_CURR_MAIL_SIZE =  cdbl(RS("UR_CURR_MAIL_SIZE"))
	      DM_NAME           =  RS("DM_NAME") 
	      csUserid          =  RS("csUserid")  
	      csUserpw          =  RS("csUserpw")    
	      cstate            =  RS("cstate") 
	      iSystem_ClassSystem  =  RS("iSystem_ClassSystem") 
          idx_ClassPostion  =  RS("idx_ClassPostion") 
          msg_chk           =  RS("msg_chk")
	      
	      IF PASSWD <> UR_PWD THEN
	         FLAG_PWD = " Style='background-color:red;'"
	      END IF
	      
 
     '   IF (ur_curr_mail_size/ur_mailboxsize)*100 >= 80 THEN
	 '       FLAG_SIZE = "red"
	 '       Mailsize = formatpercent(ur_curr_mail_size/ur_mailboxsize)
     '   elseIF (ur_curr_mail_size/ur_mailboxsize)*100 >= 50 and (ur_curr_mail_size/ur_mailboxsize)*100 <= 80 THEN
	 '       FLAG_SIZE = "blue"
	 '       Mailsize = formatpercent(ur_curr_mail_size/ur_mailboxsize)
     '   END IF

	      
        %>
        <TR Onclick="show_data('<%=ID%>','<%=PASSWD%>','<%=EMAIL%>','<%=Server.HtmlEncode(SUNG)%>','<%=TELE%>','<%=TELE1%>','<%=DDATE%>','<%=SEX%>','<%=IDUSE%>','<%=COP%>','<%=FAC%>','<%=DEPT%>','<%=JIK%>','<%=GUBUN%>','<%=WS_GUBUN%>','<%=UCASE(MAIL_DB)%>','<%=UCASE(ERP_DB)%>','<%=DM_ID%>','<%=MS_ID%>','<%=UR_ID%>','<%=FAC_GUBUN%>','<%=CHARSET%>','<%=rs("UR_MAILBOXSIZE")%>','<%=UR_PERMITPROTOCOL%>','<%=csUserpw%>','<%=iSystem_ClassSystem%>','<%=idx_ClassPostion%>','<%=msg_chk%>','<%=Nvcnt%>');" style="cursor:hand;">
		    <TD ALIGN="CENTER" WIDTH="05%" CLASS="TBL_DRW_00" <%=FLAG_PWD%>><%=nVcnt%></td>
		    <TD ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><%=SUNG%></td>
		    <TD ALIGN="CENTER" WIDTH="08%" CLASS="TBL_DRW_03"><%=ID%></td>
		    
		    
		    <TD ALIGN="CENTER" WIDTH="08%" CLASS="TBL_DRW_01"><%=dept_nm%>&nbsp;</td>
		    <TD ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01"><%=jik_nm%>&nbsp;</td>
		    <TD ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01"><%=PASSWD%>&nbsp;</td>
		    
		    <TD ALIGN="left"   WIDTH="15%" CLASS="TBL_DRW_01" style="padding-left:4px;"><%=EMAIL%>@<%=DM_NAME%></td>
		    <TD ALIGN="left"   WIDTH="10%" CLASS="TBL_DRW_01"><div style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr= '#E3EFFF',endColorStr= '#B5D2F7',gradientType= '0'); height:17;margin:0;border:1px solid #666666;padding:1px;color:<%=FLAG_SIZE%>;"/><%=Mailsize%></div></td>
		    <TD ALIGN="right" WIDTH="08%" CLASS="TBL_DRW_01" style="color:<%=FLAG_SIZE%>;"><%=Mailsize_conv(ur_curr_mail_size)%></td>
		    <TD ALIGN="right" WIDTH="07%" CLASS="TBL_DRW_01">M</td>
		    
		    <TD ALIGN="CENTER" WIDTH="04%" CLASS="TBL_DRW_01"><% if isnull(csUserid) then%>X<%else %>○<%end if %></td>
		    <TD ALIGN="CENTER" WIDTH="07%" CLASS="TBL_DRW_01"><%=csUserpw%>&nbsp;</td>
		    <TD ALIGN="CENTER" WIDTH="04%" CLASS="TBL_DRW_01"><%=cstate%>&nbsp;</td>
        </TR>       
        <%
        
        Mailsize    = ""
        FLAG_SIZE   = ""
        FLAG_PWD    = ""
        nVcnt = nVcnt + 1
        Rs.MoveNext
        Wend

        Rs.Close : Set Rs = Nothing
        %></FORM>
        </TABLE>
        </DIV>
    </DIV>
    <!-- CONTENT 부분 끝 -->

    <!-- 데이터 등록처리 부분 시작 -->
	<DIV class="layer_02" style="background-color:buttonface;border-top:0px solid #666666;height:340;">

		<TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" Width="980" HEIGHT="100%" BGCOLOR="buttonface">
		<TR><FORM NAME="form" ACTION="sys_user.asp" method="POST">
			<input type="hidden" name="nvcnt">
			<input type="hidden" name="ms_id" value="2" />
			<input type="hidden" name="ur_id"/>

			<TD VALIGN="TOP" ALIGN="center">		
			<FIELDSET STYLE="WIDTH:100%;">
				<LEGEND>
					<input type="radio" <%=a1%> name="slt"  value="등록" id="a1" checked>등록
					<input type="radio" <%=a2%> name="slt"  value="수정" id="a2">수정 
					<input type="radio" <%=a3%> name="slt"  value="삭제" id="a3">삭제 
				</LEGEND>

	            <TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="98%" ALIGN="CENTER">
                    <tr><td colspan="6" height=10></td></tr>
	                <tr><TD width="13%">사번/비번</TD>
                        <TD width="20%"><INPUT TYPE="TEXT" NAME="id" <%=DBL_SCPT%>> 
                                        <INPUT TYPE="TEXT" NAME="passwd" value="11111" <%=DBL_SCPT%>></TD>
                        <TD width="13%">성 명</TD>
                        <TD width="20%"><INPUT TYPE="TEXT" NAME="sung" <%=ENT_SCPT %>></TD>     
                        <TD width="13%">이메일계정</TD>
                        <TD width="20%"><INPUT TYPE="TEXT" NAME="email" <%=ENT_SCPT %>></TD>
                    </TR>
                    <TR>	
                        <TD>소속회사</TD>
                        <TD><SELECT NAME="fac" <%=SLT_SCPT%>>
		                    <%
			                SQL = "SELECT DISTINCT FAC_CD,FAC_NM,FAC_FGUBUN_NM,FAC_FGUBUN FROM ADMIN_FAC_GUBUN ORDER BY FAC_CD ASC" 
			                SET RS = SYSCON.EXECUTE(SQL)

			                DO UNTIL RS.EOF
			                    RESPONSE.WRITE "<OPTION VALUE='"& TRIM(RS("FAC_CD")) &"/"& TRIM(RS("FAC_FGUBUN")) &"'>"
			                    RESPONSE.WRITE TRIM(RS("FAC_NM")) &" "& TRIM(RS("FAC_FGUBUN_NM")) &"</OPTION>"
			                RS.MOVENEXT
			                LOOP
			                
			                RS.CLOSE : SET RS = NOTHING
			                %>
			                </SELECT></TD>
                        <TD>부서</TD>
                        <TD><SELECT NAME="dept" <%=SLT_SCPT%>>
                            <%
			                SQL = "SELECT * FROM ADMIN_DEPT_GUBUN ORDER BY DEPT_CD ASC" 
			                SET DTRS = SYSCON.EXECUTE(SQL)

			                DO UNTIL DTRS.EOF
			                    RESPONSE.WRITE "<OPTION VALUE='"& TRIM(DTRS("DEPT_CD")) &"'>"& TRIM(DTRS("DEPT_NM")) &"</OPTION>"
			                DTRS.MOVENEXT
			                
                            LOOP
                            DTRS.CLOSE : SET DTRS = NOTHING
			                %>
			                </SELECT></TD>
                		<TD>직급</TD>
                		<TD><SELECT NAME="jik" <%=SLT_SCPT%>>
		                    <%
			                SQL = "SELECT * FROM ADMIN_JIK_GUBUN  order by jik_cd desc" 
			                SET JKRS = SYSCON.EXECUTE(SQL)
			                
			                DO UNTIL JKRS.EOF
			                    RESPONSE.WRITE "<OPTION VALUE='"& TRIM(JKRS("JIK_CD")) &"'>"& TRIM(JKRS("JIK_NM")) &"</OPTION>"
			                JKRS.MOVENEXT
			                LOOP
                            JKRS.CLOSE : SET JKRS = NOTHING
		                    %>
			                </SELECT></TD>
                    </TR>   
                    <TR>
                     <TR><TD>업무DB명</TD>
                         <TD><SELECT NAME="erp_db" <%=SLT_SCPT%>>
                            <option value="vserp">vserp</option>
                            </SELECT></TD>
	                     <TD>메일DB명</TD>
	                     <TD><SELECT NAME="mail_db" <%=SLT_SCPT%>>
                            <option value="MAIL">MAIL</option>
                            </SELECT></TD>
                        <TD>HP/내선</TD>
                        <TD><INPUT TYPE="TEXT"  NAME="tele1" class="input" style="width:100;" onkeypress="return handleEnter(this, event);"> 
                            <INPUT TYPE="TEXT"  NAME="tele" class="input" style="width:45;" onkeypress="return handleEnter(this, event);"></TD>
	                </TR>   
                    <TR> 
                        <TD>설정언어</TD>
                        <TD><SELECT NAME="charset" <%=SLT_SCPT%>>
                            <option VALUE="KOR"> 한국</option>
	                        <option VALUE="ENG"> 영문</option>
	                        <option VALUE="CHN"> 중문</option></SELECT></TD>
                        <TD>성별</TD>
                        <TD><SELECT NAME="sex" <%=SLT_SCPT%>>
                            <option VALUE="0"> 남자</option>
	                        <option VALUE="1"> 여자</option></SELECT></TD>
                        <TD>당직설정</TD>
                        <TD><% CALL date_frm("A","-","","ddate","","","") %></TD>
                    </TR>
                    <TR>     
		                <TD>도메인</TD>
                        <TD><SELECT NAME="dm_id" <%=SLT_SCPT%>>
	                        <%
                            SQL ="SELECT DM_ID,MS_ID,DM_NAME FROM DOMAIN ORDER BY DM_ID desc"
                            SET DMRS = SYSCON.EXECUTE(SQL)

                            DO UNTIL DMRS.EOF
                                RESPONSE.WRITE "<OPTION VALUE='"& DMRS("DM_ID") &"'>"& TRIM(DMRS("DM_NAME")) &"</OPTION>"
                            DMRS.MOVENEXT
                            LOOP
                            DMRS.CLOSE : SET DMRS = NOTHING
                            %>	  
	                        </SELECT></TD>
                        <TD>용량(Mbyte)</TD>
                        <TD><INPUT TYPE=TEXT NAME="ur_mailboxsize" VALUE="5000" <%=ENT_SCPT %>></TD>
                         
                        <TD>사용유무</TD>
                        <TD><INPUT TYPE="RADIO" NAME="iduse"    VALUE="A"  id="iduse_1" checked> 정상 
                             <INPUT TYPE="RADIO" NAME="iduse"    VALUE="N"  id="iduse_2"> 일시중지</TD>
                        <!--<TD>설정기능</TD>
                        <TD><INPUT TYPE="CHECKBOX" NAME="ur_permit_1" VALUE="1" id="ur_permit_1" checked> SMTP
                            <INPUT TYPE="CHECKBOX" NAME="ur_permit_2" VALUE="2" id="ur_permit_2" > POP3
                            <INPUT TYPE="CHECKBOX" NAME="ur_permit_3" VALUE="4" id="ur_permit_3" checked> WEB</TD>--> 
                    </TR>	                
	              
                </TABLE>
			</FIELDSET>
			
			<div style="padding-top:5;"></div>
			
			<FIELDSET STYLE="WIDTH:100%;">

	            <TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="98%" ALIGN="CENTER">
	                <tr><TD width="13%">구분</TD>
	                    <TD width="20%"><input type="checkbox" name="msg_chk" id="msg_chk"  onclick="enable_msg_chk();"/> 메신저사용</TD>	
	                    <TD width="13%">부서명직책</TD>
	                    <TD width="20%"><SELECT NAME="isystem_classsystem" id="isystem_classsystem">
	                        <%
                            SQL ="select idx,csName From mc_ClassSystem A where idx_Uplevel <> '0' "
                            SET BizRS = SYSCON.EXECUTE(SQL)

                            DO UNTIL BizRS.EOF
                                RESPONSE.WRITE "<OPTION VALUE='"& BizRS("idx") &"'>"& TRIM(BizRS("csName")) &"</OPTION>"
                            BizRS.MOVENEXT
                            LOOP
                            BizRS.CLOSE : SET BizRS = NOTHING
                            %>	  
	                        </SELECT><SELECT NAME="idx_classpostion" id="idx_classpostion">
	                        <%
                            SQL ="select * From mc_ClassPosition"
                            SET BizRS = SYSCON.EXECUTE(SQL)

                            DO UNTIL BizRS.EOF
                                RESPONSE.WRITE "<OPTION VALUE='"& BizRS("idx") &"'>"& TRIM(BizRS("csName")) &"</OPTION>"
                            BizRS.MOVENEXT
                            LOOP
                            BizRS.CLOSE : SET BizRS = NOTHING
                            %>	  
	                        </SELECT></TD>	
	                    <TD width="13%">비밀번호</TD>	
	                    <TD width="20%"><INPUT TYPE=TEXT NAME="csuserpw" id="csuserpw" <%=ENT_SCPT %>></TD>		
                    </TR>	                
	              
                </TABLE>
			</FIELDSET>		

			<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="1" Width="100%" STYLE="TABLE-LAYOUT: FIXED;">
				<TR><TD HEIGHT="5"></TD></TR>
				<TR><td id="msg" style="padding-left:5px;"><%=MSG_CONT%> </td>
				    <TD ALIGN="RIGHT">
					   <INPUT TYPE="BUTTON" VALUE=" 확 인 " CLASS="BTN" ONCLICK="Prepared_Submit();"> 
					   <INPUT TYPE="BUTTON" VALUE=" 리 셋 " CLASS="BTN" ONCLICK="show_clear();">
					</TD></FORM>
				</TR>                
			</TABLE>		
		
		</TD></TR>
		</TABLE>
	
	</DIV>
    <!-- 데이터 등록처리 부분 끝 -->

	<SCRIPT LANGUAGE="JavaScript" TYPE="TEXT/JAVASCRIPT">
	<!--	

    /*********************************************************************************
	*  마우스 롤오버 관련 스크립트 시작
    *********************************************************************************/
    var selectedRowIndex = ""; 
	function selectrow()
	{
	    var oRow = window.event.srcElement.parentNode;
        var i, j

        // 일단 모든 Row 정상화처리
		if(selectedRowIndex!==""){
            for (i=0;i<datalist.rows(Number(selectedRowIndex)).cells.length;i++)
            {
                datalist.rows(Number(selectedRowIndex)).cells(i).style.backgroundColor = '';
                datalist.rows(Number(selectedRowIndex)).cells(i).style.color = '';
            }
        }
        
        // 해당 Row 만 반전처리
        for (j=0;j<datalist.rows(oRow.rowIndex).cells.length;j++)
        {
            datalist.rows(oRow.rowIndex).cells(j).style.backgroundColor = 'activecaption';
            datalist.rows(oRow.rowIndex).cells(j).style.color = 'white';
        }
        selectedRowIndex = oRow.rowIndex;
	}

    function show_data(id,passwd,email,sung,tele,tele1,ddate,sex,iduse,cop,fac,dept,jik,gubun,ws_gubun,mail_db,erp_db,dm_id,ms_id,ur_id,fac_gubun,charset,ur_mailboxsize,ur_permitprotocol,csuserpw,isystem_classsystem,idx_classpostion,msg_chk,nvcnt) 
	{
			var form = document.form;
			form.id.value	    = id;					
			form.passwd.value   = passwd;					
			form.email.value    = email;				
			form.sung.value		= sung;		
			
		    form.fac.value	    = fac+"/"+fac_gubun;	
						
			form.dm_id.value    = dm_id;			
			form.dept.value		= dept;			
			form.jik.value		= jik;						
			form.tele1.value    = tele1;					
			form.tele.value		= tele;		
			form.ddate.value    = ddate;
			form.sex.value      = sex;
			form.erp_db.value   = erp_db;					
			form.mail_db.value  = mail_db;
			form.charset.value  = charset;
			form.ur_id.value  = ur_id;
			form.ur_mailboxsize.value   = ur_mailboxsize;
			
			if(iduse=="A"){
			    document.getElementById("iduse_1").checked = true;
			}else{
			    document.getElementById("iduse_2").checked = true;
			}
			
			if(msg_chk=="0"){
			    
			    document.getElementById("msg_chk").checked = false;
			    document.getElementById("isystem_classsystem").disabled = true;
			    document.getElementById("idx_classpostion").disabled = true;
			    document.getElementById("csuserpw").disabled = true;
			    form.isystem_classsystem.value    = "";
			    form.idx_classpostion.value    = "";
			    form.csuserpw.value    = "";					
			}else{
		        document.getElementById("msg_chk").checked = true;
			    document.getElementById("isystem_classsystem").disabled = false;
			    document.getElementById("idx_classpostion").disabled = false;
			    document.getElementById("csuserpw").disabled = false;
			    form.isystem_classsystem.value    = isystem_classsystem;
			    form.idx_classpostion.value    = idx_classpostion;
			    form.csuserpw.value    = csuserpw;			
			
			}
			
                
			/*if(ur_permitprotocol=="1"){
			    document.getElementById("ur_permit_1").checked = true;
			    document.getElementById("ur_permit_2").checked = false;
			    document.getElementById("ur_permit_3").checked = false;
			}else if(ur_permitprotocol=="5"){
			    document.getElementById("ur_permit_1").checked = true;
			    document.getElementById("ur_permit_2").checked = false;
			    document.getElementById("ur_permit_3").checked = true;
			}else{
			    document.getElementById("ur_permit_1").checked = true;
			    document.getElementById("ur_permit_2").checked = true;
			    document.getElementById("ur_permit_3").checked = true;
			}*/
			
			disable_field('0');
			selectrow()
			document.all("A2").checked = true;
			document.all("A1").disabled = true;
			document.all("A2").disabled = false;
			document.all("A3").disabled = false;
	}
	
	function disable_field(str)
	{   
	    if(str=="0"){
	        document.getElementById("id").readOnly  = true;
	        document.getElementById("email").readOnly = true;
	        document.getElementById("id").style.background  = "#cccccc";
	        document.getElementById("email").style.background  = "#cccccc";
	    }else if(str=="1"){
	        document.getElementById("id").readOnly = false;
	        document.getElementById("email").readOnly = false;
	        document.getElementById("id").style.background  = "#ffffff";
	        document.getElementById("email").style.background  = "#ffffff";
	    }
	}
	
	function enable_msg_chk()
	{
	    document.getElementById("isystem_classsystem").disabled = false;
	    document.getElementById("idx_classpostion").disabled = false;
	    document.getElementById("csuserpw").disabled = false;			
	}

	function show_clear()
	{
	    disable_field('1');
		document.form.reset();
		document.all("A1").checked = true;
		document.all("A1").disabled = false;
		document.all("A2").disabled = true;
		document.all("A3").disabled = true;
	}
	
	function Prepared_Submit()
	{
	
	    var form = document.form; 
	    var querystring     = ""; 
	    var querystring1    = "";
	    
	    if(document.getElementById("a1").checked==true){
	        var SLT ="등록";
	    }else if(document.getElementById("a2").checked==true){
	        var SLT ="수정";
	    }else if(document.getElementById("a3").checked==true){
	        var SLT ="삭제";
        }
        
		if (form.id.value==""){alert("사번을 입력하세요\n\n삼성화학은 s+사번\n명성정밀은 m+사번\n유창라바는 y+사번\n동아중국은 c+사번\n\n 처럼 입력바랍니다.");
			form.id.focus();return;
		}
		if (form.passwd.value==""){alert("비밀번호를 입력하세요!");form.passwd.focus();return;}
		if (form.sung.value==""){alert("성명을 입력하세요!");form.sung.focus();return;}
		if (form.email.value==""){alert("메일계정을 입력하세요!");form.email.focus();return;}
		if (form.passwd.value==""){alert("비밀번호를 입력하세요!");form.passwd.focus();return;}
		
		if (form.dept.value==""){alert("부서를 선택하세요!");form.dept.focus();return;}		
		if (form.jik.value==""){alert("직책을 선택하세요!");form.jik.focus();return;}
		
		
		if (form.ur_mailboxsize.value==""){alert("메일용량을 입력하세요!");form.ur_mailboxsize.focus();return;}
		
		if (document.getElementById("msg_chk").checked == true){		
		    if (form.isystem_classsystem.value==""){alert("메신저를 사용하시려면 부서를 선택하세요!");form.isystem_classsystem.focus();return;}
		    if (form.idx_classpostion.value==""){alert("메신저를 사용하시려면 직책을 선택하세요!");form.idx_classpostion.focus();return;}
		    if (form.csuserpw.value==""){alert("메신저를 사용하시려면 메신저 비밀번호를 입력하세요!");form.csuserpw.focus();return;}
		    
            var isystem_classsystem = form.isystem_classsystem.value;
            var idx_classpostion = form.idx_classpostion.value;
        }else{
            var isystem_classsystem = 0;
            var idx_classpostion = 0;
        }

		if (document.getElementById("iduse_1").checked == true){
		    var idusev = "A"
        }else if(document.getElementById("iduse_2").checked == true){
		    var idusev = "N"
        }else{
		    var idusev = "A"
        }

		if(document.getElementById("msg_chk").checked == true){
	        var msg_chkv = "Y";	
		}else{
	        var msg_chkv = "N";		
		}
                
        if (confirm("정말로 "+SLT+"하시겠습니까?")) {
                
            var fac_temp = form.fac.value;
            var ftemp = fac_temp.split("/");

            if(SLT=="등록"){
                querystring = " exec sp_signon 'ADD',"
                            + " '"+form.id.value+"','"+form.passwd.value+"','"+form.sung.value+"',"
                            + " '"+form.email.value+"','"+ftemp[0]+"','"+ftemp[1]+"','"+form.dept.value+"',"
                            + " '"+form.jik.value+"','"+form.erp_db.value+"','"+form.mail_db.value+"',"
                            + " '"+form.tele1.value+"','"+form.tele.value+"','"+form.charset.value+"',"
                            + " '"+form.sex.value+"','"+form.ddate.value+"',"+form.dm_id.value+",'"+form.ur_id.value+"',"
                            + " '"+form.ur_mailboxsize.value+"','"+idusev+"','"+msg_chkv+"',"
                            + " "+isystem_classsystem+","+idx_classpostion+",'"+form.csuserpw.value+"',"
                            + " "+form.ms_id.value+",'<%=Request.ServerVariables("REMOTE_ADDR")%>'";
            }else if(SLT=="수정"){
                querystring = " exec sp_signon 'MOD',"
                            + " '"+form.id.value+"','"+form.passwd.value+"','"+form.sung.value+"',"
                            + " '"+form.email.value+"','"+ftemp[0]+"','"+ftemp[1]+"','"+form.dept.value+"',"
                            + " '"+form.jik.value+"','"+form.erp_db.value+"','"+form.mail_db.value+"',"
                            + " '"+form.tele1.value+"','"+form.tele.value+"','"+form.charset.value+"',"
                            + " '"+form.sex.value+"','"+form.ddate.value+"',"+form.dm_id.value+","+form.ur_id.value+","
                            + " '"+form.ur_mailboxsize.value+"','"+idusev+"','"+msg_chkv+"',"
                            + " "+isystem_classsystem+","+idx_classpostion+",'"+form.csuserpw.value+"',"
                            + " "+form.ms_id.value+",'<%=Request.ServerVariables("REMOTE_ADDR")%>'";
            }else if(SLT=="삭제"){   
                querystring = " exec sp_signon 'DEL',"
                            + " '"+form.id.value+"','"+form.passwd.value+"','"+form.sung.value+"',"
                            + " '"+form.email.value+"','"+ftemp[0]+"','"+ftemp[1]+"','"+form.dept.value+"',"
                            + " '"+form.jik.value+"','"+form.erp_db.value+"','"+form.mail_db.value+"',"
                            + " '"+form.tele1.value+"','"+form.tele.value+"','"+form.charset.value+"',"
                            + " '"+form.sex.value+"','"+form.ddate.value+"',"+form.dm_id.value+","+form.ur_id.value+","
                            + " '"+form.ur_mailboxsize.value+"','"+idusev+"','"+msg_chkv+"',"
                            + " "+isystem_classsystem+","+idx_classpostion+",'"+form.csuserpw.value+"',"
                            + " "+form.ms_id.value+",'<%=Request.ServerVariables("REMOTE_ADDR")%>'";
            }       
            
            
            var Result = RSExecute("/admin/sys_user.asp","sys_user","add",querystring);
            var retval=Result.return_value; 
            
            if(retval=='OK'){  
                document.getElementById("msg").innerHTML = querystring;            
            }else{
                document.getElementById("msg").innerHTML = querystring;
                //document.getElementById("msg").innerHTML = "<span style='color:red;'>오류발생</span>";            
            }
            setTimeout('clear_msg()',3000);                
        }       
        
       
	}
	
    /********* Message Clear 로직 **************/
	function clear_msg()
	{
            document.getElementById("msg").innerHTML = ""; 
	}	
	//-->
	</SCRIPT>
        
    <%

    Class tscript

        public function sys_user(slt,querystring) 
        
            syscon.execute(querystring)
            sys_user      = "OK"
                         

        exit function
        end function

    end class

    set public_description = new tscript

    RSDispatch

    %> 
    <!--#include Virtual = "/INC/_ScriptLibrary/rs.asp"  -->       
    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
    

