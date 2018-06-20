    <!--#include Virtual = "/inc/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 사용자 관리                              			            --> 
    <!-- 내용: 그룹웨어 사용자 관리                                                 --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 14일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/inc/INC_BODY.ASP" -->
    <%
        Public Function dept_convert(str)
	
	    sql="select dept_cd,dept_nm from ADMIN_DEPT_GUBUN where dept_cd='" & str &"' "
	    set deptrs=syscon.execute(sql)

	    IF deptrs.EOF or deptrs.BOF then
	    dept_convert="---"
	    else
	    dept_convert=deptrs("dept_nm")
	    end if

	    deptrs.close
	    set deptrs=nothing

        End Function
    
        slt			    = Trim(Request("slt"))
        idx			    = Trim(Request("idx"))
        name		    = Trim(Request("name"))
        uid		        = Trim(Request("uid"))
        passwd	        = Trim(Request("passwd"))
        email		    = Trim(Request("email"))
        tele		    = Trim(Request("tele"))
        tele1		    = Trim(Request("tele1"))
        addr		    = Trim(Request("addr"))
        fg			    = Trim(Request("fg"))
        stat		    = Trim(Request("stat"))		   
        sex		        = Trim(Request("sex"))
        iduse		    = Trim(Request("iduse"))
        cop		        = Trim(Request("cop"))
        ip			    = Trim(Request("ip"))
        fac		        = Trim(Request("fac"))
        
        IF fac <> "" THEN
        fac_temp	    = split(fac,"/")
        fac = fac_temp(0) : fac_gubun = fac_temp(1)
        END IF
        
        dept		    = Trim(Request("dept"))
        jik			    = Trim(Request("jik"))
        ilnkyn		    = Trim(Request("linkyn"))
	    gubun		    = Trim(Request("gubun"))  
        ws_gubun	    = Trim(Request("ws_gubun"))
        mail		    = Trim(Request("mail"))
        erp		        = Trim(Request("erp"))
        dm_id		    = Request("dm_id")
	    ms_id		    = Request("ms_id") : IF	ms_id = "" THEN ms_id=1 END IF
	    ddate		    = Trim(Request("ddate"))
	    ur_mailboxsize  = Request("ur_mailboxsize")
	    ur_permit_1     = Request("ur_permit_1")
	    ur_permit_2     = Request("ur_permit_2")
	    ur_permit_3     = Request("ur_permit_3")
	    
	    ur_permitProtocol	= Cint(ur_permit_1)+Cint(ur_permit_2)+Cint(ur_permit_3)
	    ur_permitProtocol	= Cint(ur_permitProtocol)

        gubun1	        = Request("gubun1")
        gubun2	        = Request("gubun2")

	  IF dept <> ""  THEN  hak = dept_convert(dept)	  END IF
        If slt = ""	     THEN  S2 = "checked"			  END IF

        SELECT case slt
		case ""
			S2 ="checked"
			S3 ="disabled='true'"
			S4 ="disabled='true'"
            case "조회"  
			S2 = "disabled='true'"
			S3 = "checked"
			S4 = ""
            case "등록"  S2 = "checked"
            case "수정"  S3 = "checked"
            case "삭제"  S4 = "checked"
        end    select

        IF  slt  <>  ""   THEN
       
               SELECT CASE  slt
               CASE  "등록"
                 IF  sung  <>  ""  and  uid  <>  ""  and  passwd  <>  ""  Then

                         Set Rs1 = syscon.Execute("select  id  from  profile  where  id='" & uid & "' and fac='"& fac &"' ")

                         IF  Not Rs1.Eof   Then
							 %><script language='javascript'>
							 <!--
							 alert("이미 등록된 사번입니다.\n\n다른 사번(ID)을 입력해주십시오.");
							 history.go(-1);
							 //-->
							 </script><%	                         

				        ELSE
                             Set Rs2 = syscon.Execute("select  email from profile where email='" & email & "' and fac='"& fac &"' ")
                             IF  Not  Rs2.Eof   Then
							 %><script language='javascript'>
							 <!--
							 alert('이미 등록된 Email ID입니다.\n\n다른 Email ID를 입력해주십시오.');
							 history.go(-1);
							 //-->
							 </script><%	                             
                             ELSE
						
						Set Rs3 = syscon.Execute("SELECT ur_id FROM Domain D INNER JOIN DomainUser U ON D.dm_id=U.dm_id WHERE D.ms_id="& ms_id &" AND D.dm_id="& dm_id &" AND U.ur_mailId='"& email &"' AND U.ur_pwd='"& passwd &"' ")

						IF  Not  Rs3.Eof   Then
							 %><script language='javascript'>
							 <!--
							 alert('DomainUser Table is used');
							 history.go(-1);
							 //-->
							 </script><%	
						ELSE

							  Set Rs4=syscon.execute(" select dm_name from domain where dm_id ="& dm_id )						
							  IF Not Rs4.EOF  THEN Femail = email&"@"&RS4("dm_name") END IF
							  
							  IF  cop =  "01"  THEN fg  =  "1" ELSE fg  =  "2" END IF			
							  IF ddate = "1"   THEN ddate =year(now())&"-"&month(now())&"-"&day(now())  ELSE	 ddate = NULL END IF
							  IF gubun = "1"  THEN gubun_value = "F0000" ELSE  gubun_value = fac END IF

							  syscon.Execute("insert into profile(id, passwd, email, femail, sung, tele, tele1, addr, ddate, hak, fg, stat, sex, iduse, cop, ip, FAC, DEPT, JIK, linkyn, gubun, ws_gubun, mail_db, erp_db,dm_id,ms_id) VALUES ('" & uid & "','" & passwd & "','" & email & "','"& Femail &"',N'" & name & "','" & tele & "','" & tele1 & "','" & addr & "','" & ddate & "','" & hak & "','" & fg & "','" & stat & "','" & sex & "','" & iduse & "','" & cop & "','" & ip & "','" & fac & "','" & dept & "','" & jik & "','N','" & gubun_value & "','" & ws_gubun & "','" & mail & "','" & erp & "'," & dm_id & "," & ms_id & ")")

							  '*** Default : Forward 허용 불가 : 0 , 사용자가 사용 가능한 서비스 지정(1: Smtp 허용 2: Pop3 불가 4: WebMail 허용)
							  syscon.Execute("INSERT INTO DomainUser (dm_id, ur_mailboxsize, ur_mailId, ur_name, ur_pwd, ur_forward_use, ur_markSpamtag, ur_permitProtocol) VALUES ("& dm_id &", "& ur_mailboxsize &", '"& email &"', N'"& name &"','"& passwd &"' , 0, 1, "& ur_permitProtocol &" )")
							  
							  %>
							  <script language='javascript'>
							  <!--
							  //opener.document.location.reload();
							  //window.close();
							  //-->
							  </script>
							  <%
					  end if   
                             end if   
                         end if    
                 end if   

		     		 

               CASE  "수정"
			   Set Rs = syscon.Execute("select  *  from  profile  where  idx=" & idx )

                     IF  Rs.Eof Then
				  %>
				  <script language='javascript'>
				  <!--
				  alert('등록된 자료가 없습니다.');
				  history.go(-1);
				  //-->
				  </script> 
				  <%
                     ELSE

				  Set Rs1=syscon.execute(" select dm_name from domain where dm_id ="& dm_id )	
				  
				  IF Not Rs1.EOF  THEN Femail = email&"@"&RS1("dm_name") END IF
							  
				  IF cop =  "01"	 THEN fg  =  "1" ELSE fg  =  "2" END IF			
				  IF ddate = "1"   THEN ddate =year(now())&"-"&month(now())&"-"&day(now())  ELSE ddate = NULL END IF
				  IF gubun = "1"  THEN gubun_value = "F0000" ELSE  gubun_value = fac END IF

				  syscon.Execute("update profile set passwd = '" & passwd & "', email = '" & email & "', femail = '"& Femail &"', sung =N'" & name & "', tele = '" & tele & "', tele1 = '" & tele1 & "', addr = '" & addr & "', ddate = '" & ddate & "', hak = '" & hak & "', fg = '" & fg & "', stat = '" & stat & "', sex  = '" & sex & "', iduse = '" & iduse & "', cop = '" & cop & "', ip = '" & ip & "', FAC = '" & fac & "', DEPT = '" & dept & "', JIK = '" & jik & "', linkyn = 'N', gubun = '" & gubun_value & "', ws_gubun = '" & ws_gubun & "', mail_db = '" & mail & "', erp_db = '" & erp & "',dm_id = " & dm_id & ", ms_id = "& ms_id &"  where idx=" & idx  )

				  Set Rs2 = syscon.Execute("SELECT ur_id FROM Domain D INNER JOIN DomainUser U ON D.dm_id=U.dm_id WHERE  U.ur_mailId='"& email &"' AND U.ur_pwd='"& passwd &"' ")

				  IF Rs2.EOF THEN

					  '*** Default : Forward 허용 불가 : 0 , 사용자가 사용 가능한 서비스 지정(1: Smtp 허용 2: Pop3 불가 4: WebMail 허용)
					  IF dm_id <> "" and ur_mailboxsize <> "" and ur_permitProtocol <> "" THEN
					  syscon.Execute ("INSERT INTO DomainUser (dm_id, ur_mailboxsize, ur_mailId, ur_name, ur_pwd, ur_forward_use, ur_permitProtocol, ur_markSpamtag) VALUES ("& dm_id &", "& ur_mailboxsize &",'"& email &"',N'"& name &"','"& passwd &"',0,"& ur_permitProtocol &" ,1 )")
					  END IF
				  
				  ELSE
				      IF ms_id <> "" THEN
					  syscon.Execute("update DomainUser set dm_id=" & dm_id & ",ur_mailboxsize=" & ur_mailboxsize & ",ur_mailId='" & email & "',ur_name=N'" & name & "',ur_pwd='" & passwd & "',ur_forward_use=0,ur_permitProtocol=" & ur_permitProtocol &", ur_markSpamtag=1 where ur_id="& Rs2("ur_id") &" ")
                      END IF
				 END IF

				 %>
				 <script language="javascript">
				 <!--
				 opener.document.location.reload();
				 window.close();
				 //-->
				 </script>
				 <%

                     END IF

               CASE  "삭제"

			   Set Rs = syscon.Execute("select  *  from  profile  where  idx=" & idx )
                     IF  Rs.Eof Then
						 %>
						 <script language='javascript'>"
						 <!--
						 alert('등록된 자료가 없습니다.');
						 history.go(-1);
						 //-->
						 </script>
						 <%
			   else
                          syscon.Execute("update profile set femail=NULL,iduse='N' where idx=" & idx & "")

				  Set Rs1 = syscon.Execute("SELECT ur_id FROM Domain D INNER JOIN DomainUser U ON D.dm_id=U.dm_id WHERE D.ms_id="& ms_id &" AND U.ur_mailId='"& email &"' AND U.ur_pwd='"& passwd &"' ")

				  IF Not Rs1.EOF THEN

				     syscon.Execute("DELETE FROM [AddressGroup] WHERE ag_id IN (SELECT ac_id FROM [AddressCard] WHERE ur_id = "& Rs1("ur_id") &" ) OR ac_id IN (SELECT ac_id FROM [AddressCard] WHERE ur_id = "& Rs1("ur_id") &" ") 

				     '3.1.2 자동 응답 삭제
				     syscon.Execute("DELETE FROM [AutoResponse] WHERE ur_id = "& Rs1("ur_id") &" ") 
				    
				     '3.1.3 NtUserId 삭제
				     syscon.Execute("DELETE FROM [NtUserId] WHERE ur_id = "& Rs1("ur_id") &" ") 

				     '3.1.4 외부 POP3 삭제
				     syscon.Execute("DELETE FROM [ExternalPop3] WHERE ur_id = "& Rs1("ur_id") &" ") 

				     '3.1.5 UserFilter 삭제
				     syscon.Execute("DELETE FROM [UserFilter] WHERE ur_id = "& Rs1("ur_id") &" ") 
				  
				     '3.1.6 UserBwList 삭제
				     syscon.Execute("DELETE FROM [UserBwList] WHERE ur_id = "& Rs1("ur_id") &" ") 

				     '3.1.7 DomainUser 삭제
				     syscon.Execute("DELETE FROM [DomainUser] WHERE ur_id = "& Rs1("ur_id") &" ") 
				 
				 END IF
				 %>
						 <script language='javascript'>
						 <!--
						 alert('해당 ID는 사용중지 되었습니다.\n\nEmail 계정이 삭제되었으므로 해당메일은 수신되지 않습니다.');
						 opener.document.location.reload();
						 window.close();
						 //-->
						 </script>
						 <%
                     end if
		   
         end select
	 
	 Set Rs = syscon.Execute("select idx, id, passwd, email, femail, sung, tele, tele1, addr, ddate, hak, fg, stat, sex, iduse, cop, ip, FAC, DEPT, JIK, linkyn, wdate, gubun, ws_gubun, mail_db, erp_db, IsNULL(dm_id,1) AS dm_id,ms_id from profile where  id='" & uid &"'" )
                     
	 IF  Not Rs.Eof THEN
		idx			    = Trim(Rs("idx"))
		name		    = Trim(Rs("sung"))
		uid			    = Trim(Rs("id"))
		passwd		    = Trim(Rs("passwd"))
		email		    = Trim(Rs("email"))
		tele			    = Trim(Rs("tele"))
		tele1		    = Trim(Rs("tele1"))
		addr		    = Trim(Rs("addr"))
		fg			    = Trim(Rs("fg"))
		stat			    = Trim(Rs("stat"))		   
		sex			    = Trim(Rs("sex"))
		iduse		    = Trim(Rs("iduse"))
		cop			    = Trim(Rs("cop"))
		ip			    = Trim(Rs("ip"))
		fac			    = Trim(Rs("fac"))
		dept		    = Trim(Rs("dept"))
		jik			    = Trim(Rs("jik"))
		ilnkyn		    = Trim(Rs("linkyn"))
		gubun		    = Trim(Rs("gubun"))
		ws_gubun	    = Trim(Rs("ws_gubun"))
		mail			    = Trim(Rs("mail_db"))
		erp			    = Trim(Rs("erp_db"))
		ddate		    = Trim(Rs("ddate"))
		dm_id		    = Rs("dm_id")
		ms_id		    = Rs("ms_id")

		IF sex   = "0"    Then  A1 = "checked"  else  A2 = "checked"  end if
		IF cop   = "01"  Then  B1 = "checked"  else  B2 = "checked"  end if 

		SELECT CASE iduse
		CASE "A"
		    iduse_1 = "checked"
		CASE "N"
		    iduse_2 = "checked"
		END SELECT
	
		IF gubun = "F0000"	  THEN gubun_01 = "checked"	  ELSE gubun_02 = "checked"		END IF	
		IF ws_gubun = "1"	  THEN ws_gubun_01 = "checked"	  ELSE ws_gubun_02 = "checked"		END IF	
		IF IsNULL(ddate)	  or ddate =""      THEN ddate_02 = "checked"	        ELSE ddate_01 = "checked"			END IF

		SQL ="select * From DomainUser where dm_id = "& dm_id &" and ur_mailId ='"& email &"' "
		Set mailRS = syscon.execute(sql)

		IF Not mailRs.EOF THEN
		    ur_mailboxsize    = mailRs("ur_mailboxsize")
		    ur_permitProtocol = mailRs("ur_permitProtocol")

		    Select Case ur_permitProtocol
		    case 1 
		    ur_permit_1_chk = "checked"
		    ur_permit_2_chk = ""
		    ur_permit_3_chk = ""
		    case 3
		    ur_permit_1_chk = "checked"
		    ur_permit_2_chk = "checked"
		    ur_permit_3_chk = ""
		    case 5
		    ur_permit_1_chk = "checked"
		    ur_permit_2_chk = ""
		    ur_permit_3_chk = "checked"
		    case 7
		    ur_permit_1_chk = "checked"
		    ur_permit_2_chk = "checked"
		    ur_permit_3_chk = "checked"
		    End Select
		    
		END IF

	END IF


    END IF

    Public Function stat_convert(str)

    IF str = "0" THEN 
    stat_convert = "일반사용자"
    ELSEIF str = "1" THEN
    stat_convert = "운영자"
    ELSEIF str = "2" THEN
    stat_convert = "문서관리자"
    ELSEIF str = "3" THEN
    stat_convert = "검토자"
    ELSEIF str = "4" THEN
    stat_convert = "결재자"
    ELSEIF str = "5" THEN
    stat_convert = "검토/결재자"
    ELSE
	END IF

    End Function


    %>    

    <BODY CLASS="POPUP_01" SCROLL="no">
    <% CALL Popup_generate("사용자 관리","등록시 반드시 형식에 맞게끔 입력하시기 바랍니다.")%>

    <Fieldset align=center style="width:98%;">
    <LEGEND><form name="form" target="" method="post"  action="sys_user_wrt.asp">
			<input type="hidden"  name="idx" value="<%=idx%>" >
            <input type="radio" <%=(S2)%> name="slt" value="등록" >등록                
            <input type="radio" <%=(S3)%> name="slt" value="수정" >수정
            <input type="radio" <%=(S4)%> name="slt" value="삭제" >삭제
    </LEGEND>

    <Div class="div_separate_10"></div>

    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" >
	    <tr><td ID="TBL_Title">ID/Email계정</td>
            <td ID="TBL_Data"><input type="text" name="uid" maxlength="07"  size="10" <% If uid <> "" then %> value="<%=uid%>" readonly style="background-color:dedede;" <% else %> value="" <% end if %>  onKeyPress="return handleEnter(this, event);" class=INPUT  > <input type="text" name="email"  size="10"  value="<%=(email)%>" onKeyPress="return handleEnter(this, event);" class=INPUT></td>     
            <td ID="TBL_Title">비밀번호</td>
            <td ID="TBL_Data"><input type="text" name="passwd" size="15" <% If request("idx") <> "" then %> value="<%=passwd%>" <% else %> value="" <% end if %> onKeyPress="return handleEnter(this, event);" class=INPUT>
        </tr>
        <tr><td ID="TBL_Title">도메인</td>
            <td ID="TBL_Data"><select name="dm_id" size="1">
	      <%
	      SQL ="select dm_id,ms_id,dm_name from Domain"
	      Set dmRs = syscon.execute(sql)
	      %>
	      <option value="">선택하세요</option>
	      <%
	      Do Until dmRs.EOF
		    %>
		    <% IF dm_id = dmRs("dm_id") THEN %>
		    <option value="<%=dm_id%>" selected><%=Trim(dmRS("dm_name"))%></option>
		    <% ELSE %>
		    <option value="<%=dmRS("dm_id")%>"><%=Trim(dmRS("dm_name"))%></option>
		    <% END  If	
	      dmRs.Movenext
	      Loop
	      dmRS.Close
	      %>	  
	      </SELECT></td>
            <td ID="TBL_Title">성 명</td>
            <td ID="TBL_Data"><input type="text" name="name"   size="20"  value="<%=(name)%>"  onKeyPress="return handleEnter(this, event);" class=INPUT></td>
        </tr>   

        <tr><td ID="TBL_Title">공장구분</td>
            <td ID="TBL_Data">

		    <%
			    SQL = "SELECT distinct fac_cd,fac_nm,fac_fgubun_nm,fac_fgubun FROM ADMIN_FAC_GUBUN order by FAC_cd desc" 
			    Set Rs = syscon.Execute(SQL)
			    %>
			    <select name="FAC" size="1">
			    <option value="">선택하세요</option>
			    <%
			    Do until Rs.EOF
			    FAC_nm = trim(rs("FAC_nm"))
			    FAC_cd = trim(rs("FAC_cd"))
			    fac_fgubun_nm = trim(rs("fac_fgubun_nm"))
			    fac_fgubun = trim(rs("fac_fgubun"))

			    IF trim(FAC)=FAC_cd then
			    %>
			    <option value="<%=trim(FAC_cd)%>/<%=trim(fac_fgubun)%>" selected><%=trim(FAC_nm)%><%=trim(fac_fgubun_nm)%></option>
			    <%
			    else
			    %>
			    <option value="<%=trim(FAC_cd)%>/<%=trim(fac_fgubun)%>"><%=trim(FAC_nm)%><%=trim(fac_fgubun_nm)%></option>
			    <%
			    end if
			    Rs.Movenext
			    Loop
			    %>
			    </select>
			    <%
			    Rs.close
			    Set Rs = Nothing
		    %>
    		
		    </td>
            <td  ID="TBL_Title">부서 / 직급</td>
            <td ID="TBL_Data">
    		
		    <%
			    SQL = "SELECT * FROM ADMIN_DEPT_GUBUN" 
			    Set Rs = syscon.Execute(SQL)
			    %>
			    <select name="dept" size="1">
			    <option value="">선택하세요</option>
			    <%
			    Do until Rs.EOF
			    dept_nm = trim(rs("dept_nm"))
			    dept_cd = trim(rs("dept_cd"))

			    IF trim(DEPT)=dept_cd then
			    %>
			    <option value="<%=trim(dept_cd)%>" selected><%=trim(dept_nm)%></option>
			    <%
			    else
			    %>
			    <option value="<%=trim(dept_cd)%>"><%=trim(dept_nm)%></option>
			    <%
			    end if
			    Rs.Movenext
			    Loop
			    %>
			    </select>
			    <%
			    Rs.close
			    Set Rs = Nothing
		    %>
    		
		    <%
			    SQL = "SELECT * FROM ADMIN_JIK_GUBUN" 
			    Set Rs = syscon.Execute(SQL)
			    %>
			    <select name="jik" size="1">
			    <option value="">선택하세요</option>
			    <%
			    Do until Rs.EOF
			    jik_nm = rs("jik_nm")
			    jik_cd = rs("jik_cd")

			    IF JIK=jik_cd then
			    %>
			    <option value="<%=trim(jik_cd)%>" selected><%=trim(jik_nm)%></option>
			    <%
			    else
			    %>
			    <option value="<%=trim(jik_cd)%>"><%=trim(jik_nm)%></option>
			    <%
			    end if
			    Rs.Movenext
			    Loop
			    %>
			    </select>
			    <%
			    Rs.close
			    Set Rs = Nothing
		    %>
		    </td>
        </tr>   
        <tr><td ID="TBL_Title">내선번호</td>
            <td ID="TBL_Data"><input type="text" name="tele" size="5"  value="<%=(tele)%>" onKeyPress="return handleEnter(this, event);" class=INPUT></td>
            <td ID="TBL_Title">사/외구분</td>
            <td ID="TBL_Data"><Input type="radio" name="cop" value="01" <%=B1%> > 사내 <Input type="radio" name="cop" value="05" <%=B2%> >외주</td>
        </tr>   

        <tr><td ID="TBL_Title">휴대번호</td>
            <td ID="TBL_Data"><input type="text" name="tele1" size="20"  value="<%=(tele1)%>" onKeyPress="return handleEnter(this, event);" class=INPUT></td>
            <td ID="TBL_Title">사용유무</td>
            <td ID="TBL_Data"><Input type="radio" name="iduse" value="A" <%=iduse_1%> > 정상 <Input type="radio" name="iduse" value="N" <%=iduse_2%> > 불가</td>
        </tr>   
        <tr><td ID="TBL_Title">유저권한</td>
            <td ID="TBL_Data">

		    <select name="stat" size="1">
		    <option value="">선택하세요</option>
		    <%
		    IF stat <> "" then
		    %>
		    <option value="<%=trim(stat)%>" selected ><%=stat_convert(stat)%></option>
		    <%
		    end if
		    %>
		    <option value="0">일반사용자</option>
		    <option value="1">운영자</option>
		    <option value="2">문서관리자</option>
		    <option value="3">검토자</option>
		    <option value="4">결재자</option>
		    <option value="5">검토/결재자</option>

		    </select>

		    </td>
            <td ID="TBL_Title">성별</td>
            <td ID="TBL_Data"><Input type="radio" name="sex" value="0" <%=A1%> > 남자 <Input type="radio" name="sex" value="1" <%=A2%> > 여자</td>
         </tr>
         <tr><td ID="TBL_Title">집주소</td>
             <td ID="TBL_Data1" colspan="3"><input type="text" name="addr"  size="60" value="<%=(addr)%>" onKeyPress="return handleEnter(this, event);" class=INPUT></td>
         </tr>
         <tr><td ID="TBL_Title">공용유무</td>
             <td ID="TBL_Data"><Input type=radio name=gubun value="1" <%=gubun_01%>> 공용그룹  <input type=radio name=gubun value="0" <%=gubun_02%>> 자체그룹</td>
	         <td ID="TBL_Title">당직설정</td>
	         <td ID="TBL_Data"><Input type=radio name="ddate" value="1" <%=ddate_01%>> 설정  <input type=radio name="ddate" value="0" <%=ddate_02%>> 미설정</td>
	     </tr>
	     <tr><td  ID="TBL_Title">업무일지사용</td>
	         <td ID="TBL_Data1" colspan="3"><Input type="radio" name="ws_gubun" value="1" <%=ws_gubun_01%> > 사용 <Input type="radio" name="ws_gubun" value="0" <%=ws_gubun_02%> > 미사용</td>
	     </tr>
         <tr><td ID="TBL_Title">업무DB명</td>
             <td ID="TBL_Data"><Input type=text name=erp size=20 value="<%=erp%>" class=INPUT></td>
	         <td ID="TBL_Title">메일DB명</td>
	         <td ID="TBL_Data"><Input type=text name=mail size=20 value="<%=mail%>" class=INPUT></td>
	    </tr>
    </table>
    <div style="padding-top:5px;"></div>

    </fieldset>

    <Fieldset align=center style="width:98%;">
    <LEGEND><span style="font-weight:bold; font-size:9pt;" > 메일계정 추가정보  </span></legend>               
    <div style="padding-top:10px;"></div>

    <table border="0" cellpadding="1" cellspacing="1" width="95%">
         <tr><td  ID="TBL_Title">메일할당용량</td>
             <td ID="TBL_Data"><Input type=text name="ur_mailboxsize" size=20 value="<%=ur_mailboxsize%>" class=INPUT> Mbytes</td>
	         <td  ID="TBL_Title">메일서버</td>
             <td ID="TBL_Data"><select name="ms_id" size="1">
		    <%
		    SQL ="select ms_id,ms_name from mailserver"
		    Set msRs = syscon.execute(sql)
		    %>
		    <option value="">선택하세요</option>
		    <%
		    Do Until msRs.EOF
		        %>
		        <% IF ms_id = msRs("ms_id") THEN %>
		        <option value="<%=ms_id%>" selected><%=msRs("ms_name")%></option>
		        <% ELSE %>
		        <option value="<%=msRs("ms_id")%>"><%=msRs("ms_name")%></option>
		        <% END  If	
		    msRs.Movenext
		    Loop
		    msRs.Close
		    %></td>
         </tr>
         <TR><td  ID="TBL_Title">설정된 기능</td>
	         <td width="500" colspan=3><Input type=checkbox name=ur_permit_1 value="1" <%=ur_permit_1_chk%>> Smtp 허용 + <Input type=checkbox name=ur_permit_2 value="2" <%=ur_permit_2_chk%>> Pop3 허용 + <Input type=checkbox name=ur_permit_3 value="4" <%=ur_permit_3_chk%>>WebMail 허용</td>
	    </tr>
    </table>	

    <div style="padding-top:5px;"></div>
    </fieldset>

    <div style="padding-top:5px;"></div>

    <table border="0" cellpadding="1" cellspacing="1" width="100%">
        <tr><td align="right" style="padding-right:15px;"><input type="submit"  value=" 확 인 " class=btn > <input type="button" value=" 닫 기 "  onclick="self.close();" class=btn></td></form>
        </tr>                
    </table>
    	   
    <!--#include Virtual = "/inc/INC_FOOT.ASP" -->  
    <SCRIPT language="JavaScript">
    <!--
    self.moveTo((screen.availWidth-700)/2,(screen.availHeight-550)/2)
    self.resizeTo(700,550)
    //-->
    </SCRIPT>
