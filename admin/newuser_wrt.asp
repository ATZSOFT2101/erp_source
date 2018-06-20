    <!--#include Virtual = "/inc/dbconnect.ASP"  -->
    <!--#include Virtual = "/INC/FUNCTION/ERP_FUNCTION.ASP" -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 회원가입                              			            --> 
    <!-- 내용: 그룹웨어 사용자 관리                                                 --> 
    <!-- 작 성 자 : 오세문                                                        --> 
    <!-- 작성일자 : 2012년 4월 12일                                                 --> 
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
	    ur_mailboxsize  = 5000
	    ur_permit_1     = Request("ur_permit_1")
	    ur_permit_2     = Request("ur_permit_2")
	    ur_permit_3     = Request("ur_permit_3")
	    
	    ur_permitProtocol	= Cint(ur_permit_1)+Cint(ur_permit_2)+Cint(ur_permit_3)
	    ur_permitProtocol	= 0

        gubun1	        = Request("gubun1")
        gubun2	        = Request("gubun2")

	  IF dept <> ""  THEN  hak = dept_convert(dept)	  END IF
      If slt = ""	 THEN  S2 = "checked"			  END IF

      fac = "F0001"
      iduse = "N"

        IF  slt  <>  ""   THEN
       
           SELECT CASE  slt
               CASE  "등록"
                 IF  name  <>  ""  and  uid  <>  ""  and  passwd  <>  ""  Then

						IF  cop =  "01"  THEN fg  =  "1" ELSE fg  =  "2" END IF			
						IF ddate = "1"   THEN ddate =year(now())&"-"&month(now())&"-"&day(now())  ELSE	 ddate = NULL END IF
						IF gubun = "1"  THEN gubun_value = "F0000" ELSE  gubun_value = fac END IF

						syscon.Execute("insert into profile(id, passwd, email, femail, sung, tele, tele1, addr, ddate, hak, fg, stat, sex, iduse, cop, ip, FAC, DEPT, JIK, linkyn, gubun, ws_gubun, mail_db, erp_db,dm_id,ms_id) VALUES ('" & uid & "','" & passwd & "','" & email & "','"& Femail &"',N'" & name & "','" & tele & "','" & tele1 & "','" & addr & "','" & ddate & "','" & hak & "','" & fg & "','" & stat & "','" & sex & "','" & iduse & "','" & cop & "','" & ip & "','" & fac & "','" & dept & "','" & jik & "','N','" & gubun_value & "','" & ws_gubun & "','" & mail & "','" & erp & "','" & dm_id & "','" & ms_id & "')")

                 end if   

		     		
         end select
	
    END IF

 

    %>    

    <BODY CLASS="POPUP_01" SCROLL="no">
    <% CALL Popup_generate("계정등록","등록시 반드시 정확하게 입력하시기 바랍니다.")%>

    <Fieldset align=center style="width:98%;">
    <LEGEND><form name="form" target="" method="post"  action="newuser_wrt.asp">
			<input type="hidden"  name="idx" value="<%=idx%>" >
            <input type="radio" <%=(S2)%> name="slt" value="등록" >등록                
            <input type="radio" <%=(S3)%> name="slt" value="수정" >수정
            <input type="radio" <%=(S4)%> name="slt" value="삭제" >삭제
    </LEGEND>

    <Div class="div_separate_10"></div>

    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="400" >
	    <tr>
            <td class="TBL_Title">아이디</td>
            <td class="TBL_Data"><input type="text" name="uid" maxlength="07"  size="20" onKeyPress="return handleEnter(this, event);" class=INPUT  ></td>     
        </tr>
        <tr>
            <td class="TBL_Title">성 명</td>
            <td class="TBL_Data"><input type="text" name="name"   size="20"  onKeyPress="return handleEnter(this, event);" class=INPUT></td>
        </tr>
        <tr>
            <td class="TBL_Title">비밀번호</td>
            <td class="TBL_Data"><input type="password" name="passwd" size="20"  onKeyPress="return handleEnter(this, event);" class=INPUT>
        </tr>
        <tr>
            <td class="TBL_Title">비밀번호 확인</td>
            <td class="TBL_Data"><input type="password" name="passwd_check"  size="20"   onKeyPress="return handleEnter(this, event);" class=INPUT></td>
        </tr>   
        <tr>
            <td class="TBL_Title">성별</td>
            <td class="TBL_Data"><Input type="radio" name="sex" value="0" <%=A1%> > 남자 <Input type="radio" name="sex" value="1" <%=A2%> > 여자</td>
        </tr>
        <tr>
            <td  class="TBL_Title">부서</td>
            <td class="TBL_Data">
    		
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
    		
		    </td>
        </tr>   
        <tr>
            <td class="TBL_Title">직급</td>
            <td class="TBL_Data">    		    		
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
        <tr>
            <td class="TBL_Title">휴대번호</td>
            <td class="TBL_Data"><input type="text" name="tele1" size="20" onKeyPress="return handleEnter(this, event);" class=INPUT></td>
        </tr>  
        <tr><td class="TBL_Title">내선번호</td>
            <td class="TBL_Data"><input type="text" name="tele" size="20" onKeyPress="return handleEnter(this, event);" class=INPUT></td>
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
    self.resizeTo(500,450)
    //-->
    </SCRIPT>
