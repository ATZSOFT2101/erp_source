    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 전자결재 문서코드 등록                              			--> 
    <!-- 작성자: 문성원(050407)                                                     --> 
    <!-- 작성일자 : 2006년 6월 19일                                                 --> 
    <!-- 내용: 전자결재 문서코드 관리                                               --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    ID						=  Request("ID")
    dcode					=  trim(Request("dcode")) 
    dname					=  trim(Request("dname"))  
    dpart					=  trim(Request("dpart"))  
    pname					=  trim(Request("pname"))  
    pname1					=  trim(Request("pname1"))  
    usechk					=  trim(Request("usechk"))  

    '****** 2006년 02월 8일 FIELD 추가 작업 *************
    doc_type1				=  trim(Request("doc_type1"))  
    doc_display				=  trim(Request("doc_display"))  
    doc_readline			=  trim(Request("doc_readline"))  
    doc_circulation			=  trim(Request("doc_circulation"))  
    doc_fileupload	 		=  trim(Request("doc_fileupload"))  
    doc_fac					=  trim(Request("doc_fac"))
    doc_dept				=  trim(Request("doc_dept"))
    doc_subject				=  trim(Request("doc_subject"))  
    doc_dbname				=  trim(Request("doc_dbname"))  
    doc_width				=  trim(Request("doc_width"))  
    doc_pcode				=  trim(Request("doc_pcode"))  
    slt						=  trim(Request("slt"))

    IF doc_dept  > "" THEN
       temp			=  split(doc_dept,",")
       dpart			=  temp(0)
       doc_dept		=  temp(1)
    END IF

    IF slt = "" Then S2 = "checked" END IF

    SELECT CASE slt
    CASE "조회"  S1 = "checked"
    CASE "등록"  S2 = "checked"
    CASE "수정"  S3 = "checked"
    CASE "삭제"  S4 = "checked"
    END SELECT

    IF slt  <>  ""   Then

        SELECT CASE  slt
        CASE  "조회"
		      Set Rs = Dbcon.Execute("select ID, dcode, dname, dpart, pname, pname1, usechk, IsNULL(doc_type1,1) AS doc_type1, IsNULL(doc_display,1) AS doc_display, IsNULL(doc_readline,1) AS doc_readline, IsNULL(doc_circulation,1) AS doc_circulation, IsNULL(doc_fileupload,1) AS doc_fileupload ,doc_fac, doc_dept, doc_dbname, IsNULL(doc_subject,1) AS doc_subject, doc_width, doc_pcode from docmst where id= "& id )

		      IF Not ( RS.EOF or RS.BOF ) then
			     ID					=  Rs("ID")
			     dcode				=  trim(Rs("dcode")) 
			     dname				=  trim(Rs("dname"))  
			     dpart				=  trim(Rs("dpart"))  
			     pname				=  trim(Rs("pname"))  
			     pname1				=  trim(Rs("pname1"))  
			     usechk				=  trim(Rs("usechk"))  
			     doc_type1			=  trim(Rs("doc_type1"))  
			     doc_display		=  trim(Rs("doc_display"))  
			     doc_readline		=  trim(Rs("doc_readline"))  
			     doc_circulation	=  trim(Rs("doc_circulation"))  
			     doc_fileupload		=  trim(Rs("doc_fileupload"))  
			     doc_fac			=  trim(Rs("doc_fac"))  
			     doc_dept			=  trim(Rs("doc_dept"))  
			     doc_dbname		    =  trim(Rs("doc_dbname"))
			     doc_subject		=  trim(Rs("doc_subject"))
			     doc_width			=  Rs("doc_width")
			     doc_pcode			=  trim(RS("doc_pcode"))
		      End if
    			
        CASE  "등록"
		      Set chkRs = Dbcon.Execute("select sum=count(*) from docmst where dcode= '"& dcode &"' " )

		      IF   chkRs("sum") = 0  then
			      Sql = "insert into docmst(dcode, dname, dpart, pname, pname1, usechk, doc_type1, doc_display, doc_readline, doc_circulation, doc_fileupload, doc_fac, doc_dept, doc_dbname, doc_subject, doc_width,doc_pcode) " 
			      Sql = Sql & " VALUES ('" & dcode & "','" & dname & "','" & dpart & "','" & pname & "','"& pname1 & "','"& usechk & "','"& doc_type1 & "','"& doc_display & "','"& doc_readline & "', '"& doc_circulation & "','"& doc_fileupload & "','"& doc_fac & "','"& doc_dept & "','"& doc_dbname & "','"& doc_subject & "',"& doc_width &",'"& doc_pcode & "')"                     
			      dbcon.Execute(Sql)

		          Alert_message "등록되었습니다.","2"
		      ELSE
		          Alert_message "중복된 코드가 있습니다.\n\n문서코드를 확인해주세요.","2"
		      END IF

        CASE  "수정"

		        Sql = "update docmst set dcode='"& dcode &"',dname='" & dname & "',dpart='" & dpart & "',"
		        Sql = Sql & " pname='" & pname & "',pname1='"& pname1 &"', usechk='"& usechk &"' , "
		        Sql = Sql & " doc_type1='" & doc_type1 & "',doc_display='"& doc_display &"', doc_readline='"& doc_readline &"' , "
		        Sql = Sql & " doc_circulation='" & doc_circulation & "',doc_fileupload='"& doc_fileupload &"', doc_fac='"& doc_fac &"' , "
		        Sql = Sql & " doc_dept='" & doc_dept & "',doc_dbname='"& doc_dbname &"',doc_subject='"& doc_subject &"',doc_width="& doc_width &", "
		        Sql = Sql & " doc_pcode='"& doc_pcode &"' where id =" & id
		        dbcon.Execute(Sql)
		        
		        Alert_message "수정되었습니다.","2"
        END SELECT 
               
    END IF

    SELECT CASE usechk
    CASE ""
    chk_01 = "checked"
    CASE "Y"
    chk_01 = "checked"
    CASE "N"
    chk_02 = "checked"
    END SELECT

    SELECT CASE doc_type1
    CASE ""
    doc_type1_01 = "checked"    
    CASE "1"
    doc_type1_01 = "checked"
    CASE "2"
    doc_type1_02 = "checked"
    CASE "3"
    doc_type1_03 = "checked"
    CASE "4"
    doc_type1_04 = "checked"
    END SELECT

    SELECT CASE doc_display
    CASE ""
    doc_display_01 = "checked"
    CASE "1"
    doc_display_01 = "checked"
    CASE "2"
    doc_display_02 = "checked"
    END SELECT

    SELECT CASE doc_readline
    CASE ""
    doc_readline_01 = "checked"
    CASE "1"
    doc_readline_01 = "checked"
    CASE "2"
    doc_readline_02 = "checked"
    END SELECT

    SELECT CASE doc_fileupload
    CASE ""
    doc_fileupload_01 = "checked"
    CASE "1"
    doc_fileupload_01 = "checked"
    CASE "2"
    doc_fileupload_02 = "checked"
    END SELECT

    SELECT CASE doc_circulation
    CASE ""
    doc_circulation_01 = "checked"
    CASE "1"
    doc_circulation_01 = "checked"
    CASE "2"
    doc_circulation_02 = "checked"
    END SELECT

    SELECT CASE trim(doc_dbname)
    CASE ""
    doc_dbname_01 = "checked"
    CASE "mail"
    doc_dbname_01 = "checked"
    CASE "dongaerp"
    doc_dbname_02 = "checked"
    CASE "popda"
    doc_dbname_03 = "checked"
    END SELECT

    SELECT CASE doc_subject
    CASE ""
    doc_subject_01 = "checked"
    CASE "1"
    doc_subject_01 = "checked"
    CASE "2"
    doc_subject_02 = "checked"
    END SELECT
    %>
    <BODY CLASS="POPUP_01" SCROLL="no">
    <% CALL Popup_generate("결재코드 등록","정확하게 입력하셔야만 합니다")%>


    <FIELDSET style="width:98%;">
    <LEGEND><form action="sys_appro_code_wrt.asp" method="POST">
			<Input type="hidden" name="id" value="<%=id%>" >
			<input type="radio" <%=(S1)%> name="slt" value="조회" >조회
			<input type="radio" <%=(S2)%> name="slt" value="등록" >등록                
			<input type="radio" <%=(S3)%> name="slt" value="수정" >수정
    </LEGEND>

    <Div class="div_separate_10"></div>

    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%"  >
    <tr><td ID=TBL_Title1>결재코드</td> 
        <td ID=TBL_Data1>		  
			  <Table border=0 cellpadding=0 cellspacing=0 width="100%">
			  <Tr><Td width="40%"><input type="text" size="35"  name="dcode" value="<%=dcode%>" class=input style="width:80%;"></td>
				  <Td width="20%">문서폭(PX)</td>
				  <Td width="40%"><input type="text" size="10"  name="doc_width" <% IF doc_width <> "" THEN %> value="<%=doc_width%>" <% ELSE %> value="700" <% END IF %>class=input> Pixel</td>
			  </tr>
			  </table>
		  </td> 
    </tr> 
    <tr><td ID=TBL_Title1>양식번호</td>
        <td ID=TBL_Data1><input type="text" size="35"  name="doc_pcode" value="<%=doc_pcode%>" class=input style="width:90%;"></td>  
	</tr>
    <tr><td ID=TBL_Title1>문서명</td>
        <td ID=TBL_Data1><input type="text" size="35"  name="dname" value="<%=dname%>" class=input style="width:90%;"></td>  
	</tr>
    <tr><td ID=TBL_Title1>사용공장/부서</td> 
        <td ID=TBL_Data1><select name="doc_Fac" style="width:44%;">
		<%
			SQL = "SELECT * FROM ADMIN_FAC_GUBUN" 
			Set facRs = syscon.Execute(SQL)
			%>
			<option value="">선택하세요</option>
			<%
			Do until facRs.EOF
			fac_nm = trim(facRs("fac_nm"))
			fac_cd = trim(facRs("fac_cd"))

			IF trim(doc_fac)=fac_cd then
			%>
			<option value="<%=trim(fac_cd)%>" selected><%=trim(fac_nm)%></option>
			<%
			else
			%>
			<option value="<%=trim(fac_cd)%>"><%=trim(fac_nm)%></option>
			<%
			end if
			facRs.Movenext
			Loop
			%>
			</select>
			<%
			facRs.close
			Set facRs = Nothing
		    %>
	        <select name="doc_dept" style="width:45%;">
			<%
			SQL = "SELECT * FROM ADMIN_DEPT_GUBUN" 
			Set deptRs = syscon.Execute(SQL)

			Do until deptRs.EOF
			dept_nm = trim(deptRs("dept_nm"))
			dept_cd  = trim(deptRs("dept_cd"))

			IF trim(doc_dept)=dept_cd then
			%>
			<option value="<%=trim(dept_nm)%>,<%=trim(dept_cd)%>" selected><%=trim(dept_nm)%></option>
			<%
			else
			%>
			<option value="<%=trim(dept_nm)%>,<%=trim(dept_cd)%>"><%=trim(dept_nm)%></option>
			<%
			end if
			deptRs.Movenext
			Loop
			%>
			</select>
			<%
			deptRs.close
			Set deptRs = Nothing
		%>
	</tr>
    <tr><Td colspan=2 align=center><hr style="width:98%;"></td></tr>
    <tr><td ID=TBL_Title1>문서타입 I</td>
        <td height=60>
		  <Table border=0 cellpadding=0 cellspacing=0 width="100%">
		  <Tr><Td width="25%"><input type=radio value="1" <%=doc_type1_01%> name="doc_type1"> 웹에디터</td>	
			  <Td width="25%"><input type=radio value="2" <%=doc_type1_02%> name="doc_type1"> 기본쿼리</td>
			  <Td width="25%"><input type=radio value="3" <%=doc_type1_03%> name="doc_type1"> 자재타입</td>
			  <Td width="25%"><input type=radio value="4" <%=doc_type1_04%> name="doc_type1"> 연구타입</td>
		  </tr>
		  <Tr><Td style="padding-left:8px;"><img src="<%=img_path%>doc_type_01.gif" border=0></td>	
			  <Td style="padding-left:8px;"><img src="<%=img_path%>doc_type_02.gif" border=0></td>	
			  <Td style="padding-left:8px;"><img src="<%=img_path%>doc_type_03.gif" border=0></td>	
			  <Td style="padding-left:8px;"><img src="<%=img_path%>doc_type_04.gif" border=0></td>	
		  </tr>
		  </table>
	  </td>  
    </tr>
    <tr><Td colspan=2 align=center><hr style="width:98%;"></td></tr>
    <tr><td ID=TBL_Title1>결재라인</td>
        <td ID=TBL_Data1>
		  <Table border=0 cellpadding=0 cellspacing=0 width="100%">
		  <Tr><Td width="33%"><input type=radio value="1" <%=doc_readline_01%> name="doc_readline"> 결재라인I</td>
			  <Td width="33%"><input type=radio value="2" <%=doc_readline_02%> name="doc_readline"> 결재라인II</td>
			  <Td width="33%"></td>
		  </tr>
		  </table>
	  </td>  
    </tr>
    <tr><td ID=TBL_Title1>첨부파일</td>
        <td ID=TBL_Data1>
		  <Table border=0 cellpadding=0 cellspacing=0 width="100%">
		  <Tr><Td width="33%"><INPUT TYPE="radio" NAme="doc_fileupload"  value="1" <%=doc_fileupload_01%>> 첨부가능  </td>
			  <Td width="33%"><INPUT TYPE="radio" NAme="doc_fileupload"  value="2" <%=doc_fileupload_02%>> 첨부불가</td>
			  <Td width="33%"></td>
		  </tr>
		  </table>
	   </td>  
	</tr>
    <tr><td ID=TBL_Title1>회람여부</td>
        <td ID=TBL_Data1>
		  <Table border=0 cellpadding=0 cellspacing=0 width="100%">
		  <Tr><Td width="33%"><INPUT TYPE="radio" name="doc_circulation" value="1"  <%=doc_circulation_01%>> 회람지정 가능  </td>
			  <Td width="33%"><INPUT TYPE="radio" name="doc_circulation" value="2"  <%=doc_circulation_02%>> 회람지정 불가</td>
			  <Td width="33%"></td>
		  </tr>
		  </table>
	   </td>  
	</tr>
    <tr><td ID=TBL_Title1>배포여부</td>
        <td ID=TBL_Data1>
		  <Table border=0 cellpadding=0 cellspacing=0 width="100%">
		  <Tr><Td width="33%"><INPUT TYPE="radio" name="doc_display" value="1"  <%=doc_display_01%>> 부서배포 가능  </td>
			  <Td width="33%"><INPUT TYPE="radio" name="doc_display" value="2"  <%=doc_display_02%>> 부서배포 불가</td>
			  <Td width="33%"></td>
		  </tr>
		  </table>
	   </td>  
	</tr>
    <tr><td ID=TBL_Title1>제목표시여부</td>
        <td ID=TBL_Data1>
		  <Table border=0 cellpadding=0 cellspacing=0 width="100%">
		  <Tr><Td width="33%"><INPUT TYPE="radio" name="doc_subject" value="1"  <%=doc_subject_01%>> 수동입력</td>
			  <Td width="33%"><INPUT TYPE="radio" name="doc_subject" value="2"  <%=doc_subject_02%>> 자동입력</td>
			  <Td width="33%"></td>
		  </tr>
		  </table>
	    </td>  
	</tr>
    <tr><td ID=TBL_Title1>참조DB명</td>
        <td ID=TBL_Data1>
		  <Table border=0 cellpadding=0 cellspacing=0 width="100%">
		  <Tr><Td width="33%"><input type=radio value="mail" <%=doc_dbname_01%> name="doc_dbname"> 참조없슴</td>
			  <Td width="33%"><input type=radio value="dongaerp" <%=doc_dbname_02%> name="doc_dbname"> 업무DB</td>
			  <Td width="33%"><input type=radio value="popda" <%=doc_dbname_03%> name="doc_dbname"> POPDA</td>
		  </tr>
		  </table>
		</td>  
	</tr>
    <tr><td ID=TBL_Title1>사용구분</td>
        <td ID=TBL_Data1>
		  <Table border=0 cellpadding=0 cellspacing=0 width="100%">
		  <Tr><Td width="33%"><INPUT TYPE="radio" name="usechk" value="Y" <%=chk_01%>> 정상</td>
			  <Td width="33%"><INPUT TYPE="radio" name="usechk" value="N" <%=chk_02%>> 불가</td>
			  <Td width="33%"></td>
		  </tr>
		  </table>
	   </td>  
	</tr>
	</table> 
	
    </FIELDSET>

    <table border="0" cellpadding="1" cellspacing="1" width="100%">
	    <Tr><Td align="center" height=10></td></tr>
        <tr><td align="right" style="padding-right:15px;"><input type="submit"  value=" 확 인 " class=btn> <input type="button" value=" 닫 기 " class=btn onclick="self.close();"></td></form></tr>                
    </table>

    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
    <SCRIPT language="JavaScript">
    <!--
    self.moveTo((screen.availWidth-500)/2,(screen.availHeight-600)/2)
    self.resizeTo(500,600)
    //-->
    </SCRIPT>
