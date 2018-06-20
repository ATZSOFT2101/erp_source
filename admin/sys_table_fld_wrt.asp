    <!--#include Virtual = "/INC1/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램 : 테이블/필드 관리                                 			    --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 08일                                                 --> 
    <!-- 내    용 : 필드 추가 프로그램                                              --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC1/INC_BODY.ASP" -->
    <%
    slt			= trim(Request("slt"))

    FLD_idx		=  Request("FLD_idx")
    FLDDB		=  trim(Request("FLDDB")) 
    FLDtable	=  trim(Request("FLDtable")) 
    FLDName		=  trim(Request("FLDName") )
    FLDTYPE		=  trim(Request("FLDTYPE")) 
    FLDLENG		=  Request("FLDLENG") 
    FLDDESC		=  trim(Request("FLDDESC")) 
    FLDNULL		=  trim(Request("FLDNULL"))
    FLDDisplay	=  trim(Request("FLDDisplay"))

    if slt = "" then S1 = "checked" end if

    select case slt
    case "등록"  S1 = "checked"
    case "수정"  S2 = "checked"
    case "조회"  S2 = "checked"
    case "삭제"  S3 = "checked"
    end select

    if  slt  <>  ""   Then

    select case  slt
    case  "조회"
		  Set Rs = syscon.Execute("select * from sys_field where FLD_idx= "& FLD_idx   )

		  IF Not ( RS.EOF or RS.BOF ) then
			 FLD_idx		=  Rs("FLD_idx")
			 FLDDB			=  Rs("FLDDB") 
			 FLDtable		=  Rs("FLDtable") 
			 FLDName		=  Rs("FLDName") 
			 FLDNO			=  Rs("FLDNO") 
			 FLDTYPE		=  Rs("FLDTYPE") 
			 FLDLENG		=  Rs("FLDLENG") 
			 FLDDESC		=  Rs("FLDDESC") 
			 FLDNULL		=  Rs("FLDNULL")  
			 FLDDisplay		=  Rs("FLDDisplay")
		  End if

    case  "등록"
		  Set Rs = syscon.Execute("select IsNULL(FLDno,0) from sys_field where FLDDB= '"& FLDDB &"' and FLDtable= '"& FLDtable &"' and FLDName= '"& FLDName &"'")
		  if Rs("FLDno")=0 tHEN
		  FLDNO = 1
		  ELSE
		  FLDNO = Rs("FLDno")+1
		  END IF

          Sql = "insert into sys_field(FLDDB, FLDtable, FLDName,FLDNO,FLDTYPE,FLDLENG,FLDDESC,FLDNULL,FLDDisplay) " 
          Sql = Sql & " VALUES ('" & FLDDB & "','" & FLDtable & "','" & FLDName & "'," & FLDNO & ",'" & FLDTYPE & "'," & FLDLENG & ",'" & FLDDESC & "','" & FLDNULL & "','"& FLDDisplay &"')"
          syscon.Execute(Sql)
          
		  Alert_message "정보가 등록되었습니다.","2"

    case  "수정"
          Sql = "update sys_field set FLDDB='" & FLDDB & "', FLDtable='" & FLDtable & "', FLDName='" & FLDName & "', FLDTYPE='" & FLDTYPE & "', FLDLENG=" & FLDLENG &", FLDDESC='" & FLDDESC & "', FLDNULL='" & FLDNULL & "', FLDDisplay='" & FLDDisplay & "' where FLD_idx= "& FLD_idx 
		  syscon.Execute(Sql)
		  
		  Alert_message "정보가 수정되었습니다.","2"
    case  "삭제"
		  Set Rs = syscon.Execute("select * from sys_field where FLD_idx= "& FLD_idx  )

          if  Rs.Eof then
		  Alert_message "삭제할 수 없습니다. 등록된 자료가 없습니다.","2"
          else
          syscon.Execute("delete from sys_field where FLD_idx= "& FLD_idx )
		  Alert_message "정보가 삭제되었습니다.","2"
          end if
    end select 
           
    end if

    %>
    <base target="_blank" />
    <BODY CLASS="POPUP_01" SCROLL="no">
    
    <% CALL Popup_generate("필드 등록","정확하게 입력하셔야만 합니다")%>

    <FIELDSET>
    <LEGEND><FORM ACTION="sys_table_fld_wrt.ASP" METHOD="POST">
			<Input type=hidden name=FLD_idx value=<%=FLD_idx%>>
			<INPUT TYPE="RADIO" <%=(S1)%> NAME="SLT"  VALUE="등록">등록
			<INPUT TYPE="RADIO" <%=(S2)%> NAME="SLT"  VALUE="수정">수정 
			<INPUT TYPE="RADIO" <%=(S3)%> NAME="SLT"  VALUE="삭제">삭제                 
    </LEGEND>

    <DIV CLASS="DIV_SEPARATE_10"></DIV>

    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%"  STYLE="TABLE-LAYOUT: FIXED;" >
       <tr>
          <td ID=TBL_TITLE1>DB명</td>
	      <td ID=TBL_Data1><select name=FLDDB><%
	      sql="SELECT name FROM master.dbo.SYSDATABASES	"
	      set masterRs=conn.execute(sql)

	      IF Not (masterRs.EOF or masterRs.BOF) then
		     While Not masterRs.EOF

		         IF UCASE(FLDDB) = UCASE(masterRs("name")) then
			        Response.write "<option value='"& FLDDB &"' selected>"& FLDDB &"</option>"
			     else
			        Response.write "<option value='"& masterRs("name") &"' >"& masterRs("name") &"</option>"
			     end if

		     masterRs.MoveNext
		     Wend
	      end if
	      %></td>
       </tr>           
       <tr>
          <td ID=TBL_TITLE1>테이블명</td>
          <td ID=TBL_Data1><input type="text" name="FLDTable" class=input onKeyPress="return handleEnter(this, event);" value="<%=FLDTable%>" style="width:95%;"></td>
       </tr>
       <tr>              
          <td ID=TBL_TITLE1>필드명</td>
          <td ID=TBL_Data1><input type="text" name="FLDName"  class=input onKeyPress="return handleEnter(this, event);" value="<%=FLDName%>" style="width:95%;"></td>
       </tr>
       <tr>              
          <td ID=TBL_TITLE1>필드타입</td>
          <td ID=TBL_Data1><select name="FLDType" style="width:95%;">
		      <% IF FLDType <> "" then %>
	          <option value="<%=FLDType%>" selected><%=FLDType%></option>
		      <% end if	%>
		      <option value="varchar">varchar</option>
		      <option value="numeric">numeric</option>
		      <option value="uniqueidentifier">uniqueidentifier</option>
		      <option value="bigint">bigint</option>
		      <option value="binary">binary</option>
		      <option value="bit">bit</option>
		      <option value="char">char</option>
		      <option value="datetime">datetime</option>
		      <option value="decimal">decimal</option>
		      <option value="float">float</option>
		      <option value="image">image</option>
		      <option value="int">int</option>
		      <option value="money">money</option>
		      <option value="nchar">nchar</option>
		      <option value="ntext">ntext</option>
		      <option value="nvarchar">nvarchar</option>
		      <option value="real">real</option>
		      <option value="smalldatetime">smalldatetime</option>
		      <option value="smallint">smallint</option>
		      <option value="smallmoney">smallmoney</option>
		      <option value="sql_variant">sql_variant</option>
		      <option value="text">text</option>
		      <option value="timestamp">timestamp</option>
		      <option value="tinyint">tinyint</option>
		      <option value="varbinary">varbinary</option>
		      </select></td>
       </tr>
       <tr>              
          <td ID=TBL_TITLE1>자리수</td>
          <td ID=TBL_Data1><input type="text" name="FLDLeng" class=input onKeyPress="return handleEnter(this, event);" value="<%=FLDLeng%>" size=3></td>
       </tr>
       <tr>              
          <td ID=TBL_TITLE1>NULL허용</td>
          <td ID=TBL_Data1><Input type="radio" name="FLDNULL" <% IF FLDNull="1" or FLDNull ="" then %>checked<% end if%> value=1> 허용 <Input type="radio" name="FLDNULL" <% IF FLDNull="0" then %>checked<% end if%> value=0> 허용하지 않음</td>
       </tr>
       <tr>              
          <td ID=TBL_TITLE1>출력허용</td>
          <td ID=TBL_Data1><Input type="radio" name="FLDDisplay" <% IF FLDDisplay="1" or FLDDisplay ="" then %>checked<% end if%> value=1> 허용 <Input type="radio" name="FLDDisplay" <% IF FLDDisplay="0" then %>checked<% end if%> value=0> 허용하지 않음</td>
       </tr>           
       <tr>              
          <td ID=TBL_TITLE1>필드설명</td>
          <td ID=TBL_Data1><input type="text" name="FLDDESC" size="40"  class=input  value="<%=FLDDESC%>"></td>
       </tr>
    </table>

    </fieldset>

    <table border="0" cellpadding="1" cellspacing="1" width="100%">
       <Tr><Td align="center" height=5></td></tr>
       <tr><td ID=TBL_TITLE1 style="padding-right:15px;"><input type="submit"  value=" 확 인 " class=btn> <input type="button" value=" 닫 기 " class=btn onclick="self.close();"></td></form>
       </tr>                
    </table>

    </div>

</BODY>
</HTML>