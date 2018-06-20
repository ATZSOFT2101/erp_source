    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : ERP 용어 관리                              			        --> 
    <!-- 작성자: 문성원(050407)                                                     --> 
    <!-- 작성일자 : 2006년 7월 20일                                                 --> 
    <!-- 내용: 언어별 ERP 용어 관리													--> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    W_IDX		=  trim(Request("W_IDX"))
    W_KWORD		=  trim(Request("W_KWORD"))
    W_EWORD	    =  trim(Request("W_EWORD")) 
    W_CWORD     =  trim(Request("W_CWORD")) 
    W_RWORD	    =  trim(Request("W_RWORD")) 
    W_GUBUN		=  trim(Request("W_GUBUN"))
    W_CONTENT	=  Request("W_KWORD")
    SLT			=  trim(Request("SLT"))

    IF SLT = "" THEN S2 = "CHECKED" END IF

    SELECT CASE SLT
    CASE ""		 S1 = "CHECKED" : S2 ="disabled='true'" : S3 ="disabled='true'"
    CASE "ADD"  S1 = "CHECKED" : S2 ="disabled='true'" : S3 ="disabled='true'"
    CASE "MOD"  S2 = "CHECKED" : S1 ="disabled='true'" 
    END SELECT

    IF  SLT  <>  ""   THEN

        SELECT CASE  SLT
        CASE  "등록"
			  SQL =" SELECT * FRom SYS_WORD Where W_KWORD ='"& W_KWORD &"' "
			  Set RS=sysCon.execute(SQL)

			  IF Not RS.EOF THEN
				 Alert_Message "같은 파일명으로는 등록되지 않습니다.","2"
			  ELSE
			  
                IF LEN(W_GUBUN) = 3 THEN
                    SQL1 =" SELECT  max(substring(W_CODE,4,4))+1 AS W_CODE_CNT FRom SYS_WORD Where LEFT(RTRIM(W_CODE),3)='"& W_GUBUN &"' "
                ELSE
                    SQL1 = "SELECT max(substring(W_CODE,5,3))+1 AS W_CODE_CNT FRom SYS_WORD Where LEFT(RTRIM(W_CODE),4)='"& W_GUBUN &"' "
                END IF
                 
            '    response.write sql1 
            '    response.end 
                
			    Set RS1=sysCon.execute(SQL1)
			    
			    SELECT CASE Len(Rs1("W_CODE_CNT"))
			    CASE 1
			    W_CODE_CNT ="000"& Rs1("W_CODE_CNT")
			    CASE 2
			    W_CODE_CNT ="00"& Rs1("W_CODE_CNT")
			    CASE 3
			    W_CODE_CNT ="0"& Rs1("W_CODE_CNT")
                CASE 4
                W_CODE_CNT = Rs1("W_CODE_CNT")
			    END SELECT 
			  
			    SQL = "INSERT INTO SYS_WORD (W_CODE,W_KWORD, W_EWORD, W_CWORD,W_RWORD, W_GUBUN, W_CONTENT, W_WDATE, W_EDATE, W_SABUN ) " 
				SQL = SQL & " VALUES ('"& W_GUBUN & W_CODE_CNT &"',N'"& W_KWORD &"',N'"& W_EWORD &"',N'"& W_CWORD &"',N'"& W_RWORD &"','"& W_GUBUN &"',N'"& W_CONTENT &"',GETDATE(),GETDATE(),'"& ID & "')"
				Syscon.EXECUTE(SQL)
              
				 msg =  "등록되었습니다"
			  END IF
			  RS.Close : SET RS=Nothing

        CASE  "수정"
		      SET RS = SYSCON.EXECUTE("SELECT * FROM SYS_WORD WHERE W_IDX= "& W_IDX )

              IF  RS.EOF THEN
				  Alert_Message "수정할 수 없습니다.\n\n등록된 자료가 없습니다.","2"
              ELSE
				  SQL = "UPDATE SYS_WORD set W_KWORD=N'"& W_KWORD &"',W_EWORD=N'"& W_EWORD &"',W_CWORD=N'"& W_CWORD &"',W_RWORD=N'"& W_RWORD &"',"
                  SQL = SQL & " W_GUBUN='"& W_GUBUN &"',W_CONTENT=N'"& W_CONTENT &"',W_EDATE=getdate(),W_SABUN='"& ID &"'"
                  SQL = SQL & " WHERE W_IDX ="& W_IDX
                  syscon.Execute(SQL)

				 msg =  "수정되었습니다"

			  END IF
			  RS.Close : SET RS=Nothing
        CASE "삭제"
		      SET RS = syscon.Execute("SELECT * FROM SYS_WORD WHERE W_IDX= "& W_IDX )

              IF  RS.EOF THEN
				  Alert_Message "삭제할 수 없습니다.\n\n등록된 자료가 없습니다.","2"
              ELSE
				  SYSCON.EXECUTE("DELETE FROM SYS_WORD WHERE W_IDX ="& W_IDX)

				  msg =  "삭제되었습니다"
			  END IF
			  RS.Close : SET RS=Nothing
        END SELECT
               
    END IF

	IF  W_IDX <> "" THEN

		SET RS = SYSCON.EXECUTE("SELECT * FROM SYS_WORD WHERE W_IDX= "& W_IDX )
		IF  NOT ( RS.EOF OR RS.BOF ) THEN
			W_IDX		=  Rs("W_IDX") 
			W_KWORD		=  Rs("W_KWORD") 
			W_EWORD		=  Rs("W_EWORD") 
			W_CWORD		=  Rs("W_CWORD") 
			W_RWORD		=  Rs("W_RWORD") 
			W_GUBUN		=  Rs("W_GUBUN")	
			W_CONTENT	=  Rs("W_CONTENT") 
			W_WDATE		=  Rs("W_WDATE") 
			W_EDATE		=  Rs("W_EDATE") 
			W_SABUN		=  Rs("W_SABUN") 
		END IF
	END IF

    %>
    <BODY CLASS="POPUP_01" SCROLL="no">
    <% CALL Popup_generate("ERP 용어관리","정확하게 입력하셔야만 합니다")%>

    <FIELDSET>
    <LEGEND><form action="sys_word_wrt.asp" method="POST" Name="form">
			<Input type="hidden" name="w_idx" value="<%=w_idx%>" >
			<input type="radio" <%=(S1)%> name="slt" value="등록" >등록                
			<input type="radio" <%=(S2)%> name="slt" value="수정" >수정                
			<input type="radio" <%=(S3)%> name="slt" value="삭제" >삭제              
    </LEGEND>

    <Div class="div_separate_10"></div>

    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%"  style="TABLE-LAYOUT: fixed;" >
    <tr><td ID=TBL_Title1>한글용어명</td> 
        <td ID=TBL_Data1><input type="text" size="35"  name="W_KWORD" class=input style='IME-MODE: active' value="<%=W_KWORD%>"></td> 
    </tr> 
    <tr><td ID=TBL_Title1>영문용어명</td>
        <td ID=TBL_Data1><input type="text" size="35"  name="W_EWORD" class=input value="<%=W_EWORD%>"></td>  
	</tr>
    <tr><td ID=TBL_Title1>중문용어명</td> 
        <td ID=TBL_Data1><input type="text" size="35"  name="W_CWORD" class=input value="<%=W_CWORD%>"></td>
	</tr>
    <tr><td ID=TBL_Title1>러시아용어</td> 
        <td ID=TBL_Data1><input type="text" size="35"  name="W_RWORD" class=input value="<%=W_RWORD%>"></td>
	</tr>
    <tr><td ID=TBL_Title1>업무구분</td>
        <td ID=TBL_Data1><% CALL FAC_slt1("","W_GUBUN","WA",""& W_GUBUN &"") %></td>  
	</tr>
    <!--<tr><td ID=TBL_Title1>용어설명</td> 
        <td ID=TBL_Data1 style="height:80px;"><textarea name="W_CONTENT" style="width:95%;height:80px;" class=input><%=W_CONTENT%></textarea></td>
	</tr>-->
	<tr><TD colspan=2 height=10></td></tR>
	<tr><TD colspan=2>
            <FIELDSET style="width:99%;background-color:White;">
           
            <Div class="div_separate_5"></div>
            
             <table border=0 cellpaddinjg=0 cellspacing=0 width="98%">
             <tr><td class=blue_normal>용어 등록시 주의사항</td></tr>
             <tr><td height=5></td></tr>
             <tr><td>한글 용어명은 반드시 입력하셔야만 합니다.</td></tr>
             <tr><td>영문 용어명의 경우 <span class=red_normal>첫글자는 반드시 대문자로, 두번째 글자부터는 소문자로 입력</span>바랍니다.</td></tr>
             <tr><td>단, <span class=red_bold>P</span>art <span class=red_bold>N</span>umber 와 같이 <span class=red_normal>공백이 있는 경우 공백이후 첫글자는 대문자로</span> 입력바랍니다.</td></tr>
             <tr><td height=5></td></tr>
             </table>
            
            
            </FIELDSET>

	</td></tr>
	</TABLE> 

    </FIELDSET>

    <table border="0" cellpadding="1" cellspacing="1" width="100%">
	    <Tr><Td height=3></td></tr>
            <tr><td><%=msg%></td>
                <td align="right" style="padding-right:15px;"><input type="button" onclick="onsave();" value=" 확 인 " class=btn> <input type="button" value=" 닫 기 " class=btn onclick="self.close();"></td></form>
        </tr>                
    </table>
           

    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
    <SCRIPT language="JavaScript">
    <!--
    //self.moveTo((screen.availWidth-500)/2,(screen.availHeight-540)/2)
    //self.resizeTo(500,540)

	function onsave()    
		{
		if (document.form.W_KWORD.value==""){
		   alert("한글 용어명은 반드시 입력해 주세요.");
		   document.form.W_KWORD.focus();
		   return;
		}
		/*if (document.form.W_CONTENT.value==""){
		   alert("용어설명은 반드시 입력해 주세요.");
		   document.form.W_CONTENT.focus();
		   return;
		} */  
		document.form.submit();    
		}

    //-->
    </SCRIPT>

