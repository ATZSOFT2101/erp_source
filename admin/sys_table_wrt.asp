    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램 : 테이블/필드 관리                                 			        --> 
    <!-- 작 성 자 : 문성원(050407)                                                    --> 
    <!-- 작성일자 : 2006년 6월 08일                                                   --> 
    <!-- 내    용 : 테이블/필드 관리                                                   --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    css1 = " style='font-family:arial;font-size:12px;' "
    SLT			    = TRIM(REQUEST("SLT"))
    RTABLEID		= TRIM(REQUEST("TABLEID")) 
    RTABLEDB		= TRIM(REQUEST("TABLEDB")) 
    RTABLENM		= TRIM(REQUEST("TABLENM")) 
    TABLE_gubun		= TRIM(REQUEST("TABLE_gubun")) 

    IF SLT = "" THEN S2 = "CHECKED" END IF

    SELECT CASE SLT
    CASE ""      S1 = "CHECKED" : S2 ="disabled='true'"  : S3 ="disabled='true'"
    CASE "조회"  S2 = "CHECKED" : S1 ="disabled='true'"
    END SELECT

    IF  SLT  <>  ""   THEN

        SELECT CASE  SLT
        CASE  "조회"
		      SET RS = SYSCON.EXECUTE("SELECT * FROM SYS_TABLE WHERE TABLEID= '"& RTABLEID &"' AND TABLEDB ='"& RTABLEDB &"'" )

		      IF NOT ( RS.EOF OR RS.BOF ) THEN
		         TABLEID	  =  RS("TABLEID") 
		         TABLEDB	  =  RS("TABLEDB") 
		         TABLENM	  =  RS("TABLENM") 
		         TABLE_gubun  =  trim(RS("TABLE_gubun"))
		      END IF

        CASE  "등록"

              SQL = "INSERT INTO SYS_TABLE(TABLEID, TABLEDB, TABLENM,table_gubun) " 
              SQL = SQL & " VALUES ('" & RTABLEID & "','" & RTABLEDB & "','" & RTABLENM & "','" & table_gubun & "')"
              SYSCON.EXECUTE(SQL)
		      
		      msg = "정보가 등록되었습니다."

        CASE  "수정"
		      'SET RS = SYSCON.EXECUTE("SELECT * FROM SYS_TABLE WHERE TABLEID= '"& TABLEID &"' AND TABLEDB ='"& TABLEDB &"' " )
    		  
              SQL = "UPDATE SYS_TABLE SET TABLEID='" & RTABLEID & "', TABLEDB='" & RTABLEDB & "', "
              sql = sql &" table_gubun='" & table_gubun & "', TABLENM='" & RTABLENM & "' WHERE TABLEID ='" & RTABLEID &"' "
		      SYSCON.EXECUTE(SQL)
		      
		      msg = "정보가 수정되었습니다."
		      
        CASE  "삭제"
		      SET RS = SYSCON.EXECUTE("SELECT * FROM SYS_TABLE WHERE TABLEID= '"& RTABLEID &"' AND TABLEDB ='"& RTABLEDB &"' " )

              IF  RS.EOF THEN
		          msg = "삭제할 수 없습니다. 등록된 자료가 없습니다."
              ELSE
                  SYSCON.EXECUTE("DELETE FROM SYS_TABLE WHERE TABLEID= '"& RTABLEID &"' ")
		          SYSCON.EXECUTE("DELETE FROM SYS_FIELD WHERE FLDTABLE= '"& RTABLEID &"' AND FLDDB ='"& RTABLEDB &"' ")
		          msg = "테이블/필드 정보가 삭제되었습니다."
		          
		          
              END IF
              
        END SELECT 
               
    END IF

    %>
    <base target="_self" />
    <BODY CLASS="POPUP_01" SCROLL="no">
    
    <% CALL Popup_generate("테이블 등록","정확하게 입력하셔야만 합니다")%>

    <FIELDSET>
    <LEGEND><FORM name="form" ACTION="SYS_TABLE_WRT.ASP" METHOD="POST">
            <INPUT TYPE=HIDDEN NAME=IDX VALUE="<%=IDX%>">
			<INPUT TYPE="RADIO" <%=(S1)%> NAME="SLT"  VALUE="등록">등록
			<INPUT TYPE="RADIO" <%=(S2)%> NAME="SLT"  VALUE="수정">수정 
			<INPUT TYPE="RADIO" <%=(S3)%> NAME="SLT"  VALUE="삭제">삭제                 
    </LEGEND>

    <DIV CLASS="DIV_SEPARATE_10"></DIV>

    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%"  STYLE="TABLE-LAYOUT: FIXED;" >
    <TR><TD ID=TBL_TITLE1>데이타베이스</TD> 
        <TD ID=TBL_DATA1>
            <Table cellpadding=0 cellspacing=0 border=0>
            <tr><td><SELECT NAME=TABLEDB <%=css1%>><%
			             SQL="SELECT NAME FROM MASTER.DBO.SYSDATABASES	"
			             SET MASTERRS=CONN.EXECUTE(SQL)

			             IF NOT (MASTERRS.EOF OR MASTERRS.BOF) THEN
				            WHILE NOT MASTERRS.EOF

			                IF RTABLEDB = Trim(MASTERRS("NAME")) THEN
				               RESPONSE.WRITE "<OPTION VALUE='"& RTABLEDB &"' SELECTED>"& trim(RTABLEDB) &"</OPTION>"
				            ELSE
				               RESPONSE.WRITE "<OPTION VALUE='"& MASTERRS("NAME") &"'>"& trim(MASTERRS("NAME")) &"</OPTION>"
				            END IF

				            MASTERRS.MOVENEXT
				            WEND
				            
				            MASTERRS.close
			             END IF
			             %></SELECT></td>
                <td  <%=css1%>><% CALL FAC_slt1("","table_gubun","WA",""& table_gubun &"") %></td>
            </tr></Table>                    
                    </TD> 
    </TR> 
    <TR><TD ID=TBL_TITLE1>테이블명</TD>
        <TD ID=TBL_DATA1><INPUT TYPE="TEXT" <%=css1%> NAME="TABLEID" CLASS=INPUT ONKEYPRESS="return formenter(this,event,1);" VALUE="<%=TABLEID%>" style="width:90%;"></TD>
    </TR>
    <TR><TD ID=TBL_TITLE1>테이블설명</TD>
        <TD ID=TBL_DATA1><INPUT TYPE="TEXT" <%=css1%> NAME="TABLENM" CLASS=INPUT  VALUE="<%=TABLENM%>" style="width:90%;"></TD>
    </TR>
    
    <Td></td>
    </TABLE>
 
    </FIELDSET>

    <TABLE BORDER="0" CELLPADDING="1" CELLSPACING="1" WIDTH="100%">
	    <TR><TD ALIGN="CENTER" height=5 colspan=2></TD></TR>
        <TR><td style="padding-left:10px;color:Blue;"><%=msg%></td>
            <TD ALIGN="RIGHT" STYLE="PADDING-RIGHT:15PX;">
            <INPUT TYPE="SUBMIT"  VALUE=" 확 인 " CLASS=BTN> 
            <INPUT TYPE="BUTTON" VALUE=" 닫 기 " CLASS=BTN ONCLICK="self.close();"></TD></FORM>
        </TR>                
    </TABLE>
       

    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
