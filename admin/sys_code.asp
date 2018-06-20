    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 공통코드 관리                              						--> 
    <!-- 내용: gubun1 테이블의 공통 코드 관리											--> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
	GUBUN_CODE = REQUEST("GUBUN_CODE")
	SLT        = REQUEST("SLT")
	RNVcnt     =   REQUEST("RNVcnt")
	
	
	SELECT CASE  SLT
    CASE "등록"
    
	     sql="select * from gubun1 where gubun_code='"& request("gubun_code") &"' or gubun_title=N'"& request("gubun_title") &"' "
	     set comRs=conn.execute(sql)

	     IF comRs.EOF or comRs.BOF then
            sql = "INSERT INTO COMCODH(COMMCD,COMMNM,LENGTH,CDRM01,CDRM02,CDRM03,CDRM04,CDRM05,CDRM06,CDRM07,CDRM08,CDRM09,CDRM10, " _
                                                     +    "CDRC01,CDRC02,CDRC03,CDRC04,CDRC05,CDRC06,CDRC07,CDRC08,CDRC09,CDRC10, " _
                                                     +    "CDCK01,CDCK02,CDCK03,CDCK04,CDCK05,CDCK06,CDCK07,CDCK08,CDCK09,CDCK10, " _
                                                     +    "CDLE01,CDLE02,CDLE03,CDLE04,CDLE05,CDLE06,CDLE07,CDLE08,CDLE09,CDLE10, " _
                                                     +    "FROMDATE,TODATE,CODEKIND,FDATE,FUSER,FFORM,LDATE,LUSER,LFORM) " _
               + "VALUES('" & request("gubun_code") & "',N'" & request("gubun_title") & "',3,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'20110101','29991231','E',getdate(),'921116','ADMIN',getdate(),'921116','ADMIN')"
		    conn.Execute(sql)
	     ELSE
		    sql = "UPDATE comcodh SET commnm = N'" & request("gubun_title") & "'  WHERE  commcd = '" & request("gubun_code") & "'"
		    conn.Execute(sql)

	     '   Alert_message "코드값이 같거나 코드명이 같은 레코드가 있습니다.","1"
	     END IF
	     comRs.Close : set comRs=nothing
    case "삭제"
    	 sql="select * from comcodd where commcd ='"& request("gubun_code") &"'"
	     set comRs=conn.execute(sql)

         IF comRs.EOF  then
            sql = "DELETE comcodh WHERE commcd = '" & REQUEST("GUBUN_CODE") & "'"
	        conn.Execute(sql)
         else
            Alert_message "Sub Code를 먼저 삭제 해주세요","1"
         end if
    CASE "등록1"
    
    	 sql="select * from gubun1 where gubun_code='"& request("gubun_code") &"' and gubun_title=N'"& request("gubun_title") &"' "
	     set comRs=conn.execute(sql)
         
         IF comRs.EOF or comRs.BOF THEN '신규등록
            Alert_message "Master Code를 먼저 등록해주세요","1"
         ELSE
         
            Set rsCnt = conn.Execute( "SELECT maxSEQ=max(GUBUN_SEQ) FROM GUBUN1 WHERE GUBUN_CODE = '"& REQUEST("GUBUN_CODE") &"'  ")
            if rsCnt("maxSEQ") = 0 then maxSEQ = 1 else maxSEQ = rsCnt("maxSEQ") + 1 
            rsCnt.Close  : Set rsCnt = Nothing         
         
            sql =  "INSERT INTO COMCODD(COMMCD,COMDCD,COMDNM,CDREF01,CDREF02,CDREF03,CDREF04,CDREF05,CDREF06,CDREF07,CDREF08,CDREF09,CDREF10,CDREFRM,CDSORT,FDATE,FUSER,FFORM,LDATE,LUSER,LFORM) " _
                   + " VALUES('" & request("gubun_code") & "',"& maxSEQ &",N'"& REQUEST("GUBUN_NAME") &"',0,0,0,0,0,0,0,0,0,0,0,1,GETDATE(),'921116','admin',GETDATE(),'921116','admin') "
            conn.Execute( sql )
         END IF
         
    CASE "수정1"
         
            sql = "UPDATE comcodd " & _
             "   SET comdnm = N'" & REQUEST("GUBUN_NAME") & "' " & _
             " WHERE commcd = '" & REQUEST("GUBUN_CODE") & "' and comdcd = '"& REQUEST("GUBUN_SEQ") & "' "
            conn.Execute( sql )
            
    CASE "삭제1"         
            sql = "DELETE comcodd WHERE commcd = '" & REQUEST("GUBUN_CODE") & "' and comdcd='"& REQUEST("GUBUN_SEQ") & "' "           
            conn.Execute( sql )
            
	END SELECT

	'SQL = "SELECT distinct GUBUN_CODE, GUBUN_TITLE FROM gubun1 "
  
     SQL = " SELECT COMMCD, COMMNM, LENGTH, FROMDATE, TODATE,"
     SQL = SQL & "        Case When Upper(CodeKind) = 'S'"
     SQL = SQL & "             then '시스템 코드'"
     SQL = SQL & "             else '사용자 코드' end,"
     SQL = SQL & "        Case When Upper(CodeKind) = 'S'"
     SQL = SQL & "             then 1"
     SQL = SQL & "             else 0 end"
     SQL = SQL & "   FROM COMCodH "
   
 '  If txtCommcd1.Text & "" <> "" Then
 '     If cboFind.ListIndex = 0 Then
 '        GS_SQL = GS_SQL & " Where CoMMCD >= '" & txtCommcd1.Text & "'"
 '     Else
 '        GS_SQL = GS_SQL & " Where CommNM like '%" & txtCommcd1.Text & "%'"
 '     End If
 '  End If
   
    SQL = SQL & " ORDER BY COMMCD"
	set RS=conn.execute(sql)
    %>    
	<SCRIPT LANGUAGE="JavaScript" TYPE="TEXT/JAVASCRIPT">
	<!--	
	<%
	Dim x : x = Request("x")
	Dim y : y = request("y")
	If x="" Then x=0 End If
	If y="" Then y=0 End If
	%>
	
	function window.onload()
	{
		scroll_Focus("x",<%= x %>,0);
		scroll_Focus("y",0,<%= y %>);

		resize_scroll(70);
		
		<% IF RNVcnt <> "" THEN %>
		sys_code('<%=SLT%>','<%=GUBUN_CODE%>','<%=COMMNM%>','<%=RNVcnt%>') 
		<% END IF %>
	}
    
	//-->
	</SCRIPT>
	<script type="text/javascript">
	    $(document).ready(function() {
	        $("#sub_table TR").live("click",function() {
	            var nodes = $(this).children();
	        //    alert(nodes.length);
	            nodes.each(function() {
	                $(this).removeClass();
	            });
	        });
	    });
	</script>
	<BODY CLASS="BODY_01" SCROLL="no" >

    <!-- HEAD 부분 시작 --> 
    <TABLE CLASS=Sch_Header id=d2>
	  <TR><TD width="50%">
		          <!-- 마스터 코드 등록 부분 시작 -->
		          <Table border=0 cellpadding=3 cellspacing=0>
		          <TR><form name=form action="sys_code.asp" method="post">
		              <input type=hidden name="SLT" value="등록" />
		              <input TYPE="HIDDEN" NAME="old_nvcnt" >
		              <td height=26 align=center style="padding-left:5px;">코드</td>
		              <td ><input type="text" name="GUBUN_CODE" size="15" maxlength="2"  class=input ></td>
		              <td height=26 align=center style="padding-left:5px;">코드타이틀</td>
		              <td ><input type="text" name="GUBUN_TITLE" size="20" maxlength="50"  class=input ></td>
			          <td width=40><input type="button"  value="Master 등록" class=btn onclick="save_mst('등록');"></td>
                      <td width=40><input type="button"  value="Master 삭제" class=btn onclick="save_mst('삭제');"></td></form>
		          </tr>
		          </table>	
		          <!-- 마스터 코드 등록 부분 끝 -->
		  </TD>
		  <TD ALIGN="right" width="50%">
		          <!-- 서브 코드 등록 부분 시작 -->
		          <Table border=0 cellpadding=3 cellspacing=0>
		          <TR><form name=form1 action="sys_code.asp" method="post">
		          
		              <input TYPE="HIDDEN" name="GUBUN_TITLE" value="<%=Request("gubun_title")%>" /> 		              
		              <input TYPE="HIDDEN" name="GUBUN_SEQ" > 
			          <INPUT TYPE="HIDDEN" NAME="Nvcnt" VALUE="<%=Nvcnt%>">
			          <INPUT TYPE="HIDDEN" NAME="RNvCnt">
			          <INPUT TYPE="HIDDEN" NAME="x">
			          <INPUT TYPE="HIDDEN" NAME="y">	
			          	              
		              <TD><input type=radio name=SLT value="등록1" id=A1> 등록 
		                  <input type=radio name=SLT value="수정1" id=A2> 수정 
		                  <input type=radio name=SLT value="삭제1" id=A3> 삭제</Td>
		              <TD height=26 align=center style="padding-left:5px;">Code</td>
		              <TD><input type="text" name="GUBUN_CODE" size="15" maxlength="2"  class=input value="<%=Trim(Request("gubun_code"))%>"></td>
		              <TD height=26 align=center style="padding-left:5px;">Name</td>
		              <TD><input type="text" name="GUBUN_NAME" size="20" maxlength="50"  class=input></td>
			          <TD width=40><input type="button"  value="Sub 등록" class=btn onclick="save_sub();"></td></form>
		          </TR>
		          </TABLE>	
		          <!-- 서브 코드 등록 부분 끝 -->

	  </TD></TR>
    </TABLE>
    <!--  HEAD 부분 끝 -->


    <!-- CONTENT 부분 시작 -->
    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="980"  style="TABLE-LAYOUT: fixed;" >
    <tr><td valign=top width="50%">

        <DIV class="layer_01">
            <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="450"  style="TABLE-LAYOUT: fixed;" >
            <TR>
		        <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="20%">순번</TD>
		        <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="30%">코드</TD>
		        <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="50%">코드명</TD>
            </TR>
            </TABLE>
          
            <DIV CLASS="scr_01" id="scr_01" style="width:100%;height:300px;overflow-y:auto;">
                <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="450" BGCOLOR=666666 style="TABLE-LAYOUT: fixed;">
                <%
                nVcnt = 1
                WHILE NOT RS.EOF 
                
                IF RS("COMMCD")=""	THEN GUBUN_CODE=""	ELSE GUBUN_CODE  = RS("COMMCD")		END IF
                IF RS("COMMNM")="" THEN GUBUN_TITLE="" ELSE GUBUN_TITLE = RS("COMMNM")	END IF
                                    
                %>
                <TR Onclick="javascript:sys_code('<%=SLT%>','<%=Server.HTMLencode(GUBUN_CODE)%>','<%=Server.HTMLencode(GUBUN_TITLE)%>','<%=nVcnt%>');" style="cursor:hand;">
                    <TD id="<%=Nvcnt%>-1"  ALIGN="CENTER"  WIDTH="20%" CLASS="TBL_DRW_01"><%=(nVcnt)%></TD>	   	
                    <TD id="<%=Nvcnt%>-2"  ALIGN="CENTER"  WIDTH="30%" CLASS="TBL_DRW_03"><%=GUBUN_CODE%></TD>	   
                    <TD id="<%=Nvcnt%>-3"  ALIGN="LEFT"    WIDTH="50%" CLASS="TBL_DRW_01"><%=GUBUN_TITLE%></TD>	    	   
                </TR>       
                <%
                FONTCLT = ""
                nVcnt = nVcnt + 1
                Rs.MoveNext    
                
                Wend
                Rs.Close : Set Rs=nothing           
               %>
                </TABLE>
	        </DIV>
        </DIV> 
    
    </td>
    <td width="1%" ></td>
    <td valign=top width="49%">
    
         <DIV id="cont1"></DIV>
    
     </td></tr>
    </TABLE>
    <!-- CONTENT 부분 끝 -->
    
    <script language="javascript">
    <!--
    //    window.onload = function() {
    //        sys_code('', '<%=REQUEST("GUBUN_CODE") %>', '', 1);
    //    }
	// 스크롤 위치 기억하기 스크립트 시작 //
	function scroll_Focus(gubun,x,y){
	  var db = (document.getElementById("scr_01")) ? 1 : 0; 
	  var scroll = (document.getElementById("scr_01").scrollTo) ? 1 : 0; 
	  if (gubun=="x"){
		document.getElementById("scr_01").scrollLeft = x;
	  }else if(gubun=="y"){
		document.getElementById("scr_01").scrollTop = y;
	  }else{
		document.getElementById("scr_01").scrollLeft = x;
		document.getElementById("scr_01").scrollTop = y;
	  }
	}

	function Row_rollover(Nvcnt,cnt)
	{
		var oldnvcnt = document.form.old_nvcnt.value;
		if(oldnvcnt==""){
			for(i=1;i<cnt;i++){
				document.getElementById(Nvcnt+"-"+i).style.backgroundColor='activecaption';
				document.getElementById(Nvcnt+"-"+i).style.color='white';
			}
		}else{
			for(i=1;i<cnt;i++){
				document.getElementById(oldnvcnt+"-"+i).style.backgroundColor='';
				document.getElementById(oldnvcnt+"-"+i).style.color='';
				document.getElementById(Nvcnt+"-"+i).style.backgroundColor='activecaption';
				document.getElementById(Nvcnt+"-"+i).style.color='white';
			}			
		}
		document.form.old_nvcnt.value = Nvcnt;
	}
    
    function save_mst(slt)
    {
       if ( '' == document.form.GUBUN_CODE.value )
       {
          alert( 'Master 코드를 입력하십시오.' );
          document.form.GUBUN_CODE.focus();
          return;
       }
       
       if ( '' == document.form.GUBUN_TITLE.value )
       {
          alert( 'Master 코드명을 입력하십시오.' );
          document.form.GUBUN_TITLE.focus();
          return;
      }
       form.SLT.value = slt;
       document.form.submit();
    }
    
    function save_sub()
    {
       if ( '' == document.form1.GUBUN_CODE.value )
       {
          alert( 'Master 코드를 입력하십시오.' );
          document.form1.GUBUN_CODE.focus();
          return;
       }
       
       if ( '' == document.form1.GUBUN_NAME.value )
       {
          alert( 'Sub Name을 입력하십시오.' );
          document.form1.GUBUN_NAME.focus();
          return;
       }   
       
		var db = (document.getElementById("scr_01")) ? 1 : 0; 
		var scroll = (document.getElementById("scr_01").scrollTo) ? 1 : 0; 
		var x = (db) ? document.getElementById("scr_01").scrollLeft : pageXOffset;
		var y = (db) ? document.getElementById("scr_01").scrollTop : pageYOffset;
		document.form1.x.value = x;
		document.form1.y.value = y;
		document.form1.submit();
		
	}  
          
    function sys_code(SLT,GUBUN_CODE,GUBUN_TITLE,Nvcnt) {

	        var form = document.form1;
	        var mform = document.form;
			//document.getElementById.("A1").value = "등록";					
			form.GUBUN_CODE.value	= GUBUN_CODE;					
			form.GUBUN_TITLE.value  = GUBUN_TITLE;
			form.GUBUN_NAME.value = "";

			mform.GUBUN_CODE.value = GUBUN_CODE;
			mform.GUBUN_TITLE.value = GUBUN_TITLE;
	          	
	
			form.Nvcnt.value		= Nvcnt;
			form.RNvCnt.value		= Nvcnt;

			Row_rollover(Nvcnt,4);
			
			document.all("A1").checked = true;
			document.all("A2").disabled = true;
			document.all("A2").disabled = true;
			document.all("A3").disabled = true;			
					
			//alert(querystring);
			var Result = RSExecute("/admin/sys_code1.asp","sys_code",GUBUN_CODE);
            var retval=Result.return_value; 
            
            document.getElementById("cont1").innerHTML = retval;
			
	}


    function select_sub(GUBUN_CODE, GUBUN_NAME, GUBUN_SEQ)
    {
        document.form1.GUBUN_CODE.value = GUBUN_CODE;
        document.form1.GUBUN_NAME.value = GUBUN_NAME;
        document.form1.GUBUN_SEQ.value = GUBUN_SEQ;

        selectrow();
        
        document.all("A2").checked = true;
        document.all("A1").disabled = true;
        document.all("A2").disabled = false;
        document.all("A3").disabled = false;
    }


    var selectedRowIndex = "";

    /*********************************************************************************
    * 마우스 롤오버 관련 스크립트 시작
    *********************************************************************************/

    function selectrow() {
        var oRow = window.event.srcElement.parentNode;
        var i, j

        // 일단 모든 Row 정상화처리
        if (selectedRowIndex !== "") {
            for (i = 0; i < datalist.rows(selectedRowIndex).cells.length; i++) {
                datalist.rows(selectedRowIndex).cells(i).style.backgroundColor = '';
                datalist.rows(selectedRowIndex).cells(i).style.color = '';
            }
        }

        // 해당 Row 만 반전처리
        for (j = 0; j < datalist.rows(oRow.rowIndex).cells.length; j++) {
            datalist.rows(oRow.rowIndex).cells(j).style.backgroundColor = 'activecaption';
            datalist.rows(oRow.rowIndex).cells(j).style.color = 'white';
        }
        selectedRowIndex = oRow.rowIndex;
    }  
    
    //-->
    </script>

    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->

