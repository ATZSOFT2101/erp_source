    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 프로그램 구성 메뉴 관리                              			--> 
    <!-- 내용: Mlevel 구성 관리                                                     --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 13일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    
       SLT		=  TRIM(REQUEST("SLT"))
       IF SLT = "" THEN S1 = "CHECKED" END IF

       SELECT CASE SLT
       CASE "ADD"  S1 = "CHECKED"
       CASE "EDIT"  S2 = "CHECKED"
       END SELECT

       IF   SLT <> "" THEN

            SELECT CASE  SLT
            
            CASE  "조회"
		          SQL = "SELECT * FROM MLEVEL WHERE ID = " & REQUEST("ID")
		          SET RS = SYSCON.EXECUTE(SQL)

		          IF NOT ( RS.EOF OR RS.BOF ) THEN
			         ID		=  RS("ID") 
			         PGID	=  RS("PGID") 
			         NAME	=  RS("NAME") 
			         GUBUN1	=  RS("GUBUN1") 
		          END IF
    		      
            CASE  "ADD"
            
		          SQL = "SELECT * FROM MLEVEL WHERE ID = " & REQUEST("pID")
		          SET RS = SYSCON.EXECUTE(SQL)

                  'SQL = "UPDATE MLEVEL SET STEP = STEP + 1 WHERE REF= " & RS("REF") & " AND STEP > " & RS("STEP")
                  'response.Write sql
                  'SYSCON.EXECUTE(SQL)
                  
                  SQL = "UPDATE MLEVEL SET GUBUN = 'parent' WHERE ID = " & REQUEST("pID")
                  SYSCON.EXECUTE(SQL)   
                  
		          'IF RS("LASTCNT") <> 0 THEN
                  'LASTCNT = RS("LASTCNT") + 1
                  'ELSE 
                  'LASTCNT = 0
                  'END IF     
                  
 		          IF RS("DIS_SEQ") <> NULL THEN
                     DIS_SEQ = RS("DIS_SEQ") + 1
                  ELSE 
                     DIS_SEQ = 1
                  END IF      
                        		  
                  SQL = "INSERT INTO MLEVEL(PGID,REF,STEP,RE_LEVEL,NAME,cname,ename,rname,GUBUN,PID,ISTDATE,UDTDATE,GUBUN1,LAST,DIS_SEQ) VALUES "
                  SQL = SQL & "("& REQUEST("PGID") &","& RS("REF") &", "& RS("STEP") + 1 &","& RS("RE_LEVEL") + 1 &", N'"& REQUEST("NAME") &"', N'"& REQUEST("cNAME") &"', N'"& REQUEST("eNAME") &"', N'"& REQUEST("rNAME") &"','child',"& REQUEST("pID") &",'"& DATE &"','"& DATE &"','"& REQUEST("GUBUN1") & "','N',"& DIS_SEQ &")"
                  SYSCON.EXECUTE(SQL)  
               
                  SQL3 = "SELECT ID FROM MLEVEL WHERE PID = " & REQUEST("pID") &" ORDER BY NAME"
                  SET RS2 = SYSCON.EXECUTE(SQL3)

                  CSTEP = RS("STEP") + 1
                  WHILE NOT RS2.EOF 
		              SYSCON.EXECUTE("UPDATE MLEVEL SET STEP = " & CSTEP & " , LAST = 'N' WHERE ID = " & RS2("ID")) 
		              CSTEP = CSTEP + 1
		              PID = RS2("ID")
		              RS2.MOVENEXT
		          WEND   
		          SYSCON.EXECUTE("UPDATE MLEVEL SET LAST = 'Y' WHERE ID = " & PID) 

			      response.Redirect "sys_menu_cont.asp?pid="& Request("pid")

            CASE  "EDIT"
            
                  if Request("chk_yn") = "on" then
                     chk_yn = "Y"
                  else
                     chk_yn = "N"
                  end if
            
		          SQL = "SELECT * FROM MLEVEL WHERE ID = " & REQUEST("ID")
		          SET RS = SYSCON.EXECUTE(SQL)
                                    
                  IF Not Rs.EOF THEN
                  Sql = "update mlevel set pgid="& Request("pgid") &","
                  sql = sql &" name =N'"& urlDecode(Request("name")) &"',"
                  sql = sql &" cname =N'"& urlDecode(Request("cname")) &"',"
                  sql = sql &" ename =N'"& urlDecode(Request("ename")) &"',"
                  sql = sql &" rname =N'"& urlDecode(Request("rname")) &"',"
                  sql = sql &" udtdate='"& date &"',gubun1 ='"& Request("gubun1")&"',"
                  sql = sql &" chk = '" & chk_yn & "' where id  =  " & Request("id")
                  syscon.Execute(Sql)  
                  END IF

			      response.Redirect  "sys_menu_cont.asp?pid="& Request("pid")

            CASE  "DEL"

                  SET RS = SYSCON.EXECUTE("SELECT * FROM MLEVEL WHERE ID  in  ("& REQUEST("CHKBOX") &") ")
                  
                  Do until Rs.EOF
                  
                      IF TRIM(RS("GUBUN")) <> "parent" THEN
                
			             SYSCON.EXECUTE("DELETE FROM MLEVEL WHERE ID = " & Rs("id"))   
			             SYSCON.EXECUTE("DELETE FROM menu_duty WHERE mr_mlvl_id = " & Rs("id")) 
			             
			             SET RS1 = SYSCON.EXECUTE("SELECT COUNT(PID) AS CNT FROM MLEVEL WHERE PID = " & RS("PID"))
			             
			             IF RS1("CNT") = 0 THEN
			                SQL = "UPDATE MLEVEL SET GUBUN = 'child' WHERE ID = " & RS("PID")
			                SYSCON.EXECUTE(SQL)      
			             END IF  
			             
		              END IF	
    	          
	              RS.MoveNext
	              loop
			    response.Redirect "sys_menu_cont.asp?pid="& Request("pid")
        	
        	
            END SELECT 
                   
       END IF    
       
       IF REQUEST("pID") = "" THEN
       pID = 324
       ELSE
       pID = REQUEST("pID")
       END IF
       
       SQL = "SELECT ID, PGID, DIS_SEQ, GUBUN1,PRG_ISTDATE, PRG_UDTDATE, A.NAME, isnull(A.cNAME,'') cname,isnull(A.eNAME,'') eNAME,isnull(A.rNAME,'') rNAME, B.PRG_ADDRESS, B.PRG_ADDRESS1,chk FROM MLEVEL "
       SQL = SQL &" A JOIN PRGID B ON A.PGID = B.PRG_ID WHERE A.PID =" & pID &"  ORDER BY DIS_SEQ"
       SET RS = SYSCON.EXECUTE(SQL)
       
    %>
    <BODY CLASS="BODY_01" SCROLL="no" style="background-color:#C3DAF9;margin:0;border:0;" onload="resize_scroll(370);" onresize="resize_scroll(370);">

    
    <!-- HEAD 부분 시작 --> 
    <TABLE CLASS=Sch_Header id=d2>
	  <TR><TD><form name="check" method="post" action="sys_menu_cont.asp?slt=DEL&pid=<%=pid%>">
		      <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="99%">
		      <TR><TD></TD>
				  <TD ALIGN="right">

		          <input TYPE=button CLASS=btn name="Submit" value=" 삭 제 " onclick="ck()" onMouseOver="this.style.cursor = 'hand'">
		          <INPUT TYPE=button CLASS=btn value="새로고침" Onclick="window.location.reload();">
		          <INPUT TYPE=button CLASS=btn value="순서정렬" Onclick="centerWindow('sys_menu_seq.asp?id=<%=pid%>','addwin','350','320','no');">
		          <INPUT TYPE=button CLASS=btn value="미리보기" Onclick="centerWindow('../DA_menu1.asp?pid=<%=pid%>','prewin','800','500','no');">

                    
			  </TD></TR>
			  </TABLE>
	  </TD></TR>
    </TABLE>
    <!--  HEAD 부분 끝 --> 
       
    <!-- CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="95%"  style="TABLE-LAYOUT: fixed;" >
        <tr>
            <TD ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="5%"><INPUT TYPE="checkbox" NAME="allbox" size=1 onclick="CheckAll();"></td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="5%">SEQ</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="14%">한글</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="14%">중문</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="15%">영문</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="15%">러시아</td>
            
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="05%">종류</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="08%">바뀐날짜</td>
            <TD ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="20%">목록주소</td>	
        </TR>
        </TABLE>
        <DIV CLASS="scr_01" id=scr_01>
           <TABLE BORDER="0" id="datalist" ELLPADDING="0" CELLSPACING="0" WIDTH="95%" style="TABLE-LAYOUT: fixed;">
	       <%   
	       WHILE NOT RS.EOF 
           %>	
           <TR ONCLICK="show_data('<%=Rs("PGID")%>','<%=Server.HtmlEncode(Trim(Rs("NAME")))%>','<%=Server.HtmlEncode(Trim(Rs("cNAME")))%>','<%=Server.HtmlEncode(Trim(Rs("eNAME")))%>','<%=Server.HtmlEncode(Trim(Rs("rNAME")))%>','<%=Trim(Rs("gubun1"))%>','<%=Rs("id")%>','<%=Rs("chk")%>');"  style="cursor:hand;" ><td ALIGN="CENTER" WIDTH="05%" CLASS="TBL_DRW_00" style="border-left:none;"><input type="checkbox" name="chkbox" value="<%=Rs("id")%>"></td>       
	           <td ALIGN="CENTER" WIDTH="05%" CLASS="TBL_DRW_01"><%=Rs("dis_seq")%>&nbsp;</td>
	           <td ALIGN="LEFT"   WIDTH="14%" CLASS="TBL_DRW_03" ><%=Server.Htmlencode(Rs("name"))%></a></td>   
	           <td ALIGN="LEFT"   WIDTH="14%" CLASS="TBL_DRW_01"><%=Rs("cname")%>&nbsp;</td>
	           <td ALIGN="LEFT"   WIDTH="15%" CLASS="TBL_DRW_01"><%=Rs("ename")%>&nbsp;</td>
	           <td ALIGN="LEFT"   WIDTH="15%" CLASS="TBL_DRW_01"><%=Rs("rname")%>&nbsp;</td>
	           <td ALIGN="CENTER" WIDTH="05%" CLASS="TBL_DRW_01"><%=Rs("gubun1")%>&nbsp;</td>
	           <td ALIGN="CENTER" WIDTH="08%" CLASS="TBL_DRW_01"><%=left(Rs("PRG_udtdate"),10)%>&nbsp;</td>
	           <td ALIGN="LEFT"   WIDTH="20%" CLASS="TBL_DRW_01"><%=Rs("PRG_address")%>&nbsp;</td>
	           <!--<td ALIGN="LEFT"   WIDTH="15%" CLASS="TBL_DRW_01"><%=Rs("PRG_address1")%>&nbsp;</td>-->
	        </tr>
	        <%
            RS.MOVENEXT               
            WEND
	        %></form>  
            </TABLE>
        </DIV>
    </DIV> 

    <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%" HEIGHT="100%" bgcolor=#D1E3FA>
    <tr><form name="form" action="sys_menu_cont.asp" method="POST">
		<Input type="hidden" name="pid" value="<%=request("pid")%>" >
		<Input type="hidden" name="id" value="" >
		<input type="hidden" name="PRG_toggle" value="1" />
		<Td height=30 style="padding-left:5px;">
		    <input type="radio" <%=(S1)%> name="slt" value="ADD" id=slt1 onclick="clear_data()">등록                
			<input type="radio" <%=(S2)%> name="slt" value="EDIT" id=slt2>수정</Td>
        <td height=30 style="padding-right:5px;text-align:right;" onclick="resize_panel();" style="cursor:hand;"><img src="<%=img_path%>colbtn.gif" border=0 /></td></tr>
    <tr><td valign=top colspan=2>
    
            <center>
            <FIELDSET style="width:99%;">
            <Div class="div_separate_5"></div>
            
            <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" align=left>
            <tr>
		        <%
		        IF REQUEST("MODE") ="메뉴" THEN
		        %>
                <td><input type="text" size="35"  name="Pgid" value="9999" class=input readonly></td> 
		        <%
		        ELSEIF REQUEST("PGID")=9999 THEN
		        %>
                <td><input type="text" size="35"  name="Pgid" value="<%=Pgid%>" class=input ></td> 
		        <%
		        ELSE
		        %>
		        <td><% CALL search_frm("sys01","","Pgid","name","searchvalue",""& Pgid &"",""& name &"") %></td> 
		        <%
		        END IF
		        %>
        	
		        <%
		        IF REQUEST("MODE") = "메뉴"  THEN
		        %>
                <td>프로그램명</td>
                <td><input type="text" size="35"  name="name" value="<%=name%>" class=input ></td> 
                
		        <%
		        ELSEIF REQUEST("PGID")=9999 THEN
		        %>
                <td>프로그램명</td>
                <td><input type="text" size="35"  name="name" value="<%=name%>" class=input ></td> 
		        <%
		        END IF
		        %>
                <td><% 
		            IF trim(gubun1) <> "" then
			           IF trim(gubun1)="폴더" then  
				          Folder_slt = "checked"
			           ELSE
				          File_slt = "checked"
			           End IF
			        ELSE
			           IF trim(request("mode"))="메뉴" then  
				          Folder_slt = "checked"
			           ELSE
				          File_slt = "checked"
			           End IF
			        END IF
			        
			        %>
			    중문 <input type="input" name="cname" class="input" style="font-family:Arial;font-size:12px;width:130;" />
			    영문 <input type="input" name="ename" class="input" style="font-family:Arial;font-size:12px;width:130;" />
			    러시아 <input type="input" name="rname" class="input" style="font-family:Arial;font-size:12px;width:130;" />
			    <input type=radio name="gubun1" value="폴더" <%=folder_slt%> id=gubun1_1  onclick="folder_data();"> 폴더 
			    <input type=radio name="gubun1" value="파일" <%=file_slt%> id=gubun1_2> 파일
			    <input type="checkbox" id="chk_yn" name="chk_yn" /> 사용유무
			    </td>  
			    <td width=100 align=left><input type="submit"  value=" 확 인 " class=btn> </td></form>
	        </tr>
	        </TABLE> 
        	
            </FIELDSET>
            
            <Div class="div_separate_10"></div>
            
            <FIELDSET style="width:99%;background-color:White;">
           
            <Div class="div_separate_5"></div>
            
             <table border=0 cellpaddinjg=0 cellspacing=0 width="98%">
             <tr><td class=blue_normal>프로그램 메뉴 구성시 주의사항</td></tr>
             <tr><td height=5></td></tr>
             <tr><td>프로그램 공용사용을 위해 반드시 동아화성 PRG ID 값으로 등록하셔야 합니다.</td></tr>
             <tr><td>기존의 삼성화학/유창/명성으로 등록된 PRG ID값들은 조회가 되지 않습니다.</td></tr>
             <tr><td height=5></td></tr>
             
             <tr><td>구분값을 <span class=blue_normal>폴더</span>로 선택하시면 자동으로 <b>PRGID</b> 값에 <b>9999</b>와 <b>PRG NAME</b>값에 <b>폴더</b>라고 입력이 됩니다.</td></tr>
             <tr><td>폴더명 등록시는 PRGID값은 9999로 두시고 프로그램명만 수정하시면 됩니다.</td></tr>
             
             
             </table>
            
            
            </FIELDSET>
          
            </center>
            
        </td></tr>
    </TABLE>                
    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->

    <script language="javascript">
    
     var selectedRowIndex    = "";

 //   document.getElementById("s_gubun").style.width = "120";
 //   document.getElementById("name").style.disabled=false;
	/*********************************************************************************
	* 마우스 롤오버 관련 스크립트 시작
    *********************************************************************************/
    
	function selectrow()
	{
	    var oRow = window.event.srcElement.parentNode;
        var i, j

        // 일단 모든 Row 정상화처리
		if(selectedRowIndex!==""){
            for (i=0;i<datalist.rows(selectedRowIndex).cells.length;i++)
            {
                datalist.rows(selectedRowIndex).cells(i).style.backgroundColor = '';
                datalist.rows(selectedRowIndex).cells(i).style.color = '';
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

 
 
    function show_data(Pgid,name,cname,ename,rname,gubun1,id,chk) {
    
        var form = document.form;
        
        form.Pgid.value = Pgid;
        form.name.value = name;
        form.cname.value = cname;
        form.ename.value = ename;
        form.rname.value = rname;

        if(gubun1=='폴더'){
        document.getElementById("gubun1_1").checked = true;
        document.getElementById("gubun1_2").checked = false;
        }else if(gubun1=='파일'){
        document.getElementById("gubun1_1").checked = false;
        document.getElementById("gubun1_2").checked = true;
        }
        if (chk == "Y")
            document.getElementById("chk_yn").checked = true
        else
            document.getElementById("chk_yn").checked = false;
            
        form.id.value = id;
        document.getElementById("slt2").checked = true;
        selectrow()
    }
    
    function folder_data() {
    
        var form = document.form;

        form.Pgid.value = 9999;
        form.name.value = "";
    
    }
    
    function clear_data() {
    
        var form = document.form;
        
        form.Pgid.value ="" ;
        form.name.value ="" ;
        
        document.getElementById("gubun1_1").checked = false;
        document.getElementById("gubun1_2").checked = true;

        form.id.value ="" ;
        document.getElementById("slt1").checked = true;
   
    }   
       
    function resize_panel() {
        var scroll_cnt;
        if(document.form.PRG_toggle.value==1){
        resize_scroll(90);
        document.form.PRG_toggle.value = 0;
        }else{
        resize_scroll(370);
        document.form.PRG_toggle.value = 1;
        }
    }
    
    function ck() {

    if (confirm("선택하신 데이타를 삭제하시겠습니까? ")) {
    //		location = 'order_list.asp?mode=del_all&countn='+nu;
        document.check.submit();
        return;
	    }
    else {
	    return;
       }   
    }

    </script>
