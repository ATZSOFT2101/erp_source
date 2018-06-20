    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램 : 테이블/필드 관리                                 			        --> 
    <!-- 작 성 자 : 문성원(050407)                                                    --> 
    <!-- 작성일자 : 2006년 6월 08일                                                   --> 
    <!-- 내    용 : 테이블/필드 관리                                                   --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
   	css1 = " style=""font-family:arial;font-size:12px;"" "

    gubun = request("gubun")

	IF request("tabledb") = ""  then
	   tabledb="syerp"
	else
	   tabledb = request("tabledb")
	end if

	SELECT CASE gubun
	CASE "rebuild_table"

		tbSQL ="select name from "& tabledb &".dbo.sysobjects where xtype='u'"
		Set tbRS=syscon.execute(tbSQL)

		IF Not tbRS.EOF THEN
		   
		   ti_Vcnt = 0
		   td_Vcnt = 0
		   While Not tbRs.EOF
				 
				 name = Trim(tbRs("name"))
				 tbSQL1 = "select tableid from SYS_Table where tableid ='"& name &"' "
				 set tbRs1=syscon.execute(tbsql1)

				 IF tbRs1.EOF THEN
					syscon.execute("insert into sys_table (TableDB, TableID, TableNM, TableDate) values ('"& tabledb &"','"& name &"','...',getdate())")
					ti_Vcnt = ti_Vcnt+1
                 'else
					'syscon.execute("delete sys_table where TableDB = '"& tabledb &"' and TableNM ='"& name &"'")
					'syscon.execute("delete sys_field where fldDB = '"& tabledb &"' and fldtable ='"& name &"'")
					'td_Vcnt = td_Vcnt+1	     
			     ENd if

		   tbRs.MoveNext
		   Wend

		   tbRs.Close
		   Set tbRs = Nothing

		   Alert_message ""& ti_Vcnt&" 건 생성 / "& td_Vcnt&" 삭제처리 ","sys_table_left.asp?tabledb="& tabledb &""

		END IF
    
    case "delete"
    
          SET RS = SYSCON.EXECUTE("SELECT * FROM sys_table WHERE table_idx  in  ("& REQUEST("CHKBOX") &") ")
          
          Do until Rs.EOF
          
        		tbSQL ="select name from "& tabledb &".dbo.sysobjects where xtype='u' and name='"& trim(rs("TableNM")) &"'"
		        Set tbRS=syscon.execute(tbSQL)
                
                if  tbrs.eof then
	                SYSCON.EXECUTE("DELETE FROM sys_table WHERE table_idx = " & Rs("table_idx")) 
	            end if  
                tbrs.close

          RS.MoveNext
          loop
          

	END SELECT
	%>
    <BODY CLASS="BODY_01" SCROLL="no" style="background-color:#C3DAF9;margin:0;border:0;"  onload="resize_scroll(65);get_data_list('<%=tabledb%>');" onresize="resize_scroll(65);">

    <div id="loading" style="position:absolute;width:100%;height:100%;top:200;left:0;z-index:10000;">
        <table border="0" cellpadding="0" cellspacing="1" width="400" style="background-color:999999;" align="center">
            <tr><td height="60" style="background-color:#ffffff;text-align:center;color:#ff9900;"><img src="<%=img_path%>load3.gif" border=0>
            <div style="padding-top:5;"></div>데이터를 불러오는 중입니다.</td></tr>
        </table>
    </div> 


    <!-- HEAD 부분 시작 -->
    <TABLE CLASS=Sch_Header id=d2>
        <TR><TD>
      
        <table cellpadding=1 cellspacing=0 border=0 width="99%">
            <tr><form name="form" method="post">
                <td><select name="table_gubun" onchange="get_data_list();" <%=css1%>><%
                sql="select gubun_name,gubun_seq,rtrim(gubun_code)+convert(char(2),gubun_seq) as gcode from gubun1 where gubun_Code='WA' order by gubun_seq	"
                set ars=conn.execute(sql)

                while not ars.eof

                response.write "<option value='"& ars("gcode") &"'>"& trim(ars("gubun_name")) &"</option>"

                ars.movenext
                wend

                ars.close
                %></select></td>
                <td><select name=tabledb onchange="get_data_list();" <%=css1%>>
                <%
                sql="select name from master.dbo.sysdatabases	"
                set masterrs=syscon.execute(sql)

                if  not (masterrs.eof or masterrs.bof) then
                    while not masterrs.eof
                    if tabledb = trim(masterrs("name")) then
	                %>
	                <option value="<%=tabledb%>" selected <%=css1 %>><%=tabledb%></option>
	                <%
	                else
	                %>
	                <option value="<%=masterrs("name")%>" <%=css1 %>><%=masterrs("name")%></option>
	                <%
	                end if
                    masterrs.movenext
                    wend
                
                end if
                %></select></td>
              <td align=right>
              <input type="button" value="새로고침" onclick="get_data_list();" class=btn>
              <input type="button" value=" 등 록 " class=btn Onclick="centerWindow('sys_table_wrt.asp?tabledb=<%=tabledb%>','sys_table_wrt_win','350','240','no');">
              <input type="button" value=" 삭 제 " class=btn Onclick="set_delete();">
			  <input type="button" value=" 가져오기 " class=btn Onclick="URL_move('sys_table_left.asp?tabledb=' + document.form.tabledb.value + ' &gubun=rebuild_table');"></TD></form>
           </tr>
           </table>
      </TD></TR>
    </TABLE>
    <!-- HEAD 부분 끝 -->

    <!-- SHEET CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="550"  style="TABLE-LAYOUT: fixed;" >
        <TR><form name="check" method="post">
            <td ALIGN="CENTER" class="tbl_rw_00" width="07%">CHK</td>
            <td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="07%">SEQ</td>
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">업무</td>	
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="23%">TABLE명</td>
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="18%">생성일자</td>	
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="35%">TABLE설명</td>	
        </TR>      
        </TABLE>
        <DIV CLASS="scr_01" id="scr_01">

        </DIV>
               
    </DIV> 

    <script type="text/javascript">
 
    var form = document.forms;
    var selectedRowIndex = "";
    
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

	/*********************************************************************************
	* 데이터 클릭시 Load 프레임 스크립트 시작
    *********************************************************************************/
    
    
    function loaddata(tablenm,tableid,tabledb)
    {
        selectrow();
        
        var url_1th = 'sys_table_cont.asp?tablenm='+tablenm+'&tableid='+tableid+'&tabledb='+tabledb;
        var url_2th = 'sys_table_preview.asp?tablenm='+tablenm+'&tableid='+tableid+'&tabledb='+tabledb
        
        loadFrames('sub_c',url_1th, 'sub_c1',url_2th);
    }
    
    
    /*********************************************************************************
	* Ajax 처리 스크립트 시작
	* 거래선 선택시 레코드 조회
    *********************************************************************************/
    function get_gugun_list_callback()
    {
          if(ajax_gugun.readyState == 4){   
               if(ajax_gugun.status == 200){
                    if(typeof(document.all.scr_01) == "object"){
                        document.all.scr_01.innerHTML = ajax_gugun.responseText;
                    }
               }
          }
     }
     
     var ajax_addr;    

     function get_data_list()
     {
     
        selectedRowIndex = "";
	    document.getElementById("loading").style.display    = "";
	            
        var tgubun      = document.form.table_gubun.value;
        var tabledb     = document.form.tabledb.value;

        ajax_addr = new ActiveXObject("Microsoft.XMLHTTP");
        ajax_addr.onreadystatechange = get_data_list_callback;
        ajax_addr.open("GET", "sys_table_leftx.asp?tabledb="+tabledb+"&table_gubun="+tgubun, true);      
        ajax_addr.send();
        
     }

     function get_data_list_callback()
     {
          if(ajax_addr.readyState == 4){
              if(ajax_addr.status == 200){
                  if(typeof(document.all.scr_01) == "object"){
                        if(ajax_addr.responseText.length > 0){
                            document.getElementById("scr_01").innerHTML = ajax_addr.responseText;
                        }
                       else
                             document.getElementById("scr_01").innerHTML = "오류발생";     
                   }
                   document.getElementById("loading").style.display = "none";
              }
         }
     } 	    
    
    /*********************************************************************************
	* 체크박스 선택 후 삭제시 스크립트
    *********************************************************************************/
    
    function set_delete()
    {
        if (confirm("선택하신 테이블을 삭제하시겠습니까? ")) {
            document.check.action = 'sys_table_left.asp?tabledb=<%=tabledb%>&gubun=delete';
            document.check.submit();
            return;
	    }else{
	        return;
        }    
    }
    
    /*********************************************************************************
	* 수정팝업창 스크립트
    *********************************************************************************/
    
    function set_win(tableid,tabledb)
    {
        centerModal('sys_table_Wrt.asp?tableid='+tableid+'&tabledb='+tabledb+'&slt=조회','350','250');
    }     
    
	document.getElementById("loading").style.display = "none";
    
    
    </script>


    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
