	<!--#include Virtual = "/INC/INC_TOP.ASP"  -->

    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
	<body class="body_01" scroll="no" style="background-color:#D1E3FA;margin:0;border:0;" onload="resize_scroll(81);" onresize="resize_scroll(81);">
        
    <div id="loading" style="position:absolute;width:100%;height:100%;top:200;left:0;z-index:10000;">
        <table border="0" cellpadding="0" cellspacing="1" width="400" style="background-color:999999;" align="center">
            <tr><td height="60" style="background-color:#ffffff;text-align:center;color:#ff9900;"><img src="<%=img_path%>load3.gif" border=0>
            <div style="padding-top:5;"></div>데이터를 불러오는 중입니다.</td></tr>
        </table>
    </div> 
    
    <!-- HEAD 부분 시작 -->
    <table class="sch_header" id="Table1">
	<tr><form name="form" method="post">
		<td><table border="0" cellpadding="0" cellspacing="0">
			<tr>
			    <td align="center"  width="40" >공장</td>
				<td><% call fac_slt1("전체","m_factcd","SS",""& m_factcd &"") %></td>
			    <td align="center"  width="40" >구분</td>
				<td><% CALL FAC_slt1("전체","m_gubun","WB",""& m_gubun &"") %></td>	
				<td><input type="button" class="btn" value= " 검 색 " onclick="get_data_list();" id="search_btn">
				<td><input type="button" class="btn" value= " 등 록 " onclick="popwin();" id="Button1"></td>
			</tr>
			</table>
		</td></form>
	</tr></table>
	<!-- HEAD 부분 끝 -->

	<!-- CONTENT 부분 시작 -->
    <div class="layer_01" id="layer_01" >
	    <table border="0" cellpadding="0" cellspacing="0" width="1100" style="table-layout: fixed;">
		    <tr>
			    <td align="center" width="05%" class="tbl_rw_01">No.</td>
			    <td align="center" width="06%" class="tbl_rw_01">구분</td>
			    <td align="center" width="09%" class="tbl_rw_01">사용자명</td>
			    <td align="center" width="13%" class="tbl_rw_01">코드</td>
			    <td align="center" width="05%" class="tbl_rw_01">노트북</td>
			    <td align="center" width="15%" class="tbl_rw_01">컴퓨터명</td>
			    <td align="center" width="12%" class="tbl_rw_01">Mac Adress</td>
			    <td align="center" width="10%" class="tbl_rw_01">IP</td>
			    <td align="center" width="10%" class="tbl_rw_01">OS</td>
			    <td align="center" width="05%" class="tbl_rw_01">Virus</td>
			    <td align="center" width="05%" class="tbl_rw_01">Bizmate</td>
			    <td align="center" width="05%" class="tbl_rw_01">Office</td>
		    </tr>
	    </table>
	    <div class="scr_01"  id="scr_01">

	    </div>
	</div>
	<!-- CONTENT 부분 끝 -->
	
    <!-- PAGING OR 합계,집계 부분 시작 -->    
	<div class="layer_03" style="border-top:none;">
	<table border="0" cellpadding="0" cellspacing="0" width="980" style="table-layout:fixed;">
		<tr>
			<td class="tbl_rw_01" style="text-align:center;">&nbsp;</td>
		</tr>
	</table>
	</div>
    <!-- PAGING OR 합계,집계 부분 끝 -->   

    <script type="text/javascript">   

	var form		= document.form;
	/*********************************************************************************
	* 마우스 롤오버 관련 스크립트 시작
    *********************************************************************************/
    
    var selectedRowIndex = "";
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
	* 검색버튼 클릭시 스크립트
    *********************************************************************************/
    
    function set_search()
	{
	    document.getElementById("loading").style.display = "";
	    form.submit();	
	}
	
    /*********************************************************************************
	* Ajax 처리 스크립트 시작 ( 레코드 조회 )
    *********************************************************************************/
    
    var ajax_gugun;
    var ajax_addr;    
    
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

     function get_data_list()
     {
        selectedRowIndex = "";
        var m_factcd = form.m_factcd.value;
        var m_gubun = form.m_gubun.value;
        
	    var url = "it_machinex.asp?m_factcd="+m_factcd+"&m_gubun="+m_gubun;
        ajax_addr = new ActiveXObject("Microsoft.XMLHTTP");
        ajax_addr.onreadystatechange = get_data_list_callback;
        ajax_addr.open("GET",url, true);      
        ajax_addr.send();
     }

     function get_data_list_callback()
     {
          if(ajax_addr.readyState == 4){
              if(ajax_addr.status == 200){
                  if(typeof(document.all.scr_01) == "object"){
                  
                        document.getElementById("scr_01").innerHTML = "";
                        if(ajax_addr.responseText.length > 0){
                             document.getElementById("scr_01").innerHTML = ajax_addr.responseText;
                        }else{
                             document.getElementById("scr_01").innerHTML = "오류발생";     
                        }                             
                    }
                    document.getElementById("loading").style.display = "none";
                    document.getElementById("search_btn").disabled = false;
              }
         }
     } 		
     
    /*********************************************************************************
	* 상세 데이터 조회
    *********************************************************************************/	
	function show_data(m_idx) 
	{   
	    selectrow();
		centerModal('it_machinew.asp?m_idx='+m_idx,'700','550');
			
	}   
	 
    /*********************************************************************************
	* 메시지 Clear 로직 스크립트
    *********************************************************************************/	
    function clear_msg() {
	    document.getElementById("msg").innerHTML = "";
	}  	


	document.getElementById("loading").style.display = "none";

	</SCRIPT>

    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
