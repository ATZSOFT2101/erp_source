    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <body class="POPUP_01" scroll="no">
<%  call Func_Loading("로딩 중입니다...",img_path & "load3.gif") %>

<Fieldset align=center style="width:95%;"><LEGEND style="font-weight:bold; font-size:9pt;" > 부서구분 설정 </LEGEND><div style="padding-top:10px;"></div>

     <form name="search_form" method="post"  action="sys_dept.asp"></form> 

    <div class="layer_02" style="height:100%;background-color:#D1E3FA;padding-top:5px;border-top:0px solid #666666;border-left:none;">
<table  border="0" cellpadding="1" cellspacing="1">
   <form name="form" method="post" action="sys_dept.asp">
   <tr><tD WIDTH=20></TD>
      <td align="right">분류코드&nbsp;</td>
      <td WIDTH=70><input type="text" name="deptcd" size="5" maxlength="5" class="input_box" /></td>
      <td align="right">부서분류&nbsp;</td>
      <td><input type="text" name="deptnm" size="20" maxlength="50" class="input_box" /></td>
      <td><input type="button" value=" 저장 " onclick="set_submit('udt');" class="btn">&nbsp;<input type="button" value=" 삭제 " onclick="set_submit('del');" class="btn"></td>        
   </tr>
   <tr><td height="30" style="padding-left:15px;color:blue;padding-left:10px;" <%=css1%> id="msg" colspan="5"></td></tr>
   </form>
</table>
    </div>

    <!-- CONTENT 부분 시작 -->
    <div class="layer_01" >
<table width="100%"  border="0" cellpadding="4" cellspacing="1" align=center BGCOLOR=666666>
  <tr align=center> 
    <td width="20%" bgcolor="#629AD9" class="FFFFFF_BOLD">분류코드</td>
    <td width="30%" bgcolor="#629AD9" class="FFFFFF_BOLD">부서분류명</td>
    <td width="20%" bgcolor="#629AD9" class="FFFFFF_BOLD">분류코드</td>
    <td width="30%" bgcolor="#629AD9" class="FFFFFF_BOLD">부서분류명</td>
  </tr>
        </table>   
        <div class="scr_01" id="scr_01"></div>
    </div> 
    <!-- CONTENT 부분 끝 -->


    <script type="text/javascript" src="/INC/js/ajax_script.js"></script>
<script language="javascript">
<!--
    get_data_list('sys_deptx.asp')
    function selectgod(deptcd, deptnm) {
        document.form.deptcd.value = deptcd;
        document.form.deptnm.value = deptnm;
        document.form.oldcd.value = deptcd;
        document.form.oldnm.value = deptnm;
    }

    function set_submit(slt) {

        //  입력 필드 유효성 확인한다.
        if (err_msg(form, "deptcd", "분류코드를 입력하십시오.") == "err") return;
        if (err_msg(form, "deptnm", "부서명을 입력하십시오.") == "err") return;

        var strSQL = "exec sp_InUpDel_Proc "     //  등록,수정,삭제에 해당하는 Stored Procedure (sp_InUpDel_Proc)
	                   + " @slt = '" + slt + "', "  //  등록,수정,삭제 구분 내용(slt)
	                   + " @tblName = 'ADMIN_DEPT_GUBUN',"     //  해당 테이블명
	                   + " @valueStr = 'N''" + form.deptcd.value + "'',N''" + form.deptnm.value + "'',', "
		               + " @fdName = 'dept_cd, dept_nm,',"
		               + " @condStr = 'dept_cd=''" + form.deptcd.value + "''', "         //   수정시 조건을 넣어준다.(@condStr)
		               + " @addCond = 'dept_cd=''" + form.deptcd.value + "''', "       // 등록시 조건을 넣어준다(@addCond)
		               + " @maxField = '', "     // 최대값을 구하는 필드 (@maxField)
		               + " @maxCond = '', "    // 최대값을 구하기 위한 조건 내용(@maxCond)
		               + " @delCond = 'DELETE ADMIN_DEPT_GUBUN where dept_cd = ''" + form.deptcd.value + "'''";   // 삭제시 조건을 넣어준다(@delCond)

        // 쿼리를 실행하기 위하여 RSExecute 한다.
        var Result = RSExecute("sys_dept.asp", "set_process", strSQL);

        // 결과 값을 반환하여 메시지 출력한다.
        func_msgSetProcess(Result, 'sys_deptx.asp');
    }
//-->
</script>

    <!--#include Virtual = "/INC/inc_classTscript.ASP" -->
	<!--#include Virtual = "/INC/_ScriptLibrary/rs.asp"  -->	
    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
