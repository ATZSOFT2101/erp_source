    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!--'####  자료 조회 VBScript 입니다. ####' -->
<%  
      slt			    = Trim(Request("slt"))
      fac_ucd           = Request("fac_ucd")
      fac_fgubun        = Trim(Request("fac_fgubun"))
      
      IF  slt  <>  ""   THEN         
	      Sql = "select * from Admin_fac_gubun where fac_cd='"& fac_ucd &"' and fac_fgubun ='"& fac_fgubun &"'"
	      Set Rs = syscon.Execute(SQL)
                     
	      IF  Not Rs.Eof THEN
              fac_ucd           = RS("fac_cd")
              fac_unm            = Trim(RS("fac_nm"))
              fac_captin        = Trim(RS("fac_captin"))
              fac_bunho         = Trim(RS("fac_bunho"))
              fac_bupin         = Trim(RS("fac_bupin"))
              fac_post          = RS("fac_post")
              fac_address       = Trim(RS("fac_address"))
              fac_tel           = Trim(RS("fac_tel"))
              fac_fax           = Trim(RS("fac_fax"))
              fac_uptae         = Trim(RS("fac_uptae"))
              fac_jongmok       = Trim(RS("fac_jongmok"))
              fac_codepage      = Trim(RS("fac_codepage"))
              fac_lang          = Trim(RS("fac_lang"))
              fac_gw_domain     = Trim(RS("fac_gw_domain"))
              fac_mail_domain   = Trim(RS("fac_mail_domain"))
              fac_wr_date       = RS("fac_wr_date")
              fac_ed_date       = RS("fac_ed_date")
              fac_sabun         = RS("fac_sabun")
              fac_fgubun        = RS("fac_fgubun")
              fac_fgubun_nm     = Trim(RS("fac_fgubun_nm"))

	      END IF	  
      END IF	
%>    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <body class="POPUP_01" scroll="no">
    
        <% CALL Popup_generate("기업정보등록","정확하게 입력바랍니다.")%>
    <fieldset style="width:98%;align:center;">
        <legend>
        <form NAME="form" ACTION="sys_fcode_wrt.asp" method="POST" >
            <input type="hidden" name="fac_sabun"  size="30" value="<%=cid%>" >
	        <input type="radio" name="slt"  value="add" id="a1" checked="checked" /> <%=get_word(WA10038,"등록")%> 
	        <input type="radio" name="slt"  value="mod" id="a2" /> <%=get_word(WA10039,"수정")%> 
	        <input type="radio" name="slt"  value="del" id="a3" /> <%=get_word(WA10040,"삭제")%> 
        </legend>

    <div class="div_separate_5"></div>
	      <table border="0" cellpadding="2" cellspacing="0" width="99%">
		      <tr><td width="20%" align=right>계열사코드</td> 
		          <td width="30%"><input type="text" name="fac_ucd" maxlength="07"  size="6" <% If fac_ucd <> "" then %> value="<%=fac_ucd%>" readonly style="background-color:dedede;" <% else %> value="" <% end if %>   onKeyPress="return handleEnter(this,event);" class="input" ></td>
		          <td align=right width="20%">공장코드/구분명</td> 
		          <td width="30%"><input type=text class=input name="fac_fgubun"  maxlength=3 style="width:30px;" onkeypress="return handleEnter(this,event);" value="<%=fac_fgubun%>"> 
		          <input type=text class=input name="fac_fgubun_nm"  style="width:98px;" onkeypress="return handleEnter(this,event);" value="<%=fac_fgubun_nm%>"></td>
		      </tr>      
              <tr><td align=right width="20%">계열사명</td> 
		          <td width="30%"><input type=text class=input name="fac_unm" style="width:130px;" onkeypress="return handleEnter(this,event);" value="<%=fac_unm%>"></td>
		          <td align=right width="20%">대표자명</td> 
		          <td width="30%"><input type=text class=input name="fac_captin"  style="width:130px;" onkeypress="return handleEnter(this,event);" value="<%=fac_captin%>"></td>
		      </tr>	
		      <tr><td align=right>사업자번호</td> 
		          <TD><input type=text class=input name="fac_bunho" style="width:130px;"  onkeypress="return handleEnter(this,event);" value="<%=fac_bunho%>"></td>
		          <TD align=right>법인번호</td> 
		          <TD><input type=text class=input name="fac_bupin" style="width:130px;" onkeypress="return handleEnter(this,event);" value="<%=fac_bupin%>"></td>
		      </tr>
              <tr><TD  align="right">주소</td>
                  <TD colspan=3><table border=0 cellpadding=0 cellspacing=0>
                      <tr><TD width=180><input type="text" name="zipp1" value="<%=left(fac_post,3) %>" class=input  style="width:50px;" onKeyPress="return handleEnter(this,event);"> - <input type="text" name="zipp2" value="<%=right(fac_post,3) %>" class=input  style="width:50px;" onKeyPress="return handleEnter(this,event);"></td>
			              <TD width="65%" style="padding-left:5px;"><input type=button onclick="javascript:fncZipSearch()" value="주소검색" class=btn></td></tr>
			          <tr><td colspan=2 height=2></td></tr>
			          <tr><TD colspan=2><input type="text" name="fac_address" value="<%=fac_address%>" style="width:360px;" class=input onKeyPress="return handleEnter(this,event);"></td></tr>
			         </table>
                  </td>
              </TR>		      
		      <TR><TD align=right>대표전화</TD> 
		          <TD><input type=text class=input name="fac_tel" style="width:130px;" onkeypress="return handleEnter(this,event);" value="<%=fac_tel%>"></TD>
		          <TD align=right>팩스번호</TD> 
		          <TD><input type=text class=input name="fac_fax" style="width:130px;" onkeypress="return handleEnter(this,event);" value="<%=fac_fax%>"></TD>
		      </TR>		      
		      <TR><TD align=right>업태</TD> 
		          <TD><input type=text class=input name="fac_uptae" style="width:130px;" onkeypress="return handleEnter(this,event);" value="<%=fac_uptae%>"></TD>
		          <TD align=right>종목</TD> 
		          <TD><input type=text class=input name="fac_jongmok" style="width:130px;" onkeypress="return handleEnter(this,event);" value="<%=fac_jongmok%>"></TD>
		      </TR>		      
		      <tr><td align=right>CODEPAGE</td> 
		          <td><input type=text class=input name="fac_codepage" style="width:130px;" value="<%=fac_codepage%>"></td>
		          <td align=right>언어</TD> 
		          <td><input type=text class=input name="fac_lang"  style="width:130px;" value="<%=fac_lang%>"></td>
		      </tr>		      
		      <tr><td align=right>그룹웨어도메인</TD> 
		          <td><input type=text class=input name="fac_gw_domain"  style="width:130px;" value="<%=fac_gw_domain%>"></td>
		          <td align=right>메일도메인</TD> 
		          <td><input type=text class=input name="fac_mail_domain" style="width:130px;" value="<%=fac_mail_domain%>"></td>
		      </tr>	
			  </table>
    </fieldset>

        <div class="div_separate_10"></div>

        <table border="0" cellpadding="1" cellspacing="1" width="100%">
            <tr><td style="padding-left:15px;color:blue;padding-left:10px;" <%=css1%> id="msg"></td>
                <td align="right" style="padding-right:15px;">
                    <input type="button" onclick="set_submit();"  value=" 확 인 " class="btn" > 
                    <input type="button" value=" 닫 기 "  onclick="self.close();" class="btn"></td></form>
            </tr>                
        </table>

    <script type="text/javascript" src="/INC/js/ajax_script.js"></script>
    <script language="javascript" type="text/javascript">
    <!--
    <%  if slt = "" then    %>
                document.all("a1").checked  = true;   // 등록 Radio Button 을  checked 시킴
		        document.all("a2").disabled = true;   // 수정 Radio Button 을  disabled 시킴
		        document.all("a3").disabled = true;   // 삭제 Radio Button 을  disabled 시킴
    <%  elseif slt = "조회"  then  %>
		        document.all("a1").disabled = true;   // 등록 Radio Button 을  disabled 시킴
		        document.all("a2").checked  = true;   // 수정 Radio Button 을  checked 시킴
    <%    end if  %>

        function fncZipSearch() {
            zipReturn = ReturnSub;
            url = "zipcode.asp"
            centerWindow(url, "Addr_Search", "550", "430", "no")
        }

        function ReturnSub(zip1, zip2, address) {
            document.form.zipp1.value = zip1;
            document.form.zipp2.value = zip2;
            document.form.fac_address.value = address;
            document.form.fac_address.focus();
            return;
        }

        function set_submit() {

            // 등록/수정/삭제  선택  Radio 버턴 처리 부분
            if (document.getElementById("a1").checked == true) {
                var slt = "ADD";
            } else if (document.getElementById("a2").checked == true) {
                var slt = "MOD";
            } else if (document.getElementById("a3").checked == true) {
                var slt = "DEL";
            }

            //  입력 필드 유효성 확인한다.
            if (err_msg(form, "fac_ucd", "계열사 코드값은 반드시 입력해 주세요.") == "err") return;
            if (err_msg(form, "fac_fgubun", "계열사 공장 코드값은 반드시 입력해 주세요.") == "err") return;

            //  등록/수정/삭제  쿼리를 완성한다.

            var fac_post = form.zipp1.value + form.zipp2.value;
             
            var strSQL = "exec sp_InUpDel_Proc "     //  등록,수정,삭제에 해당하는 Stored Procedure (sp_InUpDel_Proc)
	                   + " @slt = '" + slt + "', "  //  등록,수정,삭제 구분 내용(slt)
	                   + " @tblName = 'Admin_fac_gubun',"     //  해당 테이블명
	                   + " @valueStr = 'N''" + form.fac_ucd.value + "'',''" + form.fac_unm.value + "'', "
                       + " N''" + form.fac_captin.value + "'',''" + form.fac_bunho.value + "'',"
                       + " N''" + form.fac_bupin.value + "'',N''" + fac_post + "'',"
	                   + " N''" + form.fac_address.value + "'',''" + form.fac_tel.value + "'',"
                       + "  ''" + form.fac_fax.value + "'',N''" + form.fac_uptae.value + "'',"
                       + "  ''" + form.fac_jongmok.value + "'',''" + form.fac_codepage.value + "'',"
                       + " N''" + form.fac_lang.value + "'',N''" + form.fac_gw_domain.value + "'',"
                       + " N''" + form.fac_mail_domain.value + "'',getdate(),getdate(),N''" + form.fac_sabun.value + "'',"
                       + " N''" + form.fac_fgubun.value + "'',N''" + form.fac_fgubun_nm.value + "'',',"
		               + " @fdName = 'fac_cd, fac_nm, fac_captin, fac_bunho, fac_bupin, fac_post, fac_address,"
                       + " fac_tel, fac_fax, fac_uptae,fac_jongmok, fac_codepage, fac_lang, fac_gw_domain, fac_mail_domain, "
		               + "  fac_wr_date, fac_ed_date, fac_sabun,fac_fgubun,fac_fgubun_nm,', "
		               + " @condStr = 'fac_cd=''" + form.fac_ucd.value + "'' and fac_fgubun =''" + form.fac_fgubun.value + "''', "         //   수정시 조건을 넣어준다.(@condStr)
		               + " @addCond = 'fac_bunho = N''" + form.fac_bunho.value + "''', "       // 등록시 조건을 넣어준다(@addCond)
		               + " @maxField = '', "     // 최대값을 구하는 필드 (@maxField)
		               + " @maxCond = '', "    // 최대값을 구하기 위한 조건 내용(@maxCond)
		               + " @delCond = 'DELETE Admin_fac_gubun where fac_cd = ''" + form.fac_ucd.value + "'''";   // 삭제시 조건을 넣어준다(@delCond)

            // 쿼리를 실행하기 위하여 RSExecute 한다.
            var Result = RSExecute("sys_fcode_wrt.asp", "set_process", strSQL);

            // 결과 값을 반환하여 메시지 출력한다.
            func_msgSetProcess(Result, '');
        }
    //-->
    </script>

    <!--#include Virtual = "/INC/inc_classTscript.ASP" -->
	<!--#include Virtual = "/INC/_ScriptLibrary/rs.asp"  -->	
    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
    <SCRIPT language="JavaScript">
    <!--
        self.moveTo((screen.availWidth - 500) / 2, (screen.availHeight - 450) / 2)
        self.resizeTo(500, 450)
    //-->
    </SCRIPT>