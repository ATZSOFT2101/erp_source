    <!-- ========================================================================== -->
    <!-- 프로그램 : 세금계산서 - 수금전표 링크                     			        --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2008년 09월 24일                                                --> 
    <!-- 내    용 : 세금계산서 - 수금전표 링크									    --> 
    <!-- ========================================================================== -->		    

	<!--#include Virtual = "/INC1/INC_TOP.ASP"  -->
    <%
    CSS = " style='font-size:12px;font-family:arial;' "
    
    
	sid         =	Request("sid")
	idx         =	Request("idx")
    sabun		=	trim(Request("sabun"))
    gubun		=	trim(Request("gubun"))
    gubun1		=	trim(Request("gubun1"))
    subject		=	trim(Request("subject"))
    sung        =   trim(Request("sung"))
    
    content		=	trim(Request("content"))
    reply		=	trim(Request("reply"))
    sangtae		=	trim(Request("sangtae"))
    method		=	trim(Request("method"))

 
    if  method = "process" then
    
        sql = "select a.*,b.femail,b.tele1 FROM it_request a inner join profile b on id = sabun where a.idx = "& idx
        set rs = syscon.execute(Sql)
        
        if  not rs.eof then
        
            sql = "update it_request set "
            sql = sql &" sabun      ='"& sabun &"',"
            sql = sql &" gubun      ='"& gubun &"',"
            sql = sql &" gubun1      ='"& gubun1 &"',"
            sql = sql &" subject    ='"& subject &"',"
            sql = sql &" content    ='"& content &"',"
            sql = sql &" reply      ='"& reply &"',"
            sql = sql &" sangtae    ='"& sangtae &"'"
            sql = sql &"  "
            
            if  (sangtae = "1" or sangtae = "2") and rs("mailsend") = "0" then
                
                title = left(rs("writedate"),10) &"에 접수된 귀하의 전산요청내역에 대한 진행내역입니다."
                
                
                content = replace(content,vblf,"<br>")
                reply   = replace(reply,vblf,"<br>")

                msg = msg & "<style type=text/css>"
                msg = msg & "td {font-size:11;font-family:arial;}"
                msg = msg & ".tda {font-size:11;font-family:arial;color:#083772;height:20;background-color:#B5D2F7;}"
                msg = msg & ".tdb {font-size:12;font-family:arial;color:#000000;height:20;background-color:#ffffff;}"
                msg = msg & ".tdc {font-size:11;font-family:arial;color:#083772;height:20;background-color:#D5E4F2;}"
                msg = msg & ".tdd {font-size:18;font-family:arial;color:#000000;height:50;}"
                msg = msg & "</style>"                
                msg = msg & left(rs("writedate"),10) &"에 접수된 전산요청내역에 대한 진행내역입니다.<br><br>"
                msg = msg &"<table border=0 cellpadding=5 cellspacing=1 width=600 bgcolor=#6593CF>"
                msg = msg &" <tr bgcolor=#B5D2F7><td align=center>귀하의 요청내용</td></tr>"
                msg = msg &" <Tr bgcolor=#ffffff><td class=tdb>"& content &"<br>"
                msg = msg &" <tr bgcolor=#B5D2F7><td align=center>처리내용</td></tr>"
                msg = msg &" <Tr bgcolor=#ffffff><td class=tdb>"& reply &"<br>"
                msg = msg &"</table><br><br>"
                msg = msg & "처리내용에 대한 문의는 전산실(내선:273)으로 연락바랍니다."
                
                CALL mailsave(sid,Session("sung"),Session("email"),rs("femail"),title,msg)
            
            
                if  rs("tele1") <> "" then
                
		            select case sangtae
		            case "1" : sangtae_ = "처리"
		            case "2" : sangtae_ = "기각"
		            end select                
                
                    smsSql = "Insert into emma..em_tran "
                    smsSql = smsSql & " (tran_phone, tran_callback, tran_status, tran_date, tran_msg) values "
                    smsSql = smsSql & " ('"& rs("tele1") &"', '055-313-1800', '1',getdate(), '귀하의 전산요청이 "& sangtae_ &" 되었습니다.')"
                    conn.Execute(smsSql)	
                end if            
            
                sql = sql &" ,responsedate = getdate() , mailsend  = '1' "
            end if
            
            sql = sql &" where idx = "& idx 
            syscon.execute(sql)  
            
            message = "처리되었습니다"

                  
        end if
            
    end if
    
    if  method = "add" then
    
        sql = "insert into it_request (sabun,gubun,gubun1,subject,content,sangtae,mailsend,writedate) values "
        sql = sql &" ('"& sabun &"','"& gubun &"','"& gubun1 &"','"& subject &"','"& content &"','0','0',getdate()) "
        syscon.execute(Sql)

            
        message = "등록되었습니다"
            
    end if    
    
    if  idx <> "" then
	    sql = "select sung,a.*,c.gubun_name as gubunnm,d.gubun_name as gubunnm1 from it_request a "
	    sql = sql &" inner join profile b on a.sabun = b.id "
	    sql = sql &" left join dongaerp..gubun1 c on c.gubun_code = 'WC' and c.gubun_Seq = rtrim(substring(a.gubun,3,2)) "
	    sql = sql &" left join dongaerp..gubun1 d on d.gubun_code = 'WD' and d.gubun_Seq = rtrim(substring(a.gubun1,3,2)) "
	    sql = sql &" where a.idx = "& idx 
	    set rs  = syscon.execute(sql)	
    	
	    if  not rs.eof then
    	    
	        nvcnt		=	nvcnt	+	1
	        idx         =	rs("idx")
	        sabun		=	trim(rs("sabun"))
	        gubun		=	trim(rs("gubun"))
	        gubun1		=	trim(rs("gubun1"))
	        subject		=	trim(rs("subject"))
	        sung        =   trim(rs("sung"))
    	    
	        content		=	trim(rs("content"))
	        reply		=	trim(rs("reply"))
	        sangtae		=	trim(rs("sangtae"))
	        mailsend    =	trim(rs("mailsend"))
	        wdate       =	trim(rs("writedate"))
	        rdate       =	trim(rs("responsedate"))
	        gubunnm     =   trim(rs("gubunnm"))
	        gubunnm1    =   trim(rs("gubunnm1"))
    	    
    	    
	        select case sangtae
	        case "0" : sangtae_ = "접수중"
	        case "1" : sangtae_ = "완료"
	        case "2" : sangtae_ = "기각"
	        end select
    	    
	        select case mailsend
	        case "0" : mailsend_ = "미전송"
	        case "1" : mailsend_ = "전송"
	        end select

	    end if
	end if
	%>
	
	<!--#include Virtual = "/INC1/INC_BODY.ASP" -->   
    <base target="_self" />
    <body CLASS="BODY_01" SCROLL="no" style="margin:0;border:0;background-color:#D8E4F3">
    
    <div id="msg" style="display:none;position:absolute;width:300;height:50;top:100;left:150;z-index:10000;">
        <table border="0" cellpadding="0" cellspacing="1" width="300" style="background-color:666666;" align="center">
            <tr><td height="40" style="background-color:#ffffff;text-align:center;" id=msgcont></td></tr>
        </table>
    </div>      
    
    
    <!-- HEAD 부분 시작 -->
    <table  style="margin:0;border:0;background-color:#D8E4F3" id=d2 >
        <tr><td>
                <table border="0" cellpadding="0" cellspacing="0" height=30>
                <tr>
				    <td width=40 align=center><img src="<%=img_path%>fol.gif" border=0 /></TD>
				    <td>완료 또는 기각을 선택하시고, 조치사항(기각사유)을 입력하시면 관련내역이 요청자에게 전송됩니다.</td>
		        </tr>
			    </table>
        </td></tr>
    </table>
    <!--  HEAD 부분 끝 -->
  

    <table border=0 cellspacing=0 cellpadding=2 width="98%">
    <tr><td align="center">

		    <fieldset style="height:120;background-color:white;">
            <table border=0 cellspacing=0 cellpadding=2 width="98%">
                <tr><form name="wform" action="it_requestw.asp" method="post">
                    <input type="hidden" name="idx" value="<%=idx%>">
                    <input type="hidden" name="sid" value="<%=sid%>">
                    <% if   method = "add" then %>
                    <input type="hidden" name="method" value="add">
                    <% else %>
                    <input type="hidden" name="method" value="process">
                    <% end if %>
                    
                    <td width="60" align="center" <%=css%>>SID</td>
                    <td style="color:red;" <%=css%>><%=SId%></td>
                </tr>
                <tr>
                    <td width="60" align="center">요 청 자</td>
                    <td><% call search_frm("insa06","","sabun","sung","searchvalue1",""& sabun &"",""& sung &"") %></td>
                </tr>
                <tr>
                    <td align="center">업무구분</td>
                    <td><% CALL FAC_slt1("","gubun","WC",""& gubun &"") %></td>
                </tr>
                <tr>
                    <td  align="center">제 &nbsp;&nbsp; &nbsp;&nbsp; 목</td>
                    <td><input type="text" name="subject" class="input" style="width:98%;" <%=CSS%> value="<%=subject%>"/></td>
                </tr>
                <tr>
                    <td  align="center">상 &nbsp;&nbsp; &nbsp;&nbsp; 태</td>
                    <td><input type="radio" name="sangtae" id="sangtae0" value="0" <% if sangtae="0" or sangtae="" then %>checked<% end if %>> 접수
	                    <input type="radio" name="sangtae" id="sangtae1" value="1" <% if sangtae="1" then %>checked<% end if %>> 완료
	                    <input type="radio" name="sangtae" id="sangtae2" value="2" <% if sangtae="2" then %>checked<% end if %>> 기각
	                </td>
                </tr>
                <tr>
                    <td  align="center">요청내용</td>
		            <td><textarea name="content" style="width:98%;height:120;" class=input <%=CSS%>><%=content%></textarea></td>
		        </tr>
                <tr>
                    <td  align="center">요청구분</td>
		            <td><% CALL FAC_slt1("","gubun1","WD",""& gubun1 &"") %></td>
		        </tr>
                <tr>
                    <td  align="center">처리내용<br />(기각내용)</td>
                   <td><textarea name="reply" style="width:98%;height:120;" class=input <%=CSS%>><%=reply%></textarea></td>
                </tr>                               
		                                 
            </table>
            </fieldset>  
            
             <table border=0 cellspacing=0 cellpadding=0 width="98%" style="margin-top:5px;">
                <tr>
                    <td style="color:red;padding-left:5px;"><%=message%></td>
                    <td style="text-align:right;font-family:arial;font-size:12px;">
                        <input type="button" value=" 저 장 " style="width:120;height:40;" class="btn" onclick="onsave();"/>
                        <input type="button" value=" 닫 기 " style="width:120;height:40;" class="btn" onclick="self.close();"/>
                    </td></form>
                </tr>           
            </table>            
              
    </td></tr>
    </table>

    <SCRIPT TYPE="TEXT/JAVASCRIPT">

    var form = document.form;
    
    document.getElementById("gubun").style.width = "223";
    document.getElementById("gubun1").style.width = "223";

    function onsave()
    {
        var form = document.wform;
        
        if(form.sabun.value==""){
            alert("사번이 존재하지 않습니다.\n\n다시 로그인해주세요.");
            return;
        }
        if(form.gubun.value==""){
            alert("업무구분을 선택해주세요");
            return;
        }
        
        if(form.subject.value==""){
            alert("제목을 입력해주세요");
            return;
        }
        if(form.content.value==""){
            alert("요청내용을 입력해주세요");
            return;
        }
    
        if((document.getElementById("sangtae0").checked==false)&&(form.reply.value=="")){
            alert("처리내용을 입력해주세요");
            return;
        }
        
        form.submit();
    
    }

    </SCRIPT>
   
    <!--#include Virtual = "/INC1/INC_FOOT.ASP" -->    