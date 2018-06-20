    <!-- ========================================================================== -->
    <!-- 프로그램 : 세금계산서 - 수금전표 링크                     			        --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2008년 09월 24일                                                --> 
    <!-- 내    용 : 세금계산서 - 수금전표 링크									    --> 
    <!-- ========================================================================== -->		    

	<!--#include Virtual = "/INC1/INC_TOP.ASP"  -->
    <%
    CSS = " style='font-size:12px;font-family:arial;' "
    
    
    m_idx           = request("m_idx")
    m_gubun         = trim(request("m_gubun"))
    m_factcd        = trim(request("m_factcd"))
    m_UserName      = trim(request("m_UserName"))
    
    m_machineName   = trim(request("m_machineName"))
    m_macAddress    = trim(request("m_macAddress"))
    m_Workgroup     = trim(request("m_Workgroup"))
    m_Ipaddress     = trim(request("m_Ipaddress"))
    m_FixedHddcnt   = request("m_FixedHddcnt")
    m_cdromcnt      = request("m_cdromcnt")
    m_NetDrivecnt   = request("m_NetDrivecnt")
    m_CPUVender     = trim(request("m_CPUVender"))
    m_CPUID         = trim(request("m_CPUID"))
    m_CPUType       = trim(request("m_CPUType"))
    m_CPUFamily     = request("m_CPUFamily")
    m_CPUModel      = request("m_CPUModel")
    m_CPUStepping   = request("m_CPUStepping")
    m_primaryHDDid  = trim(request("m_primaryHDDid"))
    m_OS            = trim(request("m_OS"))
    m_VirusChaser   = trim(request("m_VirusChaser"))
    m_Bizmate       = trim(request("m_Bizmate"))
    m_office       = trim(request("m_office"))
    method          = trim(request("method"))
    m_notebook      = trim(request("m_notebook"))
    m_remark      = trim(request("m_remark"))

 
    if  method = "process" then
    
        sql = "select * from  machine where rtrim(m_macaddress) = '"& trim(m_macAddress) &"' "
        set subrs = syscon.execute(sql)
        
        if  not subrs.eof then
        
        
            yymmdd = now()
            yy = year(yymmdd)
            mm = right("0" & month(yymmdd),2)
            dd = right("0" & day(yymmdd),2)
            hh = right("0" & hour(yymmdd),2)
            mi = right("0" & minute(yymmdd),2)
            ss = right("0" & second(yymmdd),2)
            
            m_code = trim(m_gubun) & yy & mm & dd  & hh & mi & ss        
            
            sql = "update machine set "
            sql = sql &" m_gubun        = '"& m_gubun &"' , "
            
            if  subrs("m_code") = "" or isnull(subrs("m_code")) then
                sql = sql &" m_Code     =  '"& m_code &"' , "
            end if
            
            sql = sql &" m_factcd       = '"& m_factcd &"' , "
            sql = sql &" m_UserName     = '"& m_UserName &"' , "            
            sql = sql &" m_machineName  = '"& m_machineName &"' , "
            sql = sql &" m_macAddress   = '"& m_macAddress &"' , "
            sql = sql &" m_Workgroup    = '"& m_Workgroup &"' , "
            sql = sql &" m_Ipaddress    = '"& m_Ipaddress &"' , "
            sql = sql &" m_FixedHddcnt  = 0 , "
            sql = sql &" m_cdromcnt     = 0 , "
            sql = sql &" m_NetDrivecnt  = 0 , "
            sql = sql &" m_CPUVender    = '"& m_CPUVender &"' , "
            sql = sql &" m_CPUID        = '"& m_CPUID &"' , "
            sql = sql &" m_CPUType      = '"& m_CPUType &"' , "
            sql = sql &" m_CPUFamily    = "& m_CPUFamily &" , "
            sql = sql &" m_CPUModel     = "& m_CPUModel &" , "
            sql = sql &" m_CPUStepping   = "& m_CPUStepping &" , "
            sql = sql &" m_primaryHDDid = '"& m_primaryHDDid &"' , "
            sql = sql &" m_OS           = '"& m_OS &"' , "
            sql = sql &" m_VirusChaser  = '"& m_VirusChaser &"' , "
            sql = sql &" m_Bizmate      = '"& m_Bizmate &"',  "
            sql = sql &" m_Office       = '"& m_Office &"',  "
            sql = sql &" m_notebook     = '"& m_notebook &"',  "
            sql = sql &" m_remark       = '"& m_remark &"',  "
            sql = sql &" userid         = '"& Request.Cookies("id") &"', moddt = getdate() "
            sql = sql &" where rtrim(m_macaddress) = '"& trim(m_macAddress) &"' "
            
            message = "수정되었습니다"

        else 
            
            yymmdd = now()
            yy = year(yymmdd)
            mm = right("0" & month(yymmdd),2)
            dd = right("0" & day(yymmdd),2)
            hh = right("0" & hour(yymmdd),2)
            mi = right("0" & minute(yymmdd),2)
            ss = right("0" & second(yymmdd),2)
            
            m_code = trim(m_gubun) & yy & mm & dd  & hh & mi & ss
        
            sql = " insert into machine ( m_code,m_gubun, m_factcd, m_UserName, m_machineName, m_macAddress,"
            sql = sql &"  m_Workgroup, m_Ipaddress, m_FixedHddcnt, m_cdromcnt, m_NetDrivecnt, m_CPUVender,"
            sql = sql &"  m_CPUID, m_CPUType, m_CPUFamily, m_CPUModel, m_CPUStepping, m_primaryHDDid, m_OS,m_notebook,"
            sql = sql &"  m_VirusChaser, m_Bizmate,m_Office,m_remark, userid, regdt )  values ('"& m_code &"','"& m_gubun &"','"& m_factcd &"','"& m_UserName &"',"
            sql = sql &"  '"& m_machineName &"', '"& m_macAddress &"','"& m_Workgroup &"', '"& m_Ipaddress &"',"
            sql = sql &"  0,0, 0, '"& m_CPUVender &"',"
            sql = sql &"  '"& m_CPUID &"', '"& m_CPUType &"', "& m_CPUFamily &", "& m_CPUModel &", "& m_CPUStepping &", '"& m_primaryHDDid &"', '"& m_OS &"',"
            sql = sql &"  '"& m_notebook &"','"& m_VirusChaser &"', '"& m_Bizmate &"','"& m_Office &"', '"& m_remark &"', '"& id &"', getdate() ) "
            
            message = "등록되었습니다"
       
        end if 
        
        syscon.execute(sql)                
    
    
    elseif method = "delete" then
    
        sql = "select * from  machine where m_idx = "& m_idx
        set subrs = syscon.execute(sql)
        
        if  not subrs.eof then
        
            syscon.execute("delete machine where  m_idx = "& m_idx)
        end if
        
        message = "삭제되었습니다"
                
    end if
    
    
    if  m_idx <> "" then
    
	    sql = "select a.*,c.gubun_name as gubunnm,d.gubun_name as gubunnm1 from machine a "
	    sql = sql &" left join dongaerp..gubun1 c on c.gubun_code = 'SS' and c.gubun_Seq = rtrim(substring(a.m_factcd,3,2)) "
	    sql = sql &" left join dongaerp..gubun1 d on d.gubun_code = 'WB' and d.gubun_Seq = rtrim(substring(a.m_gubun,3,2)) "
	    sql = sql &" where m_idx = "& m_idx
	    set rs  = syscon.execute(sql)
    	
	    if  not rs.eof then
    	    
            m_idx           = rs("m_idx")
            m_gubun         = trim(rs("m_gubun"))
            m_code          = trim(rs("m_code"))
            m_factcd        = trim(rs("m_factcd"))
            m_UserName      = trim(rs("m_UserName"))
            
            m_machineName   = trim(rs("m_machineName"))
            m_macAddress    = trim(rs("m_macAddress"))
            m_Workgroup     = trim(rs("m_Workgroup"))
            m_Ipaddress     = trim(rs("m_Ipaddress"))
            m_FixedHddcnt   = rs("m_FixedHddcnt")
            m_cdromcnt      = rs("m_cdromcnt")
            m_NetDrivecnt   = rs("m_NetDrivecnt")
            m_CPUVender     = trim(rs("m_CPUVender"))
            m_CPUID         = trim(rs("m_CPUID"))
            m_CPUType       = trim(rs("m_CPUType"))
            m_CPUFamily     = rs("m_CPUFamily")
            m_CPUModel      = rs("m_CPUModel")
            m_CPUStepping   = rs("m_CPUStepping")
            m_primaryHDDid  = trim(rs("m_primaryHDDid"))
            m_OS            = trim(rs("m_OS"))
            m_VirusChaser   = trim(rs("m_VirusChaser"))
            m_Bizmate       = trim(rs("m_Bizmate"))
            m_office        = trim(rs("m_office"))
            m_notebook      = trim(rs("m_notebook"))
            m_remark        = trim(rs("m_remark"))

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

    <table border=0 cellspacing=0 cellpadding=2 width="98%">
    <tr><td align="center">

		    <fieldset style="height:120;background-color:white;">
            <table border=0 cellspacing=0 cellpadding=2 width="98%">
                <tr><form name="wform" action="it_machinew.asp" method="post">
                    <input type="hidden" name="m_idx" value="<%=m_idx%>">
                    <input type="hidden" name="method" value="process">
                    <input type="hidden" name="m_notebook" value="<%=m_notebook%>">
                    <input type="hidden" name="m_Bizmate" value="<%=m_Bizmate%>">
                    <input type="hidden" name="m_VirusChaser" value="<%=m_VirusChaser%>">
                    <input type="hidden" name="m_office" value="<%=m_office%>">
                    <td align="center">구분</td>
                    <td><table border=0 cellpadding=1 cellspacing=0>
                        <tr>
                            <td width="50"><% CALL FAC_slt1("","m_factcd","SS",""& m_factcd &"") %></td>
                            <td width="50"><% CALL FAC_slt1("","m_gubun","WB",""& m_gubun &"") %></td>
                            <td><input type="checkbox" id="m_notecheck" name="m_notecheck" <% if m_notebook = "1" then %> checked <% end if %>/> 노트북</td>
                        </tr>
                        </table></td>
                </tr>
                <tr>
                    <td align="center">코드</td>
                    <td><table border=0 cellpadding=0 cellspacing=0 style="width:100%;">
                        <tr>
                            <td width="150"><input type="text" name="m_code" class="input" style="width:98%;" <%=CSS%> value="<%=m_code%>"/></td>
                            <td width="80" align=center>사용자</td>
                            <td width="150"><input type="text" name="m_userName" class="input" style="width:98%;" <%=CSS%> value="<%=m_userName%>"/></td>
                            <td width="80" align=center>컴퓨터이름</td>
                            <td width="150"><input type="text" name="m_machineName" class="input" style="width:98%;" <%=CSS%> value="<%=m_machineName%>"/></td>
                        </tr>
                        </table></td>
                </tr>
                <tr>
                    <td  align="center">맥어드레스</td>
                    <td><input type="text" name="m_macAddress" class="input" style="width:98%;" <%=CSS%> value="<%=m_macAddress%>"/></td>
                </tr>
                <tr>
                    <td  align="center">WORKGROUP</td>
                    <td><input type="text" name="m_Workgroup" class="input" style="width:98%;" <%=CSS%> value="<%=m_Workgroup%>"/></td>
                </tr> 
                <tr>
                    <td  align="center">IP Address</td>
                    <td><input type="text" name="m_Ipaddress" class="input" style="width:98%;" <%=CSS%> value="<%=m_Ipaddress%>"/></td>
                </tr>
                <tr>
                    <td  align="center">HDD 수</td>
                    <td><table border=0 cellpadding=0 cellspacing=0>
                        <tr>
                            <td width="50"><input type="text" name="m_FixedHddcnt" class="input" style="width:98%;" <%=CSS%> value="<%=m_FixedHddcnt%>"/></td>
                            <td width="80" align=center>CD-ROM 수</td>
                            <td width="50"><input type="text" name="m_cdromcnt" class="input" style="width:98%;" <%=CSS%> value="<%=m_cdromcnt%>"/></td>
                            <td width="80" align=center>NetDrive</td>
                            <td width="50"><input type="text" name="m_NetDrivecnt" class="input" style="width:98%;" <%=CSS%> value="<%=m_NetDrivecnt%>"/></td>
                        </tr>
                        </table></td>
                </tr>
                <tr>
                    <td  align="center">CPU Vender</td>
                    <td><input type="text" name="m_CPUVender" class="input" style="width:98%;" <%=CSS%> value="<%=m_CPUVender%>"/></td>
                </tr> 
                <tr>
                    <td  align="center">CPU ID</td>
                    <td><input type="text" name="m_CPUID" class="input" style="width:98%;" <%=CSS%> value="<%=m_CPUID%>"/></td>
                </tr>
                <tr>
                    <td  align="center">CPU Type</td>
                    <td><input type="text" name="m_CPUType" class="input" style="width:98%;" <%=CSS%> value="<%=m_CPUType%>"/></td>
                </tr> 
                <tr>
                    <td  align="center">CPU Family</td>
                    <td><input type="text" name="m_CPUFamily" class="input" style="width:98%;" <%=CSS%> value="<%=m_CPUFamily%>"/></td>
                </tr>
                <tr>
                    <td  align="center">CPU Model</td>
                    <td><input type="text" name="m_CPUModel" class="input" style="width:98%;" <%=CSS%> value="<%=m_CPUModel%>"/></td>
                </tr>
                <tr>
                    <td  align="center">CPU Stepping</td>
                    <td><input type="text" name="m_CPUStepping" class="input" style="width:98%;" <%=CSS%> value="<%=m_CPUStepping%>"/></td>
                </tr>
                <tr>
                    <td  align="center">HDD ID</td>
                    <td><input type="text" name="m_primaryHDDid" class="input" style="width:98%;" <%=CSS%> value="<%=m_primaryHDDid%>"/></td>
                </tr>  
                <tr>
                    <td  align="center">OS</td>
                    <td><input type="text" name="m_OS" class="input" style="width:98%;" <%=CSS%> value="<%=m_OS%>"/></td>
                </tr> 
                <tr>
                    <td  align="center">설치</td>
                    <td><input type="checkbox" name="m_VirusChaser_chk" <% if m_VirusChaser = "True" then %> checked <% end if %>/> 바이러스체이서
                        <input type="checkbox" name="m_Bizmate_chk" <% if m_Bizmate = "True"  then %> checked <% end if %>/> 메신저
                        <input type="checkbox" name="m_office_chk" <% if m_office = "True"  then %> checked <% end if %>/> MS오피스
                    
                    </td>
                </tr>    
                <tr>
                    <td  align="center">비고</td>
                    <td><input type="text" name="m_remark" class="input" style="width:98%;" <%=CSS%> value="<%=m_remark%>"/></td>
                </tr>                          
		                                 
            </table>
            </fieldset>  
            
             <table border=0 cellspacing=0 cellpadding=0 width="98%" style="margin-top:5px;">
                <tr>
                    <td style="padding-left:5px;"><input type="button" value=" 삭 제 " style="width:120;height:40;" class="btn" onclick="ondelete();"/></td>
                    <td style="text-align:right;font-family:arial;font-size:12px;">
                        
                        <input type="button" value=" 저 장 " style="width:120;height:40;" class="btn" onclick="onsave();"/>
                        <input type="button" value=" 닫 기 " style="width:120;height:40;" class="btn" onclick="self.close();"/>
                    </td></form>
                </tr>           
            </table>            
              
    </td></tr>
    </table>

    <SCRIPT TYPE="TEXT/JAVASCRIPT">

    var form = document.wform;
   
    function onsave()
    {
        if(form.m_notecheck.checked==true){
            form.m_notebook.value = "1"
        }else{
            form.m_notebook.value = "0"
        }
        if(form.m_VirusChaser_chk.checked==true){
            form.m_VirusChaser.value = "1"
        }else{
            form.m_VirusChaser.value = "0"
        }
        if(form.m_Bizmate_chk.checked==true){
            form.m_Bizmate.value = "1"
        }else{
            form.m_Bizmate.value = "0"
        }
        if(form.m_office_chk.checked==true){
            form.m_office.value = "1"
        }else{
            form.m_office.value = "0"
        }
        
        form.submit();
    
    }
    
    
    function ondelete()
    {
        if(form.m_idx.value==""){
            alert("삭제할 데이터가 없습니다.");
            return;
        }
        
        form.method.value="delete";
        form.submit();
    
    }
    </SCRIPT>
   
    <!--#include Virtual = "/INC1/INC_FOOT.ASP" -->    