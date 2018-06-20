    <!--#include Virtual = "/INC1/INC_TOP.ASP"  -->
    <!--#include Virtual = "/INC1/INC_BODY.ASP"  -->
    <%
    Response.CharSet = "UTF-8"
    
    css1 = " style='font-family:arial;font-size:12px;' "
    css2 = " type=text class=input style='font-family:arial;font-size:12px;width:98%;text-align:right;' "
    css3 = " type=text class=input style='font-family:arial;font-size:12px;width:98%;text-align:right;background-color:#dedede;' readonly=true"
    css4 = " type=text class=input style='font-family:arial;font-size:12px;width:98%;text-align:right;background-color:#ccffff;' readonly=true"
    
    wdate1  = request("wdate1")
    wdate2  = request("wdate2")
    gubun   = request("gubun")
    gubun1  = request("gubun1")
    sangtae = request("sangtae")
    sabun   = request("sabun")
    
    wdate1 = wdate1 &"-01"
    if  mid(wdate2,6,2) = "02" then
        wdate2 = wdate2 &"-28"
    else
        wdate2 = DateSerial(year(wdate2), month(wdate2), 31)   
    end if
    
  
	sql = "select sung as insa_name,a.*,c.gubun_name as gubunnm,d.gubun_name as gubunnm1 from dasystem..it_request a "
	sql = sql &" left join dasystem..profile b on a.sabun = b.id "
	sql = sql &" left join dongaerp..gubun1 c on c.gubun_code = 'WC' and c.gubun_Seq = rtrim(substring(a.gubun,3,2)) "
	sql = sql &" left join dongaerp..gubun1 d on d.gubun_code = 'WD' and d.gubun_Seq = rtrim(substring(a.gubun1,3,2)) "
	sql = sql &" where writedate between '"& wdate1 &" 00:00:000' and '"& wdate2 &" 23:59:000' "
	
	if  gubun <> "" then
	    sql = sql &" and a.gubun = '"& gubun &"' "
    end if
    
	if  gubun1 <> "" then
	    sql = sql &" and a.gubun1 = '"& gubun1 &"' "
    end if
    
	if  sangtae <> "" then
	    sql = sql &" and a.sangtae = '"& sangtae &"' "
    end if
    
	if  sabun <> "" then
	    sql = sql &" and b.id = '"& sabun &"' "
    end if
     
	sql = sql &" order by writedate"
	'response.Write sql
	'response.end
	set rs  = conn.execute(sql)
	
    if  not rs.eof then    
        %>
	    <table id='datalist' border="0" cellpadding="0" cellspacing="0" width="1100" style="table-layout: fixed;">
	    <%
	    while not rs.eof
	    
		    nvcnt		=	nvcnt	+	1
		    idx         =	rs("idx")
		    sabun		=	trim(rs("sabun"))
		    gubun		=	trim(rs("gubun"))
		    gubun1		=	trim(rs("gubun1"))
		    subject		=	trim(rs("subject"))
		    sung        =   trim(rs("insa_name"))
		    
		    content		=	trim(rs("content"))
		    sangtae		=	trim(rs("sangtae"))
		    mailsend    =	trim(rs("mailsend"))
		    wdate       =	trim(rs("writedate"))
		    rdate       =	trim(rs("responsedate"))
		    gubunnm     =   trim(rs("gubunnm"))
		    gubunnm1     =   trim(rs("gubunnm1"))
		    
		    
		    select case sangtae
		    case "0" : sangtae_ = "접수중"
		    case "1" : sangtae_ = "<span style='color:blue;'>완료</span>"
		    case "2" : sangtae_ = "<span style='color:red;'>기각</span>"
		    end select
		    
		    select case mailsend
		    case "0" : mailsend_ = "미전송"
		    case "1" : mailsend_ = "전송"
		    end select
		    		    
	        %>
		    <tr style="cursor:hand;">
			    <td align="center" width="05%" class="tbl_drw_00" <%=css1%> style="border-left:none;">
                <a onclick="set_delete('<%=idx%>');" style='cursor:hand;'><img src='../image/btn_close.gif' border=0></a></td>
			    <td align="center" width="15%" class="tbl_drw_01" <%=css1%> onclick="show_data('<%=idx%>');" ><%=wdate%></td>
			    <td align="center" width="08%" class="tbl_drw_01" <%=css1%> onclick="show_data('<%=idx%>');" ><%=gubunnm%></td>
			    
			    <td align="center" width="07%" class="tbl_drw_03" <%=css1%> onclick="show_data('<%=idx%>');"  ><%=sung%></td>
			    <td align="left"   width="30%" class="tbl_drw_03" <%=css1%> onclick="show_data('<%=idx%>');" ><%=subject%></td>
			    <td align="center" width="09%" class="tbl_drw_01" <%=css1%> onclick="show_data('<%=idx%>');" ><%=gubunnm1%>&nbsp;</td>
			    <td align="center" width="06%" class="tbl_drw_01" <%=css1%> onclick="show_data('<%=idx%>');" ><%=sangtae_%>&nbsp;</td>
			    <td align="center" width="06%" class="tbl_drw_08" <%=css1%> onclick="show_data('<%=idx%>');" ><%=mailsend_%></td>
			    <td align="center" width="14%" class="tbl_drw_08" <%=css1%> onclick="show_data('<%=idx%>');" ><%=rdate%>&nbsp;</td>
		    </tr>
	        <%            
		rs.movenext
	    wend
	    rs.close : set rs = nothing
	    %>        
	    </table>
        <%        
    end if
    %>

