    <!--#include Virtual = "/INC1/INC_TOP.ASP"  -->
    <!--#include Virtual = "/INC1/INC_BODY.ASP"  -->
    <%
    Response.CharSet = "UTF-8"
    
    css1 = " style='font-family:arial;font-size:12px;' "
    css2 = " type=text class=input style='font-family:arial;font-size:12px;width:98%;text-align:right;' "
    css3 = " type=text class=input style='font-family:arial;font-size:12px;width:98%;text-align:right;background-color:#dedede;' readonly=true"
    css4 = " type=text class=input style='font-family:arial;font-size:12px;width:98%;text-align:right;background-color:#ccffff;' readonly=true"
    
    m_factcd = request("m_factcd")
    m_gubun = request("m_gubun")
    
	sql = "select a.*,c.gubun_name as gubunnm,d.gubun_name as gubunnm1 from machine a "
	sql = sql &" left join dongaerp..gubun1 c on c.gubun_code = 'SS' and c.gubun_Seq = rtrim(substring(a.m_factcd,3,2)) "
	sql = sql &" left join dongaerp..gubun1 d on d.gubun_code = 'WB' and d.gubun_Seq = rtrim(substring(a.m_gubun,3,2)) "
	sql = sql &" where m_ipaddress <> '' "
	
	if  m_gubun <> "" then
	    sql = sql &" and m_gubun = '"& m_gubun &"' "
	end if
	
	if  m_factcd <> "" then
	    sql = sql &" and m_factcd = '"& m_factcd &"' "
	end if
	
	

	sql = sql &" order by m_Ipaddress "
	set rs  = syscon.execute(sql)
	
    if  not rs.eof then    
        %>
	    <table id='datalist' border="0" cellpadding="0" cellspacing="0" width="1100" style="table-layout: fixed;">
	    <%
	    while not rs.eof
	    
		    nvcnt		=	nvcnt	+	1

		    select case rs("m_VirusChaser")
		    case true : m_VirusChaser = "O"
		    case false : m_VirusChaser = "X"
		    end select
		    
		    select case rs("m_Bizmate")
		    case true : m_Bizmate = "O"
		    case false : m_Bizmate = "X"
		    end select
		    
		    select case rs("m_office")
		    case true : m_office = "O"
		    case false : m_office = "X"
		    end select
		    
		    select case rs("m_notebook")
		    case "1" : m_notebook = "O"
		    case "0" : m_notebook = "X"
		    end select
		    
		    
	        %>
		    <tr  onclick="show_data('<%=rs("m_idx")%>');" style="cursor:hand;">
			    
			    <td align="center" width="05%" class="tbl_drw_00" <%=css1%> style="border-left:none;"><%=nvcnt%></td> 
			    <td align="center" width="06%" class="tbl_drw_01" <%=css1%>><%=rs("gubunnm")%>&nbsp;</td>
			    <td align="center" width="09%" class="tbl_drw_01" <%=css1%>><%=rs("m_userName")%>&nbsp;</td>
			    <td align="center" width="13%" class="tbl_drw_01" <%=css1%>><%=rs("m_code")%>&nbsp;</td>
			    <td align="center" width="05%" class="tbl_drw_01" <%=css1%>><%=m_notebook%>&nbsp;</td>
			    
			    <td align="center" width="15%" class="tbl_drw_03" <%=css1%>><%=rs("m_machineName")%></td>
			    <td align="left"   width="12%" class="tbl_drw_03" <%=css1%>><%=rs("m_macAddress")%></td>
			    <td align="center" width="10%" class="tbl_drw_01" <%=css1%>><%=rs("m_Ipaddress")%>&nbsp;</td>
			    <td align="center" width="10%" class="tbl_drw_01" <%=css1%>><%=rs("m_OS")%>&nbsp;</td>
			    <td align="center" width="05%" class="tbl_drw_08" <%=css1%>><%=m_VirusChaser%></td>
			    <td align="center" width="05%" class="tbl_drw_08" <%=css1%>><%=m_Bizmate%>&nbsp;</td>
			    <td align="center" width="05%" class="tbl_drw_08" <%=css1%>><%=m_office%>&nbsp;</td>
		    </tr>
	        <%     
	        
	        m_VirusChaser = ""
	        m_Bizmate = ""
	        m_office = ""
	        m_notebook = ""
	               
		rs.movenext
	    wend
	    rs.close : set rs = nothing
	    %>        
	    </table><br /><br />
        <%        
    end if
    %>

