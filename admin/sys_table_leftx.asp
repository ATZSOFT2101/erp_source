    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!--#include Virtual = "/INC/INC_BODY.ASP"  -->
    <%
    Response.CharSet = "UTF-8"
    
    css1 = " style='font-family:arial;font-size:12px;' "
    css2 = " style='font-family:arial;font-size:12px;color:red;' "
    
    
    tabledb   = request("tabledb")
    table_gubun = request("table_gubun")
    
    
    SQL = "SELECT Table_idx, TableID, TableNM, B.Name ,a.table_gubun,convert(char(10),crdate,111) as crdate,gubun_name FROM SYS_Table A " 
	Sql = Sql & " left join "& tabledb &".dbo.sysobjects B on xtype = 'U' and Name=tableid"
	Sql = Sql & " left join gubun1 c on gubun_code = left(table_gubun,2) and gubun_seq = substring(table_gubun,3,2) "
	Sql = Sql & " where tabledb = '"& tabledb &"' "
	
    if  table_gubun <> "WA1" and table_gubun <> "" then
    
	    sql = sql &" and table_gubun = '"& table_gubun &"' "
    end if
	
	sql = sql &"order by Tableid asc "

	Set Rs = syscon.Execute(Sql)
    
    %>
    <table border="0" id='datalist' cellpadding="0" cellspacing="0" width="550" style="table-layout: fixed;">
	<%
	vcnt=1
	while not rs.eof

	table_idx   =  rs("table_idx")
	tableid		=  rs("tableid") 
	tablenm		=  rs("tablenm")
	table_idx   =  rs("table_idx")
	crdate      =  rs("crdate")
	tgubun      =  rs("table_gubun")
	gubun_name  =  rs("gubun_name")
	
	%>
	<tr style="cursor:hand;">
	    <td width="07%" class="tbl_drw_00" align="center" style="border-left:none;"><input type="checkbox" name="chkbox" value="<%=table_idx%>"></td>
	    <td width="07%" class="tbl_drw_01" align="center" <%=css2%> onclick="set_win('<%=tableid%>','<%=tabledb%>');"><%=vcnt%></td>
	    <td width="10%" class="tbl_drw_01" align="center"<%=css1%> ><%=gubun_name%>&nbsp;</td>
	    <td width="23%" class="tbl_drw_03" <%=css1%> onclick="loaddata('<%=tablenm%>','<%=tableid%>','<%=tabledb%>');"><%=ucase(tableid)%></td>
	    <td width="18%" class="tbl_drw_01" align="center"<%=css1%>><%=crdate%></td>
	    <td width="35%" class="tbl_drw_01" <%=css1%> onclick="loaddata('<%=tablenm%>','<%=tableid%>','<%=tabledb%>');"><%=tablenm%></td>
    </tr>
    <%
    rs.movenext
    vcnt=vcnt+1
    wend

    rs.close : set rs = nothing
    %>
    </table></form> <br /><br />    


