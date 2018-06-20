    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 프로그램 구성 메뉴 관리                              			--> 
    <!-- 내용: Mlevel 구성 관리                                                     --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 13일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <Base target="c">
    <BODY CLASS="BODY_01" SCROLL="no" onload="resize_scroll(35);" onresize="resize_scroll(35);" style="background-color:#C3DAF9;margin:0;border:0;">
    
    <div style="height:25;text-align:right;">
    <input type=button value="새로고침" onclick="document.location.reload();;" class=btn style="width:90;">
    
    </div>
    
    
    <div class="scr_01" id="scr_01" style="border:1px solid #6593CF;">
    <table border="0" cellpadding="0" cellspacing="0"  width="100%" height="100%"  bgcolor="white">
	    <Tr><td valign="top" style="padding-left:10px;padding-top:10px;"> 
	    <%
	    Dim oTree, ParentID, Cn, Rs
    	
	    Set oTree = Server.CreateObject("obout_ASPTreeview_2.Tree")

        oTree.foldericons   = "../gware/inc/tree/icons"
        oTree.folderscript  = "../gware/inc/js"
        oTree.folderstyle   = "../gware/inc/tree/win2003"	
            	
        
	    oTree.AddRootNode "DONGAERP SYSTEM" 
        oTree.KeepLastExpanded = 5
        Width = "200px"



	    Set Rs = Server.CreateObject("ADODB.Recordset")
	    Rs.Open "SELECT id,pid,name,expanded FROM mlevel where pgid = 9999 ORDER BY re_level, pid, dis_seq ,id", syscon, 0, 3

	    Do Until Rs.EOF
		    If IsNULL(Rs("pid")) or Rs("pid") = 0 Then 
			    ParentID = "root"
		    Else
			    ParentID = "id" & trim(Rs("pid"))  
		    End If

		    img="foldericon.png"

		    If IsNULL(rs("EXPANDED")) then		
		    EXPANDED=0
		    else
		    EXPANDED=rs("EXPANDED")
		    end if
            
		    oTree.Add ParentID, "id" & trim(Rs("id")) , "<A href=sys_menu_cont.asp?pID="& Rs("id") &" target='c' style='font-size:12px;font-family:arial;'>"& Server.HTMLEncode(Rs("name")) &"</a>", Rs("EXPANDED"), img
		    oTree.SelectedID = Request.QueryString("id")

		    Rs.MoveNext
	    Loop

        oTree=replace(oTree.HTML(),"<div style=""font:11px verdana; background-color:white; border:3px solid #cccccc; color:#666666; width:160px; text-align:center;"" align=center>&nbsp;<br>","")
        oTree = replace(oTree,"<a href=""http://www.ASPTreeView.com"">ASPTreeView.com</a><br>&nbsp;<br>Evaluation has <span style=""display:none;"">","")
        oTree = replace(oTree,"</span>expired.<br><a href=""http://www.obout.com/t2/purchase.aspx#sale"">Info...</a><br>","")

        oTree = replace(oTree,"??","")
        oTree = replace(oTree,"?&nbsp;","")
	            	
	    Response.Write oTree

	    Set oTree = Nothing
	    Rs.Close
    %>
    </td></tr>
    </table>
    </div>
    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->

