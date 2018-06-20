    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 사용자 권한 관리                                   			--> 
    <!-- 내용: 프로그램별 사용자 권한 구성 관리                                     --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 14일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    
    Server.ScriptTimeOut = 999999

	ReID	    = Request("REID")
    opt         = Request("opt")
	REID_NAME	= Request("REID_NAME")
    
	IF REQUEST("SLT") = "ADD" THEN
       TEMP =  SPLIT(REQUEST("chk_pgid"),",")
       CNT = 0

       response.Write "확인1"
   '    response.end 

       SYSCON.EXECUTE("DELETE FROM MENU_REF WHERE MR_PRO_ID = '"& REQUEST("ReID") &"'")

       
       IF   opt ="Y" THEN
            syscon.execute("delete menu_duty where mr_pro_id = '"& REQUEST("ReID") &"' ")
       END IF
              response.Write "확인2"
 	   for cnt=0 to (UBound(TEMP))
          if Not isNULL(TEMP) and TEMP(CNT) <> 0 then  
                    response.Write "확인3"
             SYSCON.EXECUTE("INSERT INTO MENU_REF(MR_PRO_ID, MR_MLVL_ID,MR_SEQ) VALUES('"& REQUEST("ReID") &"'," & TEMP(CNT) & ","& CNT & ")")
             
             
             IF  opt ="Y" THEN
             
                 subSQL = "exec sp_mlevel " & TEMP(CNT) &",'"& REQUEST("ReID") &"'  "
                 response.Write subsql
               '  response.end 
                 Set subRS=Syscon.execute(subSQL)
                 
                 While Not subRS.EOF
                    
                    sql = "select * From menu_duty where mr_pro_id = '"& REQUEST("ReID") &"' and mr_mlvl_id= '"& subrs("pgid") &"' "
                    response.Write sql & "<br>"
                    set crs = syscon.execute(sql)
                    
                    if  crs.eof then                    
                        SYSCON.EXECUTE("INSERT INTO MENU_DUTY (MR_MLVL_ID, MR_PRO_ID, MR_READ, MR_ADD ,MR_DELETE, MR_PRINT, MR_EXCEL, MR_EDIT, MR_APPRO) Values ("& subRs("pgid") &",'"& REQUEST("ReID") &"','1','1','1' ,'1' ,'1' ,'1','1' )")
                    
                    end if
                    
                     crs.close                
                 
                  subRs.Movenext
                 Wend

                 subRs.Close : SEt subRs = nothing
             END IF
             
          end if  	
	   next
        msg = "처리되었습니다"
	
    END IF    

    %>
  
	<BODY CLASS="POPUP_01" SCROLL="no" onload="resize_scroll(60);" onresize="resize_scroll(60);" style="background-color:#C3DAF9;margin:0;border:0;">
    
    <!-- HEAD 부분 시작 --> 
    <TABLE style="margin:0;width:100%;" border=0 bgcolor="#D8E4F3">
      <Tr><form name=form action="sys_duty_left1.asp" method="post">
		  <td>	
          <table cellspacing="0" cellpadding=0 border=0>
	         <tr><td><% CALL search_frm("sys02","성명","REID","REID_NAME","searchvalue",""& REID &"",""& REID_NAME &"") %></td>
	            <td ><input type=submit  value=" 검 색 " class=btn></td>
	         </tr>
	      </table>	
	    
	  </TD></TR>
    </TABLE>
    
	<!-- HEAD 부분 끝 --> 
    

	<div class="scr_01" id="scr_01" style="border:1px solid #6593CF;border-left:1px solid #6593CF;" >
    
	<table border="0" cellpadding="0" cellspacing="0"  width="100%" height="100%"  bgcolor="white">
	    
		<Tr><td valign="top" style="padding-left:10px;padding-top:10px;">  
	    
		<%
		Dim oTree, ParentID, Cn, Rs
		Set oTree = Server.CreateObject("obout_asptreeview_2.tree")

        oTree.foldericons   = "../gware/inc/tree/icons"
        oTree.folderscript  = "../gware/inc/js"
        oTree.folderstyle   = "../gware/inc/tree/Classic"	
		
	
		
		'oTree.AddRootNode "DONGAERP SYSTEM" 
		oTree.KeepLastExpanded = 5
		Width = "150px"
            
		Set Rs = Server.CreateObject("ADODB.Recordset")
		Rs.Open "SELECT A.id,A.pid,B.mr_mlvl_id,name,expanded FROM mlevel A left join MENU_REF B On A.id=B.mr_mlvl_id and mr_pro_id ='"& ReID &"' where pgid = 9999   ORDER BY re_level, pid ,id", syscon, 0, 3

		Do Until Rs.EOF
            If	IsNULL(Rs("pid")) or Rs("pid") = 0 Then 
				ParentID = "root"
			Else
				ParentID = "id" & trim(Rs("pid"))  
			End If

			img="foldericon.png"

			If  IsNULL(rs("EXPANDED")) then		
				EXPANDED=0
			else
				EXPANDED=rs("EXPANDED")
			end if
			
			mr_mlvl_id = rs("mr_mlvl_id")
			pid        = rs("pid")

			IF  isNULL(mr_mlvl_id) THEN
				html = "<input type='checkbox' name='chk_pgid' value="& Rs("id") &" style='height:13px; width:13px; margin:0px;'> "
			ELSE
				html = "<input type='checkbox' name='chk_pgid' value="& Rs("id") &" style='height:13px; width:13px; margin:0px;' checked> "
			END IF
			
			id = Rs("id")
		    
			oTree.Add ParentID, "id" & trim(Rs("id")) , html &"<A href='sys_duty_cont.asp?PID="& Rs("id") &"&REID="& REID &"' target='sub_c'>"& Rs("name") &"</a>", Rs("EXPANDED"), img
		   
			oTree.SelectedID = Request.QueryString("id")

			Rs.MoveNext
			Loop
			
			oTree=replace(oTree.HTML(),"<div style=""font:11px verdana; background-color:white; border:3px solid #cccccc; color:#666666; width:160px; text-align:center;"" align=center>&nbsp;<br><a href=""http://www.ASPTreeView.com"">ASPTreeView.com</a><br>&nbsp;<br>Evaluation has <span style=""display:none;"">","")
    	
			Response.Write oTree

			Set oTree = Nothing
			Rs.Close
			%>
    
		</td></tr>
	</table>
	</div>
    
	<!-- Bottom 부분 시작 -->
    <TABLE CLASS=Sch_Header style="margin:0;background-color:666666;">
	  <TR><TD align=right>

        <table border="0" cellpadding="0" cellspacing="0">
		<tr><td><%=msg %></td>
		    <td align=right><input type="checkbox" name="opt" id="opt" value"Y" checked/> 세부권한 <input type="button" value="권한저장" onclick="this.value='처리중';disabled=true;duty_save();" class=btn></td></form></tr>
		</table>

	  </TD></TR>
    </TABLE>
	<!-- Bottom 부분 끝 -->     
   
	<!--#include Virtual = "/INC/INC_FOOT.ASP"              -->
    
    <SCRIPT TYPE="TEXT/JAVASCRIPT">
    
    function duty_count() {
	    var f = document.form;
	    var check_count = 0;
	    for(i = 0 ; i < f.elements.length; i++) {
		    if(f.elements[i].checked == true)
			    duty_count =duty_count + 1
		    }
	    return duty_count;
    }
    
    function duty_save() {

        var myform =	document.form;
        var param = "0" ;
        if (duty_count() <= 0) {
		        alert("선택된 메뉴가 없습니다\n\n적어도 하나이상을 선택해주세요")
		        return;
        }    
        if (typeof(form.chk_pgid.length) != "undefined") {  // 목록이 여러개일때
            for (var i = 0; i < form.chk_pgid.length; i ++) {
                if (form.chk_pgid[i].checked) {             // 체크된것만 param 변수에 저장
		        param = param + "," +  form.chk_pgid[i].value;
		    }
            }
        } else {    // 목록이 하나일때
            if (form.chk_pgid.checked) {
                param = param + "," +  form.chk_pgid.value;
            }
            if (form.chk_pgid.checked == false){ 
               alert("체크해 주세요"); 
               return false;
               }
        }
        
        if(document.getElementById("opt").checked==true){
            var opt = "Y"
        }else{
            var opt = "N"
        }
        location.href='sys_duty_left1.asp?opt='+opt+'&reid=<%=reid%>&slt=ADD&chk_pgid='+ param ;
    }
    </SCRIPT>    
