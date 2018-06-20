    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램 : 테이블/필드 관리                                 			        --> 
    <!-- 작 성 자 : 문성원(050407)                                                    --> 
    <!-- 작성일자 : 2006년 6월 08일                                                   --> 
    <!-- 내    용 : 테이블/필드 관리                                                   --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    gubun   = Request("gubun")
    FLd_idx = Request("FLd_idx")
    
	IF request("tableid") <> "" then

	   tableid	= request("tableid")
	   tablenm	= request("tablenm")
	   tabledb	= request("tabledb")

	   IF gubun = "rebuild_field" THEN
	   
	      'sys_FIELD 테이블과 syscolumns 테이블 비교하여 없는 필드는 Insert 시켜주자
          FDSQL = "USE "& TABLEDB &""	      
          FDSQL = FDSQL& " INSERT INTO SYS_FIELD (FLDDB,FLDTABLE,FLDNO,FLDNAME,FLDTYPE, FLDLENG, FLDNULL, FLDDESC, FLDDISPLAY) "
          FDSQL = FDSQL& " SELECT '"& TABLEDB &"','"& TABLEID &"',A.colorder,A.NAME, B.NAME,A.LENGTH,A.ISNULLABLE,'...','1'  FROM "& TABLEDB &".DBO.SYSCOLUMNS A "
          FDSQL = FDSQL& " inner JOIN "& TABLEDB &".DBO.SYSTYPES B ON A.XTYPE = B.XTYPE and B.NAME <> 'sysname' "
          FDSQL = FDSQL& " WHERE   ID=OBJECT_ID('"& TABLEID &"') AND NOT EXISTS( "
          FDSQL = FDSQL& "        SELECT FLDNAME,FLDTYPE,FLDNO, FLDLENG, FLDNULL  "
          FDSQL = FDSQL& "        FROM SYS_FIELD  C "
          FDSQL = FDSQL& "        WHERE  "
          FDSQL = FDSQL& "        A.NAME = C.FLDNAME AND  "
          FDSQL = FDSQL& "        B.NAME = C.FLDTYPE  AND  "
          FDSQL = FDSQL& "        A.LENGTH = C.FLDLENG  AND  "
          FDSQL = FDSQL& "        A.ISNULLABLE = C.FLDNULL  AND FLDTABLE='"& TABLEID &"' "
          FDSQL = FDSQL& "       )  "
          'RESPONSE.WRITE FDSQL
          SYSCON.EXECUTE(FDSQL)

	      'SYS_FIELD 테이블과 SYSCOLUMNS 테이블 비교하여 불필요한 필드는 DELETE 시켜주자
	      FDSQL = "DELETE SYS_FIELD WHERE FLDNAME = ("
          FDSQL = FDSQL& " SELECT FLDNAME FROM SYS_FIELD A "
          FDSQL = FDSQL& "       WHERE  FLDTABLE= '"& TABLEID &"' AND NOT EXISTS "
          FDSQL = FDSQL& " ( SELECT B.NAME, C.NAME,B.LENGTH,B.ISNULLABLE FROM "& TABLEDB &".DBO.SYSCOLUMNS B "
          FDSQL = FDSQL& " LEFT JOIN "& TABLEDB &".DBO.SYSTYPES C ON C.XTYPE = B.XTYPE WHERE "
          FDSQL = FDSQL& " ID = OBJECT_ID('"& TABLEID &"') AND B.NAME = A.FLDNAME )"
          FDSQL = FDSQL& " ) "
          SYSCON.EXECUTE(FDSQL)
        
       ELSEIF  gubun = "delete" THEN 
        
          syscon.execute("Delete sys_field where fld_idx ="& fld_idx )
       
	   END IF
	   
	   IF gubun = "rebuild_field" THEN
	   sql = "SET NOCOUNT ON "
	   slq = sql & " USE DASYSTEM "
	   END IF
	   
	   Sql = sql &" SELECT * FROM sys_field WHERE FLDTable= '"& tableid &"' and FLDDb='"& tabledb &"' order by FLDNo"
	   'response.Write sql
	   Set Rs = syscon.Execute(Sql)
	   
	ENd if
	%>

    <BODY CLASS="BODY_01" scroll=no >

    <!-- HEAD 부분 시작 -->
    <TABLE CLASS=Sch_Header>
      <TR><TD><table cellpadding=2 cellspacing=0 border=0 width="99%">
          <tr><td><B>Table Name</b> : <%=tableid%>  /  <B>Description</B>: <%=tablenm%></TD>
              <TD align=right id=d2>
              <input type="button" value="새로고침" onclick="JavaScript:window.location.reload();" class=btn>
              <INPUT TYPE=button CLASS=btn value="필드추가" Onclick="javascript:centerWindow('sys_table_Fld_Wrt.asp?tableid=<%=tableid%>&tablenm=<%=tablenm%>','sys_table_fld_wrt_win','350','370','no');"> 
			  <% IF tableid <> "" and tabledb <> "" THEN %>
			  <INPUT TYPE=button CLASS=btn value="Rebuild" Onclick="javascript:URL_move('sys_table_cont.asp?tabledb=<%=tabledb%>&tableid=<%=tableid%>&gubun=rebuild_field&tablenm=<%=tablenm%>');">
			  <% END IF %></TD>
           </tr>
           </table>
      </TD></TR>
    </TABLE>
    <!-- HEAD 부분 끝 -->

    <!-- SHEET CONTENT 부분 시작 -->
    <DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="530"  style="TABLE-LAYOUT: fixed;" >
        <TR><td ALIGN="CENTER" CLASS="TBL_RW_00" WIDTH="05%">SEQ</td>
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="25%">필드명</td>
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="20%">필드타입</a></td>
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">NULL</td> 	
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="10%">출력</td> 
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="25%">Description</td> 
			<td ALIGN="CENTER" CLASS="TBL_RW_01" WIDTH="05%">X</td> 
		</TR>
		</table>
		
        <DIV CLASS="scr_01" id="scr_01" style="height:260px;">
        <TABLE BORDER="0" id='datalist' CELLPADDING="0" CELLSPACING="0" WIDTH="530" style="TABLE-LAYOUT: fixed;">		
        <%
		IF request("tableid") <> "" then

		nVcnt=1
		While Not Rs.EOF

		FLD_idx			=  Rs("FLD_idx") 
		FLDDB			=  Rs("FLDDB") 
		FLDtable		=  Rs("FLDtable") 
		FLDName			=  Rs("FLDName") 
		FLDNO			=  Rs("FLDNO") 
		FLDTYPE			=  Rs("FLDTYPE") 
		FLDLENG			=  Rs("FLDLENG") 
		FLDDisplay		=  Rs("FLDDisplay") 

		IF FLDLENG=2147483647 then
		FLDLENG=16
		END IF

		FLDDESC			=  Rs("FLDDESC") 
		FLDNULL			=  Rs("FLDNULL") 
		%>
		<tr style="cursor:hand;">
		    <td ALIGN="CENTER" WIDTH="05%" CLASS="TBL_DRW_01" style="font-family:Arial;font-size:12px;"><%=(Nvcnt)%></td>
		    <td ALIGN="LEFT" WIDTH="25%" CLASS="TBL_DRW_03" Onclick="centerModal('sys_table_Fld_Wrt.asp?FLDtable=<%=FLDtable%>&FLDdb=<%=FLDdb%>&FLDName=<%=FLDName%>&FLD_idx=<%=FLD_idx%>&slt=조회','350','370');" style="color:Red;font-family:Arial;font-size:12px;"><%=UCASE(FLDName)%></td>
		    <td ALIGN="LEFT" WIDTH="20%" CLASS="TBL_DRW_01" style="font-family:Arial;font-size:12px;"><%=FLDTYPE%>(<%=FLDLENG%>)</td>			
		    <td ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01"><Input type=checkbox name=FLDNULL <% IF FLDNull="1" then %>checked<% end if%> value=<%=FLDNull%>></td> 			
		    <td ALIGN="CENTER" WIDTH="10%" CLASS="TBL_DRW_01">
		   
		   <%
			SELECT CASE trim(FLDDisplay)
			CASE "1"
		   %>
		   <INPUT TYPE=checkbox name=FLDDisplay value="1" checked onclick="javascript:valueWindow('sys_table_Fld_change.asp?FLDDisplay=0&FLd_idx=<%=FLd_idx%>');">
		   <%
			CASE "0"
		   %>
		   <INPUT TYPE=checkbox name=FLDDisplay value="0" onclick="javascript:valueWindow('sys_table_Fld_change.asp?FLDDisplay=1&FLd_idx=<%=FLd_idx%>');"> 
		   <%
			end select
		   %></td>
		   <td ALIGN="LEFT" WIDTH="25%" CLASS="TBL_DRW_01"><%=FLDDESC%></td>
		   <td ALIGN="center" WIDTH="05%" CLASS="TBL_DRW_01"><a href="sys_Table_cont.asp?tablenm=<%=tablenm%>&tableid=<%=tableid%>&tabledb=<%=tabledb%>&FLd_idx=<%=FLd_idx%>&gubun=delete"><img src="<%=img_path%>msg_x.gif" border=0></a></td>
	   </tr>
	   <%
	   Nvcnt=Nvcnt+1
	   Rs.MoveNext
	   Wend

	   'Rs.Close
	   'Set Rs = Nothing
	   ELSE

		   Response.write "<Tr><td ALIGN=CENTER COLSPAN=6 CLASS=TBL_DRW_01 style='height:50;'>테이블을 선택해주세요</td></tr>"


	   END IF
	   %>        
	</table>  
    </DIV> 
    </DIV>    
    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
