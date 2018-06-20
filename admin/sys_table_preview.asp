    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램 : 테이블/필드 관리                                 			        --> 
    <!-- 작 성 자 : 문성원(050407)                                                    --> 
    <!-- 작성일자 : 2006년 6월 08일                                                   --> 
    <!-- 내    용 : 테이블/필드 관리                                                   --> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <BODY CLASS="BODY_01" >
    <%
	tableid = trim(request("tableid"))
    tablenm	= request("tablenm")
    tabledb	= trim(request("tabledb"))

	sql = "SELECT Name FROM "& tabledb &"..sysobjects WHERE type = 'U' and Name='"& tableid &"' "
	Set sysRs = syscon.Execute(sql)

	IF sysRs.EOF or sysRs.BOF then
	%>
		<DIV class="layer_01" align=left>
        <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%"  style="TABLE-LAYOUT: fixed;" >
		<Tr><Td align=center CLASS=TBL_DRW_01 style='height:50;'>테이블이 존재하지 않습니다.</td></tr>
		</table>
		</div>

	<%
	ELSE

		makesql= " select top 100 * from "& tabledb &".."& tableid 
		Set tblRs = syscon.Execute(makesql)
	
		%>
		
        <!-- SHEET CONTENT 부분 시작 -->
        <DIV class="layer_01" align=left>
            <TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%"  >
		        <%
			    ' 쿼리문에 여러 개의 레코드셋을 가져오는 경우도 있으므로
			    ' 다음과 같이 모든 레코드를 가져올 수 있도록 처리한다.
			    Do Until tblRs is Nothing
				intCount = tblRs.Fields.Count

				if ( intCount <> 0 ) then
					Response.Write "<TR>"
					response.write "<td ALIGN=CENTER CLASS=TBL_RW_00 width=30>no</td>"
					' 컬럼의 이름을 출력하는 부분
					For intLoop = 0 to intCount - 1
						Response.Write "<td ALIGN=CENTER CLASS=TBL_RW_01>"
						Response.Write tblRs.Fields( intLoop ).Name		' 컬럼의 이름을 가져온다.
						Response.Write "</TD>"
					Next

					Response.Write "</TR><Tr><TD colspan="& intCount &" height=1 bgcolor=666666></td></tr>"

					' 실제 레코드를 출력하는 부분
					Nvcnt =1
					Do Until ( tblRs.EOF = True )
						Response.Write "<TR>"
						response.write "<TD ALIGN='CENTER' CLASS='TBL_DRW_00' width=30>"& Nvcnt &"</td>"

						For intLoop = 0 to intCount - 1
							Response.Write "<TD ALIGN='CENTER' CLASS='TBL_DRW_01'>"
							Response.Write tblRs( intLoop )&"&nbsp; "
							Response.Write "</TD>"
						Next

						Response.Write "</TR>"
						tblRs.MoveNext
						Nvcnt = Nvcnt +1
					Loop
				end if

				Set tblRs = tblRs.NextRecordSet
			Loop
		%>
            </TABLE>
        </DIV>
    </DIV>
	
	<%
	END if

    %>
    <!--#include Virtual = "/INC/INC_FOOT.ASP" -->
