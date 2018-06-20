    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : DB Query analyzer                              					--> 
    <!-- 내용: DB Query analyzer												    --> 
    <!-- 작 성 자 : 문성원(050407)                                                    --> 
    <!-- 작성일자 : 2006년 6월 19일                                                   --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
	<%
	strDataBase = Request.Form("DataBase")
	strQueryStatement = Request.Form("QueryStatement")
	%>

    <BODY CLASS="BODY_01" >

	<Table border=0 cellpadding=0 cellspacing=1 width="100%" align=center>
	<TR><Td align=center>
	<% 
		if ( strDataBase <> "" ) then

			' Connection String 을 작성한다.


			' 쿼리문이 Insert, Update, Delete의 경우에는 결과물이 레코드셋이 아니므로 따로 처리한다.
			if ( Left( LCase( strQueryStatement ), 6 ) = "insert" ) or ( Left( LCase( strQueryStatement ), 6 ) = "update" ) or ( Left( LCase( strQueryStatement ), 6 ) = "delete" ) then
				On Error Resume Next

				Conn.Execute strQueryStatement

				if ( Err.Number <> 0 ) then
					Response.Write "오류 발생 : " & Err.Description
					Response.End
				else
					Response.Write "쿼리문이 제대로 수행 되었습니다."
					Response.End
				end if
			else
				Set objRs = Conn.Execute( strQueryStatement )

				Dim elements
				Dim elem
				Dim intCount
				Dim intLoop
			%>

			<table width="100%"  border="0" cellpadding="0" cellspacing="1" align=center BGCOLOR=666666>


			<%
				' 쿼리문에 여러 개의 레코드셋을 가져오는 경우도 있으므로
				' 다음과 같이 모든 레코드를 가져올 수 있도록 처리한다.
				Do Until objRs is Nothing
					intCount = objRs.Fields.Count

					if ( intCount <> 0 ) then
						Response.Write "<TR height=22>"

						' 컬럼의 이름을 출력하는 부분
						For intLoop = 0 to intCount - 1
							Response.Write "<TD bgcolor=#629AD9 CLASS=paging_white_normal align=center style='padding-left:3px;'>"
							Response.Write objRs.Fields( intLoop ).Name		' 컬럼의 이름을 가져온다.
							Response.Write "</TD>"
						Next

						Response.Write "</TR>"

						' 실제 레코드를 출력하는 부분
						Do Until ( objRs.EOF = True )
							Response.Write "<TR height=22>"

							For intLoop = 0 to intCount - 1
								Response.Write "<TD bgcolor=#ffffff style='padding-left:3px;'>"
								Response.Write objRs( intLoop )&" "
								Response.Write "</TD>"
							Next

							Response.Write "</TR>"
							objRs.MoveNext
						Loop
					end if

					Set objRs = objRs.NextRecordSet
				Loop
	%>

			</table>

	<%
			end if

		else
	%>
	<table border=0 cellpadding=0 cellspacing=0 width="100%" height="100%">
	<Tr><Td align=center valign=middle>QUERY 검색결과가 보여집니다.</td></tR>
	</table>
	<%
		end if
	%>
	</td></tr>
	</table>


    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->
