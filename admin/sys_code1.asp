<!--#include Virtual = "/INC/config.asp"  -->
<%

Class tscript

public function sys_code(GUBUN_CODE) 


        '******************* 코드 조회 부분 시작 ************************************
        
	    SQL = "select GUBUN_CODE,GUBUN_TITLE,GUBUN_SEQ,GUBUN_NAME from GUBUN1 where GUBUN_CODE='"& GUBUN_CODE &"' order by GUBUN_SEQ "
		SET RS = CONN.EXECUTE(SQL)

        HTML = "<DIV class='layer_01' style='height:500px;'> "&vblf

        HTML = HTML &"<TABLE BORDER='0' CELLPADDING='0' CELLSPACING='0' WIDTH='450'  style='TABLE-LAYOUT: fixed;' > "&vblf
        HTML = HTML &"<TR> "&vblf
		HTML = HTML &"<TD ALIGN='CENTER' CLASS='TBL_RW_01' WIDTH='15%'>순번</TD> "&vblf
		HTML = HTML &"<TD ALIGN='CENTER' CLASS='TBL_RW_01' WIDTH='20%'>부모코드</TD> "&vblf
		HTML = HTML &"<TD ALIGN='CENTER' CLASS='TBL_RW_01' WIDTH='15%'>SEQ</TD> "&vblf
		HTML = HTML &"<TD ALIGN='CENTER' CLASS='TBL_RW_01' WIDTH='50%'>NAME</TD> "&vblf
        HTML = HTML &"</TR> "&vblf
        HTML = HTML &"</TABLE> "&vblf
       
        HTML = HTML &"<DIV CLASS='scr_01' style='height:500px;'> "&vblf
        HTML = HTML &"<TABLE id='datalist' BORDER='0' CELLPADDING='0' CELLSPACING='0' WIDTH='450' BGCOLOR=666666 style='TABLE-LAYOUT: fixed;'> "&vblf
        
        IF Not Rs.EOF THEN
            nVcnt = 1
            Do until Rs.EOF
            
                IF RS("GUBUN_SEQ")   = ""	THEN GUBUN_SEQ   = ""	ELSE GUBUN_SEQ   = RS("GUBUN_SEQ")		END IF
                IF RS("GUBUN_NAME")  = ""	THEN GUBUN_NAME  = ""	ELSE GUBUN_NAME  = RS("GUBUN_NAME")	END IF

                HTML = HTML &"<TR onclick=""select_sub('"& GUBUN_CODE &"','"& Server.HTMLencode(""& GUBUN_NAME &"") &"','"& GUBUN_SEQ &"');"" style='cursor:hand;'> "&vblf 
                HTML = HTML &"<TD ALIGN='CENTER' WIDTH='15%' CLASS='TBL_DRW_01'>"& nVcnt &"</TD> "&vblf    	
                HTML = HTML &"<TD ALIGN='CENTER' WIDTH='20%' CLASS='TBL_DRW_03'>"& GUBUN_CODE &"</TD> "&vblf 
                HTML = HTML &"<TD ALIGN='CENTER' WIDTH='15%' CLASS='TBL_DRW_01'>"& GUBUN_SEQ &"</TD> "&vblf 	    	   
                HTML = HTML &"<TD ALIGN='LEFT'   WIDTH='50%' CLASS='TBL_DRW_01'>"& GUBUN_NAME &"</TD> "&vblf 	    	   
                HTML = HTML &"</TR> "&vblf    

            nVcnt = nVcnt + 1
            Rs.MoveNext
            Loop
			HTML = HTML &"</TABLE>"&vblf
            HTML = HTML &"</DIV>"&vblf
            HTML = HTML &"</DIV>"&vblf 
           
		END IF

        '******************* 코드 조회 부분 끝 ************************************
        
		sys_code = HTML
		
        Rs.Close : Set Rs = Nothing

   exit function
end function

end class

set public_description = new tscript

RSDispatch

%>
<!--#include Virtual = "/INC/_ScriptLibrary/rs.asp"  -->

