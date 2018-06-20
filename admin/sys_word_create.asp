    <!--#include Virtual = "/INC1/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : ERP 용어 관리                              			--> 
    <!-- 작성자: 문성원(050407)                                                   --> 
    <!-- 작성일자 : 2006년 11월 16일                                        --> 
    <!-- 내용: 언어별 ERP 용어 관리												--> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC1/INC_BODY.ASP" -->    
    <% 
    
    edit_Host = Request.ServerVariables("REMOTE_ADDR")
    wdate=now()
    characterset = request("characterset")

    select case characterset
    case "KOR"
	    confname = gw_rpath & "inc1\language_k.asp"
	    conftext = "한국어"
    case "ENG"
	    confname = gw_rpath & "inc1\language_e.asp"
	    conftext = "영어"
    case "CHN"
	    confname = gw_rpath & "inc1\language_c.asp"
	    conftext = "중국어"
	end select

    Set FS = Server.CreateObject("Scripting.FileSystemObject")
	Set aw = Fs.OpenTextFile(confname,2,-1)  
				
        sql = " declare @charset char(3)"
        sql = sql &" set @charset = '"& characterset &"' "
        sql = sql &" select w_code,case when @charset = 'KOR' then w_kword "
        sql = sql &"                    when @charset = 'CHN' then "
        sql = sql &"                            case when w_cword is null or w_cword ='' then w_kword else w_cword end "
        sql = sql &"                    when @charset = 'ENG' then "
        sql = sql &"                            case when w_eword is null or w_eword ='' then w_kword else w_eword end "
        sql = sql &"               end as w_word ,w_content from sys_word a "
        sql = sql &" left join dongaerp.dbo.gubun1 on gubun_code=left(w_gubun,2) and gubun_seq=right(w_gubun,2) order by w_code asc"
        set rs = syscon.execute(sql)
    	
	    aw.writeline ("<%")
	    aw.writeline ("'이 파일은 "& conftext &" Library Language 파일입니다.")
	    aw.writeline ("'이 파일은 임의대로 수정을 할수 없으며, 반드시 업무용어관리에서만 생성하실 수 있습니다. ")
	    aw.writeline ("'이 파일은 "& wdate &"에 "& edit_Host &"에 의해 새롭게 생성되었습니다.")
    	
        while not rs.eof
        
            w_code      =   Trim(RS("w_code"))
            w_word      =   Trim(RS("w_word"))
            w_content   =   replace(Trim(server.HTMLEncode(RS("w_content")))," ","")

	        aw.writeline (w_code &" = "& chr(34) & w_word & chr(34)) & vclf

	        Rs.MoveNext
	        Vcnt=Vcnt+1
	        
	    Wend

	Rs.Close : Set Rs = Nothing
	
	aw.writeline(chr(37)&">")	
	
    %>
    <!--#include Virtual = "/INC1/INC_FOOT.ASP" -->    
