    <!--#include Virtual = "/INC1/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : ERP 용어 관리                              			--> 
    <!-- 작성자: 문성원(050407)                                                   --> 
    <!-- 작성일자 : 2006년 11월 16일                                        --> 
    <!-- 내용: 언어별 ERP 용어 관리												--> 
    <!-- ========================================================================== -->		    
    <!--#include Virtual = "/INC1/INC_BODY.ASP" -->    
    <%
    
    server.Execute("sys_word_create.asp?characterset=KOR")
    server.Execute("sys_word_create.asp?characterset=CHN")
    server.Execute("sys_word_create.asp?characterset=ENG")
	
	Alert_message "생성되었습니다","2"
    %>
    <!--#include Virtual = "/INC1/INC_FOOT.ASP" -->    
