	<!--METADATA NAME="Microsoft ActiveX Data Objects 2.5 Library" TYPE="TypeLib" UUID="{00000205-0000-0010-8000-00AA006D2EA4}"-->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 도움말 정보관리                              			        --> 
    <!-- 내용: 도움말 정보관리                                                      --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 20일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
	<%
	PRG_NAME		= REQUEST("PRG_NAME")
	HELP_IDX		= REQUEST("HELP_IDX")
	HELP_PRG_ID		= REQUEST("HELP_PRG_ID") 
	HELP_TITLE		= TRIM(REQUEST("HELP_TITLE"))  
	SLT				= TRIM(REQUEST("SLT"))

	Function AddSlashes(strSrc)								
		AddSlashes = Replace(strSrc, "'", "\'")
	End Function

	Function StripSlashes(strSrc)							
		StripSlashes = Replace(strSrc, "\'", "'")
	End Function

	Function HtmlSpecialChars(strSrc)						
		Dim result
		result = Replace(strSrc, ">", "&gt;")
		result = Replace(result, "<", "&lt;")
		result = Replace(result, "&nbsp;", "&amp;nbsp;")
		result = Replace(result, """", "&quot;")
		HtmlSpecialChars = result
	End Function

    SELECT CASE SLT
	   CASE ""      SLT_GUBUN = "none" : S1 = "checked" : S2="disabled='true'" : S3="disabled='true'"
	   CASE "조회"  S1 = "disabled='true'" : S2="checked" : S3=""
       CASE "수정"  S2 = "checked" : S1="disabled='true'"
    END SELECT    
    
    SELECT CASE  SLT
    CASE  "조회" 

		SET RS = SYSCON.EXECUTE("SELECT * FROM SYS_HELP WHERE HELP_IDX= "& HELP_IDX )
		IF NOT ( RS.EOF OR RS.BOF ) THEN						  
		   HELP_IDX		=  RS("HELP_IDX") 
		   HELP_PRG_ID  =  RS("HELP_PRG_ID") 
		   HELP_TITLE	=  RS("HELP_TITLE") 
		   HELP_CONTENT =  RS("HELP_CONTENT") 
		   HELP_CONTENT = AddSlashes(HELP_CONTENT)
		END IF
        RS.CLOSE : SET RS=NOTHING

    CASE  "등록"

        For i = 1 To Request.Form("HELP_CONTENT").Count						' 본문 : 100K 씩 나눠서 온 걸 하나로 합친다
            msgbody = msgbody + Request.Form("HELP_CONTENT")(i)
        Next

        filedate = now()
        fileyy= year(filedate)
        filemm = right("0" & month(filedate),2)
        filedd = right("0" & day(filedate),2)
        filemi = right("0" & minute(filedate),2)
        filese = right("0" & second(filedate),2)
        filename = fileyy & filemm & filedd & filemi & filese

        uploadPath = BBS_UPLOAD_PATH & filename &"/" 						' 파일을 저장할 경로명
        uploadUrl  = BBS_UPLOAD_URL & filename &"/" 						' 첨부 파일을 읽어 들일 URL

        Set namoMime = Server.CreateObject("NamoMIME.MIMEObject")			' NamoMIME 유틸리티인 COM Class를 불러옴
        Set fso = Server.CreateObject("Scripting.FileSystemObject")			' 파일 시스템용 오브젝트를 생성

		'On Error Resume Next												' 에러 처리
        fso.CreateFolder uploadPath											' 폴더 생성
		Err.Clear

        namoMime.Decode msgbody, uploadPath									' 업로드 경로에 있는 파일들을 디코드함
        Set ts = fso.OpenTextFile(uploadPath & "noname.htm", 1)
        converted = ts.ReadAll
        ts.Close
        fso.DeleteFile uploadPath & "noname.htm"							' 디코드한 메세지와 파일을 지움.

        Set folder = fso.GetFolder(uploadPath)
        Set files = folder.Files
        For each f in files
            converted = Replace(converted, f.name, uploadUrl & f.name)
        Next
        converted = AddSlashes(Left(converted, Len(converted) - 1))			' EOF가 있는 HTML을 읽어 들임.
		converted = replace(converted,uploadPath&"\","")
		converted = replace(converted,"'","")
			  
		SQL = "INSERT INTO SYS_HELP (HELP_PRG_ID, HELP_TITLE, HELP_CONTENT, HELP_ID) " 
		SQL = SQL & " VALUES (" & HELP_PRG_ID & ",'" & HELP_TITLE & "','" & CONVERTED & "','"& cid &"')"                     
		SYSCON.EXECUTE(SQL)

        Alert_Message "처리되었습니다","2"

    CASE  "수정" 

		Set Rs = syscon.Execute("select * from SYS_HELP where help_idx= "& help_idx )

        IF Rs.EOF THEN
           Alert_Message "데이터가 없습니다\n\n확인바랍니다.","2"
        ELSE
		
			For i = 1 To Request.Form("HELP_CONTENT").Count						' 본문 : 100K 씩 나눠서 온 걸 하나로 합친다
				msgbody = msgbody + Request.Form("HELP_CONTENT")(i)
			Next

			filedate = now()
			fileyy= year(filedate)
			filemm = right("0" & month(filedate),2)
			filedd = right("0" & day(filedate),2)
			filemi = right("0" & minute(filedate),2)
			filese = right("0" & second(filedate),2)
			filename = fileyy & filemm & filedd & filemi & filese

			uploadPath = BBS_UPLOAD_PATH & filename &"/" 					' 파일을 저장할 경로명
			uploadUrl  = BBS_UPLOAD_URL & filename &"/" 	' 첨부 파일을 읽어 들일 URL

			Set namoMime = Server.CreateObject("NamoMIME.MIMEObject")			' NamoMIME 유틸리티인 COM Class를 불러옴
			Set fso = Server.CreateObject("Scripting.FileSystemObject")			' 파일 시스템용 오브젝트를 생성

			'On Error Resume Next												' 에러 처리
			fso.CreateFolder uploadPath											' 폴더 생성
			Err.Clear

			namoMime.Decode msgbody, uploadPath									' 업로드 경로에 있는 파일들을 디코드함
			Set ts = fso.OpenTextFile(uploadPath & "noname.htm", 1)
			converted = ts.ReadAll
			ts.Close
			fso.DeleteFile uploadPath & "noname.htm"							' 디코드한 메세지와 파일을 지움.

			Set folder = fso.GetFolder(uploadPath)
			Set files = folder.Files
			For each f in files
				converted = Replace(converted, f.name, uploadUrl & f.name)
			Next
			converted = AddSlashes(Left(converted, Len(converted) - 1))			' EOF가 있는 HTML을 읽어 들임.
			converted = replace(converted, "'", "''")
			converted = replace(converted,uploadPath&"\","")
			
			SQL = "UPDATE SYS_HELP SET HELP_PRG_ID="& HELP_PRG_ID &",HELP_TITLE='"& HELP_TITLE &"',HELP_CONTENT='"& converted &"',HELP_ID = '"& cid &"'"
			SQL = SQL & " WHERE HELP_IDX ="& HELP_IDX
			SYSCON.EXECUTE(SQL)

			RS.CLOSE : SET RS=NOTHING

			Alert_Message "수정 처리되었습니다","2"

		END IF

    CASE  "삭제" 

		SET RS = SYSCON.EXECUTE("SELECT * FROM SYS_HELP WHERE HELP_IDX= "& HELP_IDX )

        IF  RS.EOF THEN
            Alert_Message "데이터가 없습니다\n\n확인바랍니다.","2"
        ELSE
            CONN.EXECUTE("DELETE FROM SYS_HELP WHERE HELP_IDX = " & HELP_IDX)
			RS.CLOSE : SET RS=NOTHING
            Alert_Message "처리되었습니다","2"
        END IF

    END SELECT      

	%>
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
	<Script Language="VBScript">
	Function Wec_OnInitCompleted 
		dim object 
		set object = document.form	
		<% IF SLT = "조회" THEN %>
		object.wec.Value = form.contents.value
		<% END IF %>
	End Function 
	</SCRIPT> 
	<BODY CLASS="POPUP_01" SCROLL="no">

    <% CALL Popup_generate("온라인 HELP 등록/수정/삭제 ","정확하게 입력하셔야만 합니다")%>

	<FIELDSET style="width:98%;">
    <LEGEND><FORM  NAME="form" method="post">
			<INPUT type="hidden" name="contents" value="<%= HtmlSpecialChars(help_CONTENT) %>">
			<Input type="hidden" name="help_PRG_id" value="<%=help_PRG_id%>" >
			<Input type="hidden" name="help_idx" value="<%=help_idx%>" >
            <input type="radio" <%=(S1)%> name="slt" value="등록" >등록                
            <input type="radio" <%=(S2)%> name="slt" value="수정" >수정
            <input type="radio" <%=(S3)%> name="slt" value="삭제" >삭제
	</LEGEND>
	
	<table border="0" cellspacing="0" width="100%"   >
    <tr><td ID="TBL_Title">프로그램명</td> 
        <td ID="TBL_DATA1"><input type="text" name="prg_name" value="<%=prg_name%>" class="INPUT_readonly" style="width:100%;" readonly></td> 
    </tr> 
    <tr><td ID="TBL_TITLE">제목</td> 
        <td ID="TBL_DATA1"><input type="text" name="help_title" value="<%=help_title%>" class="INPUT" style="width:100%;"></td> 
    </tr> 
    <tr><td colspan="2"><script src="/inc/js/NamoWec.js"></script></td></tr>
	</table> 
	
	</FIELDSET>

    <table border="0" cellpadding="1" cellspacing="1" width="100%">
       <Tr><Td align="center" height=5></td></tr>
       <tr><td style="padding-right:10px;" align=right><input type="button"  value=" 확 인 " class=btn onclick="javascript:Onsave();"> <input type="button" value=" 닫 기 " class=btn onclick="self.close();"></td></form>
       </tr>                
    </table>
       
    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->

    <SCRIPT language="JavaScript">
    <!--
    self.moveTo((screen.availWidth-780)/2,(screen.availHeight-690)/2)
    self.resizeTo(770,710)

	var action;
	function DoSubmit()
	{
        var form = document.form;
        form.action = action;
        form.submit();
	}

	function DivideString(strSrc)
	{
		var strTag = strSrc;
		var form = document.form;
		var tags;

			while(strTag.length > 0)
			{
				tags = document.createElement("TEXTAREA");
				tags.name = "HELP_CONTENT";
				tags.value = strTag.substr(0, 102400);
				form.appendChild(tags);
				strTag = strTag.substr(102400);
			}

		tags = document.createElement("TEXTAREA");
		tags.name = "HELP_CONTENT";
		tags.value = strTag;
		form.appendChild(tags);
	}

    // 실제 입력폼 검증 부분 SCRIPT 처리
    function Onsave(){

        var form = document.form;
        var wec  = document.wec;

        if (form.help_title.value=="") {
        alert("제목을 선택해주십시오.")
        form.help_title.focus();
        return;    
        }

        if (confirm("등록을 완료 하시겠습니까?")) {
    
        DivideString(wec.MIMEValue);
	    action = "sys_HELP_wrt.asp";
        DoSubmit();
	    return;
	    }
    else {
		    return;
       }   
    }
    //-->
    </SCRIPT>    

