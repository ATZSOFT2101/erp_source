    <!--#include Virtual = "/INC/INC_TOP.ASP"  -->
    <!-- ========================================================================== -->
    <!-- 프로그램명 : 프로그램 구성 메뉴 관리                              			--> 
    <!-- 내용: Mlevel 구성 관리                                                     --> 
    <!-- 작 성 자 : 문성원(050407)                                                  --> 
    <!-- 작성일자 : 2006년 6월 13일                                                 --> 
    <!-- ========================================================================== -->		
    <!--#include Virtual = "/INC/INC_BODY.ASP" -->
    <%
    IF REQUEST("CHK") = "수정" THEN
       TEMP =  SPLIT(REQUEST("DEST"),",")
       CNT = 0
       WHILE  TEMP(CNT) <> "" 
       'RESPONSE.WRITE CNT + 1 & "___" & TEMP(CNT) & ","
       SYSCON.EXECUTE("UPDATE MLEVEL SET DIS_SEQ =" & CNT + 1 & " WHERE ID = " & TEMP(CNT))
       CNT = CNT + 1
       WEND
    END IF

    'RESPONSE.WRITE REQUEST("ID")
    SET RS  =  SYSCON.EXECUTE("EXPAND1 "& REQUEST("ID") &",'"& id &"','syerp' ")
    %> 
    <script type="text/javascript">
    <!--

    var optionsMessage = "Select a filter from the list to edit the options";
    var lightMessage = "The light filter doesn't work like other filters. To add light sources you call on different methods.<br><br>See the documentation at <a style='color: blue; text-decoration: underline' href='http://msdn.microsoft.com/workshop/author/filter/reference/filters/light.asp' target='_top'>Microsoft Internet Client SDK</a>";

    function addWebFilter(el) {
	    if (el.selectedIndex == -1) return;
	    var opt = document.createElement("OPTION");
	    opt.text = el.item(el.selectedIndex).text;
	    if (addedFilters.selectedIndex != -1) {
		    addedFilters.add(opt, addedFilters.selectedIndex + 1);
		    addedFilters.selectedIndex += 1;		    
	    }
	    else {
		    addedFilters.add(opt);
		    addedFilters.selectedIndex = addedFilters.length -1;
	    }
        el.remove(el.selectedIndex);
	    el.selectedIndex = el.length -1;    
    }

    function removeWebFilter(el) {
	    if (el.selectedIndex == -1) return;
	    var opt = document.createElement("OPTION");
	    opt.text = el.item(el.selectedIndex).text;
	    if (filterSelect.selectedIndex != -1) {
		    filterSelect.add(opt, filterSelect.selectedIndex + 1);
		    filterSelect.selectedIndex += 1;		    
	    }
	    else {
		    filterSelect.add(opt);
		    filterSelect.selectedIndex = filterSelect.length -1;
	    }
	    el.remove(el.selectedIndex);
	    el.selectedIndex = el.length -1;
    }


    function moveUp(el) {
	    var index = el.selectedIndex;

	    if ((index < 1) || (index == -1)) return;

	    var tmp = el.item(index - 1);

	    var opt = document.createElement("OPTION");
	    opt.text = tmp.text;
	    opt.value = tmp.value;
	    opt.attrib = tmp.attrib;
	    addedFilters.remove(index-1);
	    addedFilters.add(opt, index);
	    addedFilters.selectedIndex = index - 1;
    	
    }

    function moveDown(el) {
	    var index = el.selectedIndex;

	    if ((index > el.length - 2) || (index == -1)) return;

	    var tmp = el.item(index + 1);

	    var opt = document.createElement("OPTION");
	    opt.text = tmp.text;
	    opt.value = tmp.value;
	    opt.attrib = tmp.attrib;
	    addedFilters.remove(index + 1);
	    addedFilters.add(opt, index);
	    addedFilters.selectedIndex = index + 1;

    }

    function makeOrder(){
        blipOrder.value=""
             for (var i=0; i<addedFilters.length;i++){
                 blipOrder.value+=(addedFilters[i].value + ",")
             }            
       //      alert(blipOrder.value);
             document.location.href = "sys_menu_Seq.asp?id=<%=Request("id")%>&chk=수정&dest=" + blipOrder.value
    }
    //-->
    </script>

    <BODY CLASS="POPUP_01" SCROLL="no">
    <% CALL Popup_generate("메뉴순서구성","상하 순서로 정렬됩니다")%>

    <FIELDSET>
    <LEGEND></LEGEND>

    <Div class="div_separate_10"></div>

	<TAble border=0 width="100%">
	<Tr><TD width=150><select name="addedFilters" size="10" ondblclick="removeWebFilter(this)" style="height: 200; width:150px;">
	<%             
	while not Rs.EOF 
          if Rs("level") = 2 then                                         
		  %>                    
          <option value="<%=trim(Rs("item"))%>"><%=Rs("item1")%></option>
		  <%                
		  end if
	Rs.MoveNext    
    wend
	%>     
	</select></td>
	<Td><Input type=button  onclick="moveUp(addedFilters)"   value="Up Level" style="height:20px; width:130px;" class=btn><Div class="div_separate_5"></div>

	    <Input type=button  onclick="moveDown(addedFilters)" value="Down Level" style="height:20px;width:130px;" class=btn><Div class="div_separate_10"></div>
	
	    <input type="button" value="Save" onclick="javascript:makeOrder();" style="height:20px; width:130px;"class=btn>
	<INPUT TYPE="hidden" NAME="blipOrder" VALUE=""></td>
	</tr>
	</table>

    </fieldset>


    
    <!--#include Virtual = "/INC/INC_FOOT.ASP"              -->
