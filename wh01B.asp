<!-- #include file="wFundProc.asp" -->
<%
sdate = trim(request("A"))
area = trim(request("B"))
kind = trim(request("C"))
if Request("D") = "" then
	Page = 1
else
	Page = CLng(Request("D"))   ' CLng 不可省略
end if
if sdate = "" then sdate="NA"
if kind  = "" then kind ="NA"
if area = "" then  area ="NA"

dim areanamearray( 4)
areanamearray(0) = "最新"
areanamearray(1) = "美國地區"
areanamearray(2) = "歐洲地區"
areanamearray(3) = "遠東地區"
areanamearray(4) = "東南亞地區"

select case area
	case "0"
		areastr = "23456789"
		areaname = areanamearray(0)
	case "1"
		areastr = "29"
		areaname = areanamearray(1)
	case "2"
		areastr = "678"
		areaname = areanamearray(2)
	case "3"
		areastr = "34"
		areaname = areanamearray(3)
	case "4"
		areastr = "5"
		areaname = areanamearray(4)
	case else
		areastr = "29"
		areaname = areanamearray(1)
end select
  
xxxIDS = ""
xxxIDS = xxxIDS & "<select onchange=""selopn(this.options[this.selectedIndex].value )"" name=""IDS"" id=""IDS"" size=""1"">" & chr(13) & chr(10)
for i = 0 to 4
	if i = CInt(area) then
		xxxIDS = xxxIDS & "<option selected value=""/w/wh/wh01b_NA_" & i & ".djhtm"">" & areanamearray(i) & "</option>" & chr(13) & chr(10)
	else
		xxxIDS = xxxIDS & "<option value=""/w/wh/wh01b_NA_" & i & ".djhtm"">" & areanamearray(i) & "</option>" & chr(13) & chr(10)
	end if
next
xxxIDS = xxxIDS & "</select>" & chr(13) & chr(10)

Response.Write GetDocProlog("研究報告", "wh01B","NA", "NA")

Response.Write "<table width=" & g_TableWidth & ">" & chr(13) & chr(10)
Response.Write "<form method=POST name=wh01b_frm>" & chr(13) & chr(10)
Response.Write "<tr><td class=wfb0c colspan=4>" & xxxIDS & "研究報告</form</td></tr>" & chr(13) & chr(10)
Response.Write "</form>" & chr(13) & chr(10)
Response.Write "<tr><td class=wfb3l>日期</td><td class=wfb3l>標題</td><td class=wfb3l>機構</td><td class=wfb3l>作者</td></tr>" & chr(13) & chr(10)
Response.write sayReportList(area, areastr, areaname, page)
Response.Write "</table>" & chr(13) & chr(10)

Response.Write GetDocEplog("R")
  
function sayReportList(areacode, areastr, areaname, page)
	dim xxx 
	xxx = ""
	if sdate = "NA" then
		sql = "exec spj_mda00911 null,'" & Areastr & "',null,null,100"
	else
		sql = "exec spj_mda00911 null,'" & areastr & "',null,'" & sdate & "',100"
	end if
	
	if OpenJust(conn, rs, sql) then
	  
		rs.PageSize = 20
		If Page > rs.PageCount Then Page = rs.PageCount
		rs.AbsolutePage = Page
	
		rno = 0
		for iPage = 1 to rs.PageSize		
			rno = rno + 1
			cs = "wfb1"
			if rno mod 2 = 0 then cs = "wfb2"
			xxx = xxx & "<tr><td class=" & cs & "l>" & formatYYMD(rs(0)) & "</td>" & chr(13) & chr(10)
			xxx = xxx & "<td class=" & cs & "l><a href=""/w/wh/wh01a_" & trim(rs(6)) & ".djhtm"">" & trim(rs(5)) & "</a></td>" & chr(13) & chr(10)
			xxx = xxx & "<td class=" & cs & "l>" & trim(rs(4)) & "</td>" & chr(13) & chr(10)
			xxx = xxx & "<td class=" & cs & "l>" & stdfmt(trim(rs(2)),0) & "</td></tr>" & chr(13) & chr(10)
			rs.MoveNext
			if rs.Eof then
				exit for
			end if		
		next
		
		xxx = xxx & "<tr><td class=wfb2l colspan=4>" & chr(13) & chr(10)
		if rs.PageCount > 1  then
			xxx = xxx & "<div align=center>" 
			xxx = xxx & "<form name=""Formpage"" onSubmit=""return go()"" style=""font-size:9pt"">" & Chr(13) & Chr(10)
			If Page <> 1 Then
				xxx = xxx & "<A HREF=""/w/wh/wh01B_" & sdate & "_" & area & "_" & kind & "_1.djhtm"">第一頁</A>　" & Chr(13) & Chr(10)
				xxx = xxx & "<A HREF=""/w/wh/wh01B_" & sdate & "_" & area & "_" & kind & "_" & Page-1 & ".djhtm"">上一頁</A>　" & Chr(13) & Chr(10)
			End If
			If Page <> rs.PageCount Then
				xxx = xxx & "<A HREF=""/w/wh/wh01B_" & sdate & "_" & area & "_" & kind & "_" & page+1 & ".djhtm"">下一頁</A>　" & Chr(13) & Chr(10)
				xxx = xxx & "<A HREF=""/w/wh/wh01B_" & sdate & "_" & area & "_" & kind & "_" & rs.PageCount & ".djhtm"">最後一頁</A>　" & Chr(13) & Chr(10)
			End If
			xxx = xxx & "輸入頁次:<select name=""page1"" onchange=""goPage()"">"
			for i = 1 to rs.PageCount 
				if page = i then
					setidx = " selected "
				else
					setidx = ""
				end if
				xxx = xxx & "<option value =""/w/wh/wh01B_" & sdate & "_" & area & "_" & kind & "_" & i & ".djhtm"" " & setidx & ">" & i & "</option>" & vbcrlf
			next
			xxx = xxx & "</select>"
			xxx = xxx & "頁次:<FONT COLOR=""Red"">/" & rs.PageCount &"</FONT>" & Chr(13) & Chr(10)
			xxx = xxx & "</FORM>"
			xxx = xxx & "</div>" & Chr(13) & Chr(10)
			xxx = xxx & "<script language=""JavaScript""><!--" & chr(13) & chr(10)
			xxx = xxx & "function goPage(){" & vbcrlf
			xxx = xxx & "document.location = document.Formpage.page1.options[document.Formpage.page1.selectedIndex].value;" & vbcflf
			xxx = xxx & "}" & vbcrlf
	
			xxx = xxx & "function go() {" & chr(13) & chr(10)
			xxx = xxx & "if(document.F.F.value <= """ & rs.PageCount & """ && document.F.F.value > ""0"" ) {" & chr(13) & chr(10)
			xxx = xxx & "  pno = parseInt(document.F.F.value);" & chr(13) & chr(10)
			xxx = xxx & "  if(pno < 1) pno = 1;" & chr(13) & chr(10)
			xxx = xxx & "  if(pno > " & rs.PageCount & ") pno = " & rs.PageCount & ";}" & chr(13) & chr(10)
			xxx = xxx & "  else pno = 1;" & chr(13) & chr(10)
			xxx = xxx & "self.location = '/w/wh/wh01B_NA_" & areacode & "_NA_' + pno + '.djhtm';" & chr(13) & chr(10)
			xxx = xxx & "return false;}" & chr(13) & chr(10) 
			xxx = xxx & "// --></script>" & chr(13) & chr(10)
		end if  
		xxx = xxx & "<div align=center>" 
		xxx = xxx & "<form name=sch onSubmit=""return go1()""  font style=""font-size:9pt"">" & Chr(13) & Chr(10)
		xxx = xxx & "以西元日期(yyyy/mm/dd)查詢<input type=text name=B size=8 value=" & datechg(sdate) & ">" & chr(13)  & chr(10)
		xxx = xxx & "<input type=button value=GO name=b1 onclick=""return go1()"">" & chr(13) & chr(10)
		xxx = xxx & "</FORM></div>" & Chr(13) & Chr(10)
		xxx = xxx & "<script language=""Javascript"" src=""/w/js/jschkd.djjs""></script>" & chr(13) & chr(10)
		xxx = xxx & "<script language=""JavaScript""><!--" & chr(13) & chr(10)
		xxx = xxx & "	function go1() {" & chr(13) & chr(10)
		xxx = xxx & "   var B = document.sch.B.value;" & chr(13) & chr(10)
		xxx = xxx & "	if (B == '') {" & chr(13) & chr(10)
		xxx = xxx & "		B='NA';" & chr(13) & chr(10)
		xxx = xxx & "	}" & chr(13) & chr(10)
		xxx = xxx & "	else if ((B = chkYDate(B,1)) == false){" & chr(13) & chr(10)
		xxx = xxx & "		return false;" & chr(13) & chr(10)
		xxx = xxx & "	}" & chr(13) & chr(10)
		xxx = xxx & "	self.location = '/w/wh/wh01B_' + B + '_" & areacode & "_NA_1.djhtm';"& chr(13) & chr(10)
		xxx = xxx & "	return false;} " & chr(13) & chr(10)
		xxx = xxx & "// --></script>" & chr(13) & chr(10)
		
		xxx = xxx & "</td></tr>" & chr(13) & chr(10)	    
		
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if
	sayReportList = xxx
end function
%>

