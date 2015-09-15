<!-- #include file="wtFundProc.asp" -->
<%
tid = ucase(trim(request("A")))
ord = ucase(trim(request("B")))
if tid="" or tid="NA" then tid = deftTID
if ord="" or ord="NA" then ord = "801"

if g_Customer_ShowFundOrderBtn_Flag  then
	colNo = 8 + g_OrderBtnCount
else
	colNo = 8
end if

Response.Write GetDocProlog("報酬依類型AA", "wq02","NA", cid,"NA")

sql = "exec TSP_GET_FUND_BY_TID '" & fundaspid & "','" & tid & "','" & ord & "'"
'Response.Write sql & "<BR>"

if OpenFundDJ(conn, rs, sql) then
	call FormHead(rs("yb800010"))
	rno = 0
	while not rs.EOF 
		rno = rno + 1
		cs = "wfb5"
		if rno mod 2 = 0 then cs = "wfb2" 
		
		Response.Write "<tr>" & chr(13) & chr(10)
		
		if g_Customer_ShowFundOrderBtn_Flag and g_OrderBtnPosition = 1 and fundaspid <> "" then
			if ucase(trim(rs("approved"))) = "Y" then
				Response.Write "<td class=" & cs & "c>" & MakeButton(rno,trim(rs("bankfundid"))) & "</td>" & chr(13) & chr(10)
			else
				Response.Write "<td class=" & cs & "c></td>" & chr(13) & chr(10)
			end if
		end if
		
		if g_Customer_ShowFundCode_Flag then
			if fundaspid <> "" then
				Response.Write "<td class=" & cs & "l><a href=""/w/wr/wr01_" & trim(rs(0)) & ".djhtm"">" & rs("bankfundid") & trim(rs(1)) & "</a></td>" & chr(13) & chr(10)
			else
				Response.Write "<td class=" & cs & "l><a href=""/w/wr/wr01_" & trim(rs(0)) & ".djhtm"">" & trim(rs(1)) & "</a></td>" & chr(13) & chr(10)
			end if
		else
			Response.Write "<td class=" & cs & "l><a href=""/w/wr/wr01_" & trim(rs(0)) & ".djhtm"">" & trim(rs(1)) & "</a></td>" & chr(13) & chr(10)
		end if    
	
		Response.Write "<td class=" & cs & "l><a href=/w/wq/wq01_" & trim(rs(2)) & ".djhtm>" & trim(rs(3)) & "</a></td>" & chr(13) & chr(10)
		Response.Write "<td class=" & cs & "r nowrap>" & stdfmt(rs("yb800010"),8) & "</td>" & chr(13) & chr(10)
		Response.Write "<td class=" & cs & "r nowrap>" & stdfmt(rs("yb800300"),2) & "</td>" & chr(13) & chr(10)
		Response.Write "<td class=" & cs & "r nowrap>" & stdfmt(rs("yb800801"),2) & "</td>" & chr(13) & chr(10)
		Response.Write "<td class=" & cs & "r nowrap>" & stdfmt(rs("yb800803"),2) & "</td>" & chr(13) & chr(10)
		Response.Write "<td class=" & cs & "r nowrap>" & stdfmt(rs("yb800806"),2) & "</td>" & chr(13) & chr(10)
		Response.Write "<td class=" & cs & "r nowrap>" & stdfmt(rs("yb800901"),2) & "</td>" & chr(13) & chr(10)
		'Response.Write "<td class=" & cs & "r nowrap>" & stdfmt(rs("yb800902"),2) & "</td>" & chr(13) & chr(10)
		'Response.Write "<td class=" & cs & "r nowrap>" & stdfmt(rs("yb800903"),2) & "</td>" & chr(13) & chr(10)
		'Response.Write "<td class=" & cs & "r nowrap>" & stdfmt(rs("yb800905"),2) & "</td>" & chr(13) & chr(10)
		if g_Customer_ShowFundOrderBtn_Flag and g_OrderBtnPosition = 2 and fundaspid <> "" then
			if ucase(trim(rs("approved"))) = "Y" then
				Response.Write "<td class=" & cs & "c>" & MakeButton(rno,trim(rs("bankfundid"))) & "</td>" & chr(13) & chr(10)
			else
				Response.Write "<td class=" & cs & "c></td>" & chr(13) & chr(10)
			end if
		end if
		Response.Write "</tr>" & chr(13) & chr(10)
		rs.MoveNext
	wend
	Response.Write "<tr><td class=wfb4l colspan=" & colNo & ">" & FundMsgText(0) & "</td></tr>" & chr(13) & chr(10)
	Response.Write "</table>" & chr(13) & chr(10)
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
else
	call FormHead(null)
	Response.Write "<tr><td class=wfb2c colspan=" & colNo & ">查無資料</td></tr>" & chr(13) & chr(10)
	Response.Write "<tr><td class=wfb4l colspan=" & colNo & ">&nbsp;</td></tr>" & chr(13) & chr(10)
end if
Response.Write "</table>" & chr(13) & chr(10)
Response.Write "<script language=""JavaScript""><!--" & chr(13) & chr(10)
Response.Write "InitComboList(document.frmSel.selTID, '/w/wq/wq02_', '_" & ord & ".djhtm', '" & tid & "', tfund_type, '');" & chr(13) & chr(10)
Response.Write "// --></script>" & chr(13) & chr(10)
Response.Write GetDocEplog("Q")

sub FormHead(dated)
	Response.Write "<table width=" & g_TableWidth & ">" & chr(13) & chr(10)
	Response.Write "<FORM method=POST name=frmSel align=center onsubmit=""return false;"">" & chr(13) & chr(10)
	Response.Write "<tr><td class=wfb0c colspan=" & colNo & ">" & chr(13) & chr(10)
	Response.Write "<SELECT name=selTID2 onchange=selopn(this.options[this.selectedIndex].value)>" & chr(13) & chr(10)
	Response.Write "<OPTION  value=""/w/wq/wq01.djhtm"">基金公司</OPTION>" & chr(13) & chr(10)
	Response.Write "<OPTION selected value=""/w/wq/wq02.djhtm"">基金類型</OPTION>" & chr(13) & chr(10)
	Response.Write "</SELECT>" & chr(13) & chr(10)
	Response.Write "<SELECT name=selTID onchange=selopn(this.options[this.selectedIndex].value)>" & chr(13) & chr(10)
	for selcnt=1 to 5
		Response.Write "<OPTION>ＸＸＸＸＸＸＸＸＸ</OPTION>" & chr(13) & chr(10)
	next
	Response.Write "</SELECT>之績效排行" & chr(13) & chr(10)
	Response.Write "<div align=left style=""font-size:9pt"">"
	Response.Write "<input type=radio name=radio1 checked onclick=""javascript:window.open('/w/wq/wq02_" & tid & "_" & ord & ".djhtm','_self');"">報酬顯示" & chr(13) & chr(10)
	Response.Write "<input type=radio name=radio1         onclick=""javascript:window.open('/w/wq/wq04_" & tid & "_" & ord & ".djhtm','_self');"">風險顯示" & chr(13) & chr(10)
	Response.Write "</div>"
	Response.Write "</td></tr>" & chr(13) & chr(10)
	Response.Write "</form>" & chr(13) & chr(10)
	Response.Write "<tr>" & chr(13) & chr(10)
	
	if g_Customer_ShowFundOrderBtn_Flag and g_OrderBtnPosition = 1 and fundaspid <> "" then
		Response.Write "<td class=wfb3c>" & g_OrderFiledTitle & "</td>" & chr(13) & chr(10)
	end if
	Response.Write "<td class=wfb3c width=""40%"">基金名稱</td>" & chr(13) & chr(10)
	Response.Write "<td class=wfb3c>基金公司</td>" & chr(13) & chr(10)
	Response.Write "<td class=wfb3c>淨值<BR>日期</td>" & chr(13) & chr(10)  
	Response.Write "<td class=wfb3c>淨值</td>" & chr(13) & chr(10)  
	Response.Write "<td class=wfb3c nowrap>一個月<br>(%)<a href=/w/wq/wq02_" & tid & "_801.djhtm><img border=0 src=/w/images/down.gif></a></td>" & chr(13) & chr(10) 
	Response.Write "<td class=wfb3c nowrap>三個月<br>(%)<a href=/w/wq/wq02_" & tid & "_803.djhtm><img border=0 src=/w/images/down.gif></a></td>" & chr(13) & chr(10) 
	Response.Write "<td class=wfb3c nowrap>六個月<br>(%)<a href=/w/wq/wq02_" & tid & "_806.djhtm><img border=0 src=/w/images/down.gif></a></td>" & chr(13) & chr(10) 
	Response.Write "<td class=wfb3c nowrap>一年<br>(%)<a href=/w/wq/wq02_" & tid & "_901.djhtm><img border=0 src=/w/images/down.gif></a></td>" & chr(13) & chr(10) 
	'Response.Write "<td class=wfb3c nowrap>二年<br>(%)<a href=/w/wq/wq02_" & tid & "_902.djhtm><img border=0 src=/w/images/down.gif></a></td>" & chr(13) & chr(10) 
	'Response.Write "<td class=wfb3c nowrap>三年<br>(%)<a href=/w/wq/wq02_" & tid & "_903.djhtm><img border=0 src=/w/images/down.gif></a></td>" & chr(13) & chr(10) 
	'Response.Write "<td class=wfb3c nowrap>五年<br>(%)<a href=/w/wq/wq02_" & tid & "_905.djhtm><img border=0 src=/w/images/down.gif></a></td>" & chr(13) & chr(10) 
	if g_Customer_ShowFundOrderBtn_Flag and g_OrderBtnPosition = 2 and fundaspid <> "" then
		Response.Write "<td class=wfb3c>" & g_OrderFiledTitle & "</td>" & chr(13) & chr(10)
	end if
	Response.Write "</tr>" & chr(13) & chr(10) 
end sub
%>
