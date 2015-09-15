<!-- #include file="wFundProc.asp" -->
<%
fid = trim(request("A"))
if fid="" or fid="NA" then fid = defFID
if id2="" then id2 = "NA"
if ix1="" then ix1 = "AI000010"
if ix2="" then ix2 = "AI000020"
if yyy="" or yyy="NA" then yyy = "1"

fname = getWFund_FIDName(fid, cid)

'=====================================  基本資料報酬等資料取得 (wb01)   ================================================

Dim s2030,s2180,s2160
Dim t0030,y0030,diff,diffPercent

'取淨值及報酬
sql = "exec wsp_get_roi_info '" & fid & "'"
if OpenWFundDB(conn, rs, sql) then
	
	s2030 = stdfmt(rs("wb102030"),4)	& " (" & trim(rs("wb102020")) & ")"    '最新淨值
	s2180 = rs("wb102180")    '本月以來報酬率
	s2160 = rs("wb102160")    '今年以來報酬率
		
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
end if

'取淨值漲跌
sql = "exec wsp_get_nav_daily '" & fid & "'"
if OpenWFundDB(conn, rs, sql) then
	for icount = 1 to 2
		if icount = 1 then
			t0030 = stdfmt(rs(1),4)
		end if
		if icount = 2 then
			y0030 = stdfmt(rs(1),4)
		end if

		rs.MoveNext
		if rs.Eof then
			exit for
		end if		
	next
		
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
	
	'計算 與昨日比較
	diff = stdfmt((t0030 - y0030),4)
	'計算 漲跌幅
	diffPercent = stdfmt((diff / y0030) * 100 ,4)
	
end if


'=====================================  基本資料報酬等資料取得 (wb01) END ===============================================
lastDate = Date
calDay = CInt(Day(lastDate))
lastDate = DateAdd("d",-calDay,lastDate)
lastDate = DateAdd("m",-4,lastDate)

xxxIDS = ""
xxxIDS = xxxIDS & "<select onchange=""selopn(this.options[this.selectedIndex].value )"" name=""IDS"" size=""1"">" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option selected " & "value=""/w/wb/wb01_" & fid & ".djhtm"">基金基本資料</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wb/wb02_" & fid & ".djhtm"">基金淨值表</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wb/wb03_" & fid & ".djhtm"">基金績效表</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wb/wb04_" & fid & ".djhtm"">基金持股狀況</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wb/wb05_" & fid & ".djhtm"">基金配息狀況</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "</select>" & chr(13) & chr(10)

Response.Write GetDocProlog("基本資料", "wb01", fid, cid)
Response.Write "<script language=""javascript"" src=""/w/js/WFundlistJS.djjs""></script>" & chr(13) & chr(10)
%>
<SCRIPT LANGUAGE=javascript>
<!--
window.onload = GoStart;

function GoStart()
{
	ComboReset('<%=fid%>');
}

function FundGoPage(sObj)
{
	var sURL = '/w/wb/wb01.djhtm?a=';
	var sFID = sObj.selFund3.options[sObj.selFund3.selectedIndex].value;
	if ( sFID != '0' )
		document.location = sURL + sFID;
}
//-->
</SCRIPT>
<%

sql = "exec wsp_get_fid_info '" & fid & "','" & fundaspid & "'  "
'Response.Write sql & "<BR>"
if OpenWFundDB(conn, rs, sql) then
	
  	Response.Write "<div class=""contentfield"">" & chr(13) & chr(10)
   'Response.Write "<div class=""article_block"">" & chr(13) & chr(10)
   Response.Write "<div class=""text_sqzer"">" & chr(13) & chr(10)
   'Response.Write "<div class=""companyselector"">" & chr(13) & chr(10)
   Response.Write "<div class=""wfb0c"">" & chr(13) & chr(10)
   
	Response.Write "<form name=wb01_frm>"
	Response.Write GenComboList(cid,fid,"wb01_frm")
	'Response.Write "<SELECT name=selFID onchange=selopn(this.options[this.selectedIndex].value)>" & chr(13) & chr(10)
	'for k=0 to 9
	'	Response.Write "<OPTION>ＸＸＸＸＸＸＸＸＸＸＸ</OPTION>" & chr(13) & chr(10)
	'next
	'Response.Write "</SELECT>" & chr(13) & chr(10)    
	Response.Write xxxIDS
	
	if g_Customer_ShowWFundOrderBtn_Flag and fundaspid <> "" then
		if ucase(trim(rs("approved"))) = "Y" then
			Response.Write MakeButton(rno,trim(rs("bankfundid"))) & chr(13) & chr(10)
		else

		end if
	end if
	Response.Write "</form>"
	Response.Write "</div>"

	
	'============================================   基本資料 (wb01)   ====================================================
	response.write "<div class=""a_tab_block tab_btm_article"">" & chr(13) & chr(10)
	response.write "  <div class=""aj_block tab_top_article sbg""></div>" & chr(13) & chr(10)
	response.write "  <div class=""squeeze"">" & chr(13) & chr(10)
	response.write "    <div class=""a_tab_block_title"">" & chr(13) & chr(10)
	response.write "      <h5>" & fname & "</h5>" & chr(13) & chr(10)
	response.write "    </div>" & chr(13) & chr(10)

	response.write "    <table>" & chr(13) & chr(10)
	
	if not isnull(rs(2)) then
		Response.Write "      <tr class=""pink"">" & chr(13) & chr(10)
		Response.Write "        <td width=""17%"">總代理公司</td>" & chr(13) & chr(10)
		Response.Write "        <td width=""17%""><span class=""bear""><a href=""/w/wc/wc01_" & trim(rs(2)) & ".djhtm"">" & trim(rs(3)) & "</a></span></td>" & chr(13) & chr(10)
		cid = trim(rs(2))
	else
		Response.Write "      <tr class=""pink"">" & chr(13) & chr(10)
		Response.Write "        <td width=""17%"">總代理公司</td>" & chr(13) & chr(10)
		Response.Write "        <td width=""17%""><span class=""bear"">" & trim(rs(3)) & "</span></td>" & chr(13) & chr(10)
		cid = "NA"
	end if
	Response.Write "        <td width=""17%"" class=""col_head"">基金類型</td>" & chr(13) & chr(10)
	Response.Write "        <td colspan=3 width=""49%""><span class=""bear"">" & trim(rs(8)) & "</span></td>" & chr(13) & chr(10)
	Response.Write "      </tr>" & chr(13) & chr(10)
	
	Response.Write "      <tr class=""pink"">" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"" class=""col_head"">成立日期</td>" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"">" & trim(rs(4)) & "</td>" & chr(13) & chr(10)
	if isnull(rs(6)) then
	    Response.Write "        <td width=""17%"" class=""col_head"">基金規模</td>" & chr(13) & chr(10)
	    Response.Write "        <td width=""16%"">N/A</td>" & chr(13) & chr(10)
	elseif Clng(rs(6)) = 0 then
		Response.Write "        <td width=""17%"" class=""col_head"">基金規模</td>" & chr(13) & chr(10)
		Response.Write "        <td width=""16%"">N/A</td>" & chr(13) & chr(10)
	else
		if CDate(trim(rs(20))) > lastDate then
			Response.Write "        <td width=""17%"" class=""col_head"">基金規模</td>" & chr(13) & chr(10)
			Response.Write "        <td width=""16%"">" & stdfmt(clng(rs(6))/1000,2) & " 百萬" & trim(rs(19)) & "(" & trim(rs(20)) & ")</td>" & chr(13) & chr(10)
		else
			Response.Write "        <td width=""17%"" class=""col_head"">基金規模</td>" & chr(13) & chr(10)
			Response.Write "        <td width=""16%"">" & stdfmt(clng(rs(6))/1000,2) & " 百萬" & trim(rs(19)) & "</td>" & chr(13) & chr(10)
		end if
	end if
	name = "&nbsp;"
	if trim(rs(17)) <> "" and trim(rs(18)) <> "" then
		name = trim(rs(17)) & "(" & trim(rs(18)) & ")"
	elseif trim(rs(17)) <> "" then
		name = trim(rs(17))
	elseif trim(rs(18)) <> "" then
		name = trim(rs(18))
	end if
	Response.Write "        <td width=""16%"" class=""col_head"">基金經理人</td>" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"">" & name & "</td>" & chr(13) & chr(10)
	Response.Write "      </tr>" & chr(13) & chr(10)
	
	Response.Write "      <tr class=""pink"">" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"" class=""col_head"">投資標的</td>" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"">" & trim(rs(10)) & "</td>" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"" class=""col_head"">投資區域</td>" & chr(13) & chr(10)
	Response.Write "        <td width=""16%"">" & trim(rs(9)) & "</td>" & chr(13) & chr(10)
	Response.Write "        <td width=""16%"" class=""col_head"">計價幣別</td>" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"">" & trim(rs(7)) & "</td>" & chr(13) & chr(10)
	Response.Write "      </tr>" & chr(13) & chr(10)
	
	Response.Write "      <tr class=""gray"">" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"" class=""col_head"">最新淨值</td>" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"">" & s2030 & "</td>" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"" class=""col_head"">漲跌幅</td>" & chr(13) & chr(10)
	if diffPercent < 0 then
		Response.Write "        <td colspan=3 width=""49%""><span class=""fall"">" & diffPercent & "%</span></td>" & chr(13) & chr(10)
	else
		Response.Write "        <td colspan=3 width=""49%"">" & diffPercent & "%</td>" & chr(13) & chr(10)
	end if
	Response.Write "      </tr>" & chr(13) & chr(10)
	
	Response.Write "      <tr class=""gray"">" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"" class=""col_head"">本月以來報酬率</td>" & chr(13) & chr(10)
	if isnull(trim(s2180)) or trim(s2180) = "" then
		Response.Write "        <td width=""17%"">NA</td>" & chr(13) & chr(10)	
	else
		if cdbl(s2180) < 0 then
			Response.Write "        <td width=""17%""><span class=""fall"">" & s2180 & "%</span></td>" & chr(13) & chr(10)
		else
			Response.Write "        <td width=""17%"">" & s2180 & "%</td>" & chr(13) & chr(10)
		end if
	end if
	Response.Write "        <td width=""17%"" class=""col_head"">與昨日比較</td>" & chr(13) & chr(10)
	if diff > 0 then
		response.write "        <td colspan=3 width=""49%"">" & diff & "↑</td>" & chr(13) & chr(10)
	elseif diff < 0 then
		response.write "        <td colspan=3 width=""49%""><span class=""fall"">" & diff & "↓</span></td>" & chr(13) & chr(10)
	else
		response.write "        <td colspan=3 width=""49%"">" & diff & "</td>" & chr(13) & chr(10)
	end if
	Response.Write "      </tr>" & chr(13) & chr(10)
	
	Response.Write "      <tr class=""gray"">" & chr(13) & chr(10)
	Response.Write "        <td width=""17%"" class=""col_head"">今年以來報酬率</td>" & chr(13) & chr(10)
	if isnull(trim(s2160)) or trim(s2160) = "" then
		Response.Write "        <td width=""17%"">NA</td>" & chr(13) & chr(10)
	else
		if cdbl(s2160) < 0 then
			Response.Write "        <td width=""17%""><span class=""fall"">" & s2160 & "%</span></td>" & chr(13) & chr(10)
		else
			Response.Write "        <td width=""17%"">" & s2160 & "%</td>" & chr(13) & chr(10)
		end if
	end if
	'風險收益等級
	Response.Write "        <td nowrap width=""17%"" class=""col_head""><a href=http://www.funddj.com/y/notes/rrnotes/rrnotes.htm target=_blank>風險收益等級</a></td>" & chr(13) & chr(10)
	Response.Write "        <td colspan=3 width=""49%""><span class=""bear"">" & GetRiskLevel(fid) & " </td>" & chr(13) & chr(10)
	Response.Write "      </tr>" & chr(13) & chr(10)
	
	Response.Write "    </table>" & chr(13) & chr(10)

	'文件下載：財務報告書/公開說明書/投資人須知 
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <ul class=""cate_addinfo_line"">" & chr(13) & chr(10)
	Response.Write "      <li><em>表單下載 : </em></li>" & chr(13) & chr(10)
	Response.Write "      <li><a href="""& GetFundInfoURL(fid,"2")&""" target=""_blank"">財務報告書</a></li>" & chr(13) & chr(10)
	Response.Write "      <li><a href="""& GetFundInfoURL(fid,"1")&""" target=""_blank"">公開說明書</a></li>" & chr(13) & chr(10)
	Response.Write "      <li><a href="""& GetFundInfoURL(fid,"3")&""" target=""_blank"">投資人須知</a></li>" & chr(13) & chr(10)
	Response.Write "    <a href="""& GetFundMonthReport(fid,int(Timer))&""" target=""_blank"">基金月報</a>" & vbcrlf
	'Response.Write "      <li><a href=""#"" onclick=""javascript:return sOpenTradeRule('" & fid & "');"" class=""end"">短線交易規定</a></li>" & chr(13) & chr(10)
	
	Response.Write "    </ul>" & chr(13) & chr(10)
	Response.write "<ul class=""cate_addinfo_line"">" 
	Response.Write "<li><em>投資人服務及保護 : </em></li>" & vbcrlf
	'Response.Write "<td class=""wfb6l"" colspan=""3"">" & vbcrlf
	Response.Write " <li><a href=""#"" onclick=""javascript:return sOpenTradeRule('" & fid & "','SwingTrade');"">短線交易規定</a></li>" & vbcrlf
	Response.Write "  <li><a href=""#"" onclick=""javascript:return sOpenTradeRule('" & fid & "','Fair');"">公平價格調整機制</a></li>" & vbcrlf
	Response.Write "  <li><a href=""#"" onclick=""javascript:return sOpenTradeRule('" & fid & "','AntiDilute');"">反稀釋機制</a></li>" & vbcrlf
	Response.Write "    </ul>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)

	  
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
	
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)
	
	
	'商品 DM
	'Response.Write "<ul class=""dmlink"">" & chr(13) & chr(10)
	'Response.Write "  <li class=""prod_dm""><a href=""#"" title=""商品DM"">商品DM</a></li>" & chr(13) & chr(10)
	'Response.Write "  <li class=""more""><a href=""#"" title=""MORE"">MORE</a></li>" & chr(13) & chr(10)
	'Response.Write "</ul>" & chr(13) & chr(10)
	
	Response.Write "    <h4>&nbsp;</h4>" & chr(13) & chr(10)
	'====================================================================================================================


	'============================================  基金淨值 (wb02)   ====================================================
	
	Response.Write "<div class=""article_inline_block pusher"">" & chr(13) & chr(10)
	Response.Write "  <div class=""squeezer"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h5>基金淨值走勢圖</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	
	Response.Write "    <table>"
	  
	sql = "exec wsp_get_fid_info '" & fid & "'"
	if OpenWFundDB(conn, rs, sql) then
		currencyType = trim(rs("wb100070"))
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if
	  
	Response.Write "      <tr><td colspan=8>" & chr(13) & chr(10)
	Response.write "        <applet archive=""CURVE1.jar"" CODE=""CURVE1.class"" codebase=""/w/jar"" HEIGHT=""182"" WIDTH=""320"" VIEWASTEXT id=Applet1>" & chr(13) & chr(10)
	Response.write "          <param name=""BCD"" value=""/w/bcd/BCDNavList_" & fid & "_1.djbcd"">" & chr(10)
	'Response.write "          <param name=""T"" value=""" & fname & "基金淨值走勢圖"">" & chr(10)
	Response.write "          <param name=""T"" value="""">" & chr(10)
	Response.write "          <param name=""U"" value=""元(" & currencyType & ")"">" & chr(10)
	Response.write "          <param name=""BC"" value=""fff8e5"">" & chr(10)
	Response.write "          <param name=""LC"" value=""0000FF"">" & chr(10)
	Response.write "        </applet>" & chr(10)
	Response.Write "      </td></tr>" & chr(13) & chr(10)
	
	
	rcnt = 0
	dim ary(30,3)
	sql = "exec wsp_get_nav_daily '" & fid & "'"
	if OpenWFundDB(conn, rs, sql) then
		idx = 0
		while not rs.EOF 
			ary(idx,0) = rs(0)
			ary(idx,1) = rs(1)
			ary(idx,2) = rs(2)
			ary(idx,3) = rs(3)
			idx = idx + 1
			rs.MoveNext
		wend
		rcnt = idx
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if
	
	if rcnt > 0 then
		Response.Write NavCol(ary, 0,  3, rcnt)
	else
		Response.Write "<tr><td>無淨值資料</td></tr>" & chr(13) & chr(10)
	end if
	
	Response.Write "</table>" & chr(13) & chr(10)
	
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)
	Response.Write "</td>" & chr(13) & chr(10)
	

	'====================================================================================================================


	'============================================  基金績效 (wb03)   ====================================================
	
	Response.Write "<div class=""article_inline_block"">" & chr(13) & chr(10)
	Response.Write "  <div class=""squeezer"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h5>基金績效勢圖</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)

	sql = "exec wsp_get_roi_info '" & fid & "'"
	if OpenWFundDB(conn, rs, sql) then
		Response.Write "    <table>" & chr(13) & chr(10)
			
		Response.Write "      <tr><td>" & chr(13) & chr(10)
		Response.Write "        <applet archive=""MCURVE5.jar"" CODE=""MCURVE5.class"" codebase=""/w/jar"" NAME=MCURVE5 HEIGHT=182 WIDTH=320 VIEWASTEXT id=MCURVE5>" & chr(13) & chr(10)
		Response.Write "          <param name=""BCD"" value=""/w/bcd/BCDROIList5_" & fid & "_NA_NA_NA_NA_NA_NA_1.djbcd"">" & chr(13) & chr(10)
		Response.Write "          <param name=""CAPTION"" value="""">" & chr(13) & chr(10)
		response.Write "          <param name=""BC"" value=""fff8e5"">" & chr(13) & chr(10)
		response.Write "          <param name=""LC"" value=""000000 00AAAA AAAA00 AA00AA 0000AA"">" & chr(13) & chr(10)
		response.Write "          <param name=""T"" value=""" & fname & """>" & chr(13) & chr(10)
		response.Write "          <param name=""U"" value=""　 　 　 　 　"">" & chr(13) & chr(10)
		response.Write "        </applet>" & chr(13) & chr(10)
		Response.Write "      </td></tr>" & chr(13) & chr(10)
		Response.Write "    </table>" & chr(13) & chr(10)
		
		Response.Write "    <table>" & chr(13) & chr(10)
		Response.Write "      <thead>" & chr(13) & chr(10)
		Response.Write "        <tr>" & chr(13) & chr(10)
		Response.Write "          <td>一個月</td>" & chr(13) & chr(10)
		Response.Write "          <td>三個月</td>" & chr(13) & chr(10)
		Response.Write "          <td>六個月</td>" & chr(13) & chr(10)
		Response.Write "          <td>年化標準差</td>" & chr(13) & chr(10)
		Response.Write "          <td>Sharpe</td>" & chr(13) & chr(10)
		Response.Write "          <td>Beta</td></tr>" & chr(13) & chr(10)
		Response.Write "        </tr>" & chr(13) & chr(10)
		Response.Write "      </thead>" & chr(13) & chr(10)
		
		Response.Write "        <tr>" & chr(13) & chr(10)
		ShowData(rs("wb102090"))
		ShowData(rs("wb102100"))
		ShowData(rs("wb102110"))
		ShowData(rs("wb103030"))
		ShowData(rs("wb103020"))
		ShowData(rs("wb103010"))
		Response.Write "        </tr>" & chr(13) & chr(10)
		
		Response.Write "</table>"
		

		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
			
	else 
		Response.Write "<table>" & chr(13) & chr(10)
		Response.Write "<tr><td class=wfb6c>無績效資料</td></tr>" & chr(13) & chr(10)
		Response.Write "</table>" & chr(13) & chr(10)
	end if

	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)

	'====================================================================================================================
	
	
	'============================================  持股比例 (wb04)   ====================================================
	
	Response.Write "<div class=""a_tab_block tab_btm_article holdpercent cleartitle"">" & chr(13) & chr(10)
	Response.Write "  <div class=""aj_block tab_top_article sbg""></div>" & chr(13) & chr(10)
	Response.Write "  <div class=""squeeze"">" & chr(13) & chr(10)
	Response.Write "    <div class=""a_tab_block_title"">" & chr(13) & chr(10)
	Response.Write "      <h5>持股比例</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	Response.Write "    <div class=""block_composer"">" & chr(13) & chr(10)
	
	x1 = GetFundHold(fid, "400")
	x2 = GetFundHold(fid, "410")
	if (x1 = "") and (x2 = "") then
		Response.Write "    <table>"
		Response.Write "      <tr><td>" & chr(13) & chr(10)
		Response.Write "        <p align=center>無持股資料</p>"
		Response.Write "      </td></tr>" & vbcrlf
		Response.Write "    </table>" & chr(13) & chr(10)
	else
		Response.Write x1
		Response.Write x2
	end if
	
	Response.Write "      <div class=""cleartitle""></div>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)
	
	
	'====================================================================================================================
	
	
	'============================================  配息狀況 (wb05)   ====================================================
	
	
	Response.Write "<div class=""inline_block_out"">" & chr(13) & chr(10)
	Response.Write "  <div class=""a_tab_block tab_btm_article holdpercent cleartitle"">" & chr(13) & chr(10)
	Response.Write "    <div class=""aj_block tab_top_article sbg""></div>" & chr(13) & chr(10)
	Response.Write "    <div class=""squeeze"">" & chr(13) & chr(10)
	Response.Write "      <div class=""a_tab_block_title"">" & chr(13) & chr(10)
	Response.Write "        <h5>配息記錄</h5>" & chr(13) & chr(10)
	Response.Write "      </div>" & chr(13) & chr(10)
	Response.Write "      <div class=""inline_block_alt"">" & chr(13) & chr(10)

	Response.Write "        <table class=""twincolour"" >" & chr(13) & chr(10)
	  
	Response.Write "          <thead>" & chr(13) & chr(10)
	Response.Write "            <tr>" & chr(13) & chr(10)
	Response.Write "              <td>日期</td>" & chr(13) & chr(10)
	Response.Write "              <td>狀態</td>" & chr(13) & chr(10)
	Response.Write "              <td>息值/比例</td>" & chr(13) & chr(10)
	Response.Write "              <td>幣別</td>" & chr(13) & chr(10)
	Response.Write "            </tr>" & chr(13) & chr(10)
	Response.Write "          </thead>" & chr(13) & chr(10)

	sql = "select * from wa210000 where wb210010='" & fid & "' order by wb210020 desc"
	'Response.Write sql & "<BR>"

	if OpenWFundDB(conn, rs, sql) then
	
		rno = 0
		for icount = 1 to 4
			rno = rno + 1
			cs = "odd"
			if rno mod 2 = 0 then cs = "" 
			if rs("wb210050")= "S" then
				if cdbl(rs(2)) >= 1 then
					Response.Write "          <tr class=" & cs & ">" & chr(13) & chr(10)
					Response.Write "            <td>" & trim(rs(1)) & "</td>" & chr(13) & chr(10)
					Response.Write "            <td>合併</td>" & chr(13) & chr(10)
					Response.Write "            <td>" & formatnumber(rs(2),2) & ":1</td>" & chr(13) & chr(10)
					Response.Write "            <td>&nbsp;</td>" & chr(13) & chr(10)
					Response.Write "          </tr>" & chr(13) & chr(10)
				else
					Response.Write "          <tr class=" & cs & ">" & chr(13) & chr(10)
					Response.Write "            <td>" & trim(rs(1)) & "</td>" & chr(13) & chr(10)
					Response.Write "            <td>分割</td>" & chr(13) & chr(10)
					Response.Write "            <td>1:" & formatnumber(1/cdbl(rs(2)),2) & "</td>" & chr(13) & chr(10)
					Response.Write "            <td>&nbsp;</td>" & chr(13) & chr(10)
					Response.Write "          </tr>" & chr(13) & chr(10)
				end if
			else
				Response.Write "          <tr class=" & cs & ">" & chr(13) & chr(10)
				Response.Write "            <td>" & trim(rs(1)) & "</td>" & chr(13) & chr(10)
				if isnull(trim(rs(5))) or trim(rs(5)) = "" then
					Response.Write "            <td>配息</td>" & chr(13) & chr(10)
					Response.Write "            <td>" & trim(rs(2)) & "</td>" & chr(13) & chr(10)
					Response.Write "            <td>" & trim(rs(3)) & "</td>" & chr(13) & chr(10)
					Response.Write "          </tr>" & chr(13) & chr(10)
				elseif isnull(trim(rs(2))) or trim(rs(2)) = "" or cdbl(trim(rs(2))) = 0 then
					Response.Write "            <td>稅後配息</td>" & chr(13) & chr(10)
					Response.Write "            <td>" & trim(rs(5)) & "</td>" & chr(13) & chr(10)
					Response.Write "            <td>" & trim(rs(3)) & "</td>" & chr(13) & chr(10)
					Response.Write "          </tr>" & chr(13) & chr(10)
				else
					Response.Write "            <td>稅前配息 / 稅後配息</td>" & chr(13) & chr(10)
					Response.Write "            <td>" & trim(rs(2)) & " / " & trim(rs(5)) & "</td>" & chr(13) & chr(10)
					Response.Write "            <td>" & trim(rs(3)) & "</td>" & chr(13) & chr(10)
					Response.Write "          </tr>" & chr(13) & chr(10)
				end if
			end if
			rs.MoveNext
			if rs.Eof then
				exit for
			end if		
		next
		
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
		
		Response.Write "      </table>" & chr(13) & chr(10)
	
	else 
		Response.Write "        <tr><td class=wfb6c colspan=4>無配息資料</td></tr>" & chr(13) & chr(10)
		Response.Write "      </table>" & chr(13) & chr(10)
	end if
	
	Response.Write "      <a href=""/w/wb/wb05_" & fid & ".djhtm"" title=""MORE"" class=""more"">MORE</a> </div>" & chr(13) & chr(10)
	Response.Write "    <h4>&nbsp;</h4>" & chr(13) & chr(10)
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)
	

	'====================================================================================================================
		
	'response.write "      </div>" & chr(13) & chr(10)
	'response.write "    </div>" & chr(13) & chr(10)
	response.write "  </div>" & chr(13) & chr(10)
	response.write "</div>" & chr(13) & chr(10)
	
else
	Response.Write "<table>" & chr(13) & chr(10)
	Response.Write "<tr><td class=wfb6c colspan=4>無基金基本資料</td></tr>" & chr(13) & chr(10)
	Response.Write "</table>" & chr(13) & chr(10)
end if

Response.Write "<script language=""javascript""><!--" & chr(13) & chr(10)
'Response.Write "InitComboList(document.wb01_frm.selFID, '/w/wb/wb01_', '.djhtm', '" & fid & "', wfund_fund);" & chr(13) & chr(10)
'== 短線交易規定 open window function ==
Response.Write "function sOpenTradeRule(sFID,sType)" & vbcrlf
Response.Write "{	" & vbcrlf
Response.Write "	var sURL = '/w/wb/wb01rule.djhtm?a='+sFID+ '&b='+sType;	" & vbcrlf
Response.Write "	window.open(sURL,'newwindow',config='width=600,height=300,top=0,left=0,toolbar=0,menubar=0,scrollbars=yes,resizable=no,location=no,status=no');	" & vbcrlf
Response.Write "	return false;	" & vbcrlf
Response.Write "}	" & vbcrlf
Response.Write "// --></script>" & chr(13) & chr(10)
  	
Response.Write GetDocEplog("Q")


'基金淨值資料
function NavCol(ary, first, last, limit)
	dim xxx,i
	xxx = ""
	rno = 0
	for i = 0 to 1
		'xxx = xxx & "		<tr>" & chr(13) & chr(10)
		if i = 0 then
			xxx = xxx & "      <tr><td nowrap>日期</td>" & chr(13) & chr(10)
		else
			xxx = xxx & "      <tr><td nowrap>淨值</td>" & chr(13) & chr(10)
		end if
		for idx=first to last
			if i = 0 then
				if idx >= limit then
					fd0 = "&nbsp;"
				else
					fd0 = stdfmt(ary(idx,0),3)
				end if
				xxx = xxx & "        <td>" & fd0 & "</td>" & chr(13) & chr(10)
			else
				if idx >= limit then
					fd1 = "&nbsp;"
					xxx = xxx & "        <td>" & fd1 & "</td>" & chr(13) & chr(10)
				else
					'fd1 = stdfmt(ary(idx,1),dot)
					fd1 = stdfmt(ary(idx,1),4)
					if fd1 < 0 then
						xxx = xxx & "        <td><span class=""fall"">" & fd1 & "</span></td>" & chr(13) & chr(10)
					else
						xxx = xxx & "        <td>" & fd1 & "</td>" & chr(13) & chr(10)
					end if
				end if
			end if
		next
		xxx = xxx & "      </tr>" & chr(13) & chr(10)
	next
	
	NavCol = xxx
end function  


'基金績效數值(判斷正負)
Function ShowData(sData)

	if not isnull(sData) then
		if cdbl(sData) < 0 then
			Response.Write "          <td nowrap><span class=""fall"">" & sData & "</span></td>" & chr(13) & chr(10)
		else
			Response.Write "          <td nowrap>" & sData & "</td>" & chr(13) & chr(10)
		end if
	else
		Response.Write "          <td nowrap>" & sData & "</td>" & chr(13) & chr(10)
	end if

end Function


'基金持有類股圖
Function GetFundHold(fid, table)
	dim MyCounter
	dim MyTotal
	dim MyTotal2
	dim T
	dim V
	dim I
	dim C
	dim S
	Dim sAppletTitle
	Dim sTotalValue
	Dim sDTable
	Dim sPercent
	Dim sValue
	dated = ""

	'-- 2006/01/19 modified by cuteduck Start --
	sName =""
	dt = "萬元"
	dValue = 10
	strSql = "exec spj_mda72151 '" & fid & "'"		
	aRs =  OpenSQL_Fund(strSql)
	if not isEmpty(aRs)  then 
		FundCurrency = trim(aRs(9,0))
		if FundCurrency = "日圓" then
			dt = "百萬"
			dValue = 1000
		end if
	end if  
	'-- 2006/01/19 modified by cuteduck End --  

	if table = "400" then 
		sql = "exec wsp_fundhold_400 '" & fid & "'"
		sAppletTitle = "地區"
	elseif table = "410" then 
		sql = "exec wsp_fundhold_410 '" & fid & "'"
		sAppletTitle = "產業"
	else
		sql = ""
	end if

	if sql = "" then 
		GetFundHold = ""
		exit function
	end if
	
	'取得合計的值
	sTotalValue = 0
	if openWFundDB(conn,rs,sql) = true then
		do While Not rs.EOF
			if trim(rs(3)) <> 0 then
					sTotalValue = sTotalValue + mFormatNumber2(cdbl(trim(rs(3))) / dValue,0)
			end if
			rs.MoveNext
		loop
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	else
		GetFundHold = ""
		exit function
	end if
	
	sDTable = ""
	sPercent = 0
	sumd = 0
	if openWFundDB(conn,rs,sql) = true then
		MyCounter = 0
		do While Not rs.EOF
			sumd = sumd +1
			MyCounter = MyCounter +1
			
		'-- 2006/01/19 modified by cuteduck Start --
			T = T & " " & trim(replace(trim(rs(2)),"��",""))  '去除全形空白
			V = V & " " & formatnumber(cdbl(trim(rs(3))) / dValue,0,0,0,0)
			C = C & " " & GetColor(MyCounter)        
		'-- 2006/01/19 modified by cuteduck End --                 
			if trim(rs(3)) <> 0 then
					sValue = mFormatNumber2(cdbl(trim(rs(3))) / dValue,0)
			else
				sValue = 0
			end if

			dated = rs(1)
			if MyCounter = 1 then
				S = "1"
			else
				S = S & " " & "0"
			end if
			
			'================以表格顯示持股分佈================
			if sumd = 1 then
				sDTable = sDTable & "        <table class=""rightside twincolour"" style=""width: 450px !important;"">" & chr(13) & chr(10)
				sDTable = sDTable & "          <thead>" & chr(13) & chr(10)
				sDTable = sDTable & "            <tr><td>名稱</td><td>值</td><td>比例</td><td>名稱</td><td>值</td><td>比例</td></tr>" & vbcrlf
				sDTable = sDTable & "          </thead>" & chr(13) & chr(10)
			end if
			if (sumd mod 2 = 1) then
				if rno mod 2 = 0 then
					sDTable = sDTable & "            <tr>" & vbcrlf
				else
					sDTable = sDTable & "            <tr class=""odd"">" & vbcrlf
				end if
			end if
			sDTable = sDTable & "              <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(sumd) & ";"">&nbsp;</span> " & trim(replace(trim(rs(2)),"��","")) & "</div></td>" & vbcrlf
			sDTable = sDTable & "              <td>" & sValue & "</td>" & vbcrlf
			sPercent = formatnumber((sValue / sTotalValue) * 100)
			sDTable = sDTable & "              <td>" & sPercent & "%</td>" & vbcrlf
			if (sumd mod 2 = 0) then
				sDTable = sDTable & "            </tr>" & vbcrlf
				rno = rno + 1
			end if
			'========================================================
			
			rs.MoveNext
		loop
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
		
		'====================顯示合計的部份======================
		'判斷是否為單數,若是,則要補齊商品右側空白的另一半
		if (sumd mod 2 = 1) then
			sDTable = sDTable & "              <td>&nbsp;</td>" & vbcrlf
			sDTable = sDTable & "              <td>&nbsp;</td>" & vbcrlf
			sDTable = sDTable & "              <td>&nbsp;</td></tr>" & vbcrlf
		end if
		
		if rno mod 2 = 0 then
			sDTable = sDTable & "            <tr class=""odd"">" & vbcrlf
		else
			sDTable = sDTable & "            <tr>" & vbcrlf
		end if
		sDTable = sDTable & "              <td colspan=3>&nbsp;</td>" & vbcrlf
		sDTable = sDTable & "              <td>合計</td>" & vbcrlf
		sDTable = sDTable & "              <td><span class=""bear"">" & sTotalValue & "</span></td>" & vbcrlf
		sDTable = sDTable & "              <td><span class=""bear"">100%</span></td>" & vbcrlf
		sDTable = sDTable & "            </tr>" & vbcrlf
		sDTable = sDTable & "          </table>" & vbcrlf

		'========================================================
		
	else
		GetFundHold = ""
		exit function
	end if
	  
	
	xxx = xxx & "      <table>" & chr(10)
	xxx = xxx & "        <tr><td valign=""top"">" & chr(10)
	xxx = xxx & "              <applet ARCHIVE=""PIE2DNoTable.jar"" codebase=""/w/jar"" CODE=""PIE2DNoTable.class""  width=220 height=187 VIEWASTEXT id=Applet1>" & chr(10)
	xxx = xxx & "                <param name=""T"" value=""基金持股分佈 " & T & """>" & chr(10)
	xxx = xxx & "                <param name=""V"" value=""" & V & """>" & chr(10)
	xxx = xxx & "                <param name=""C"" value=""" & C & """>" & chr(10)
	xxx = xxx & "                <param name=""S"" value=""" & S & """>" & chr(10)
	xxx = xxx & "                <param name=""F"" value=""1"">" & chr(10)
	xxx = xxx & "                <param name=""BkColor"" value=""fff8e5"">" & chr(10)
	xxx = xxx & "              </applet>" & chr(10)
	xxx = xxx & "            </td>" & chr(10)
	xxx = xxx & "            <td valign=""top"">" & chr(10)
	xxx = xxx & sDTable
	xxx = xxx & "            </td></tr></table>" & chr(10)
	
	GetFundHold = xxx
end function

function GetFundtopHold(fid)
	if Trim(fid & "") = "" then
		exit function
	end if
	
	xxx = ""
	sql = "exec spj_mda72853 '" & fid & "'"
	if OpenFundDJ(conn, rs, sql) then
		dated = rs(0)

		xxx = xxx & "      <table valign=""top"" class=""twincolour"" >" & chr(13) & chr(10)
		xxx = xxx & "        <thead>" & chr(13) & chr(10)
		xxx = xxx & "          <tr>" & chr(13) & chr(10)
		xxx = xxx & "            <td>持股名稱</td>" & chr(13) & chr(10)
		xxx = xxx & "            <td>比例&nbsp;</td>" & chr(13) & chr(10)
		xxx = xxx & "          </tr>" & chr(13) & chr(10)
		xxx = xxx & "        </thead>" & chr(13) & chr(10)

		rno = 0
		do while not rs.EOF 
			rno = rno + 1
			cs = "odd"
			if rno mod 2 = 0 then cs = "" 
						      
			xxx = xxx & "        <tr class=" & cs & ">" & chr(13) & chr(10)
			xxx = xxx & "          <td>" & stdfmt(rs(1),0) & "</td>" & chr(13) & chr(10)
			if isnull(rs(3)) then
				rrr = "&nbsp;"
			else
				rrr = formatnumber(rs(3),2) & "%"
			end if

			xxx = xxx & "          <td>" & rrr & "&nbsp;</td>" & chr(13) & chr(10)
			xxx = xxx & "        </tr>" & chr(13) & chr(10)
			rs.MoveNext
		loop
			    
		xxx = xxx & "      </table>" & vbcrlf
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if
	GetFundtopHold = xxx
end function

%>
