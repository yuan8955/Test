<!-- #include file="wtFundProc.asp" -->
<%
fid = ucase(trim(request("A")))
if fid="" or fid="NA" then fid = deftFID

idname = getFundDJ_IDName(fid)

'=====================================  基本資料報酬等資料取得 (wr01)   ================================================
Dim s801,s300,s310,s320,s330,s340,fname,s120

sql = "exec tsp_get_fund_info_800 '" & fid & "'"

if OpenFundDJ(conn, rs, sql) then
	dot = 2
	dcg = 6
	getDecimal rs("yb800120"), dot, dcg
	
	s300 = stdfmt(rs("yb800300"),dot) & " (" & trim(rs("yb800010")) & ")"     '最近淨值
	s330 = stdfmt(rs("yb800330"),2)       '成立日起報酬率
	s340 = stdfmt(rs("yb800340"),2)       '今年以來報酬率
	s801 = stdfmt(rs("yb800801"),2)       '一個月報酬率
	s310 = stdfmt(rs("yb800310"),2)       '與昨日比較
	s320 = stdfmt(rs("yb800320"),2)       '漲跌幅
	fname = trim(rs("fundname"))
	cid =  trim(rs("yb800140"))
	
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
end if

'最近一次配息資料
sql = "exec YS130000 '" & fid & " '"
if OpenFundDJ(conn, rs, sql) then
	s120 = trim(rs(1)) & " (" & trim(rs(0)) & ")"
end if
if s120 = "" then
	's120 = "沒有配息資料"
	s120 = "N/A"
end if

'============================================  持股比例計算 (wr04)   ====================================================
dim MyCounter
dim MyTotal
dim MyTotal2
dim T
dim V
dim I
dim C
dim S
dim showFlag
showFlag = false

dim sitebase
sitebase = 0
sql = "exec tsp_get_fund_info '" & fid & "'"
if OpenFundDJ(conn, rs, sql) then
	if not isnull(trim(rs(9))) and trim(rs(9)) <> "" then
		sitebase = cdbl(trim(rs(9))) *100
	end if
		
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
end if  
  
sql = "exec tsp_get_fund_hold120 '" & fid & "'"
'sql = "exec tsp_get_fund_hold120 'ACCA20'"
sumd = 0
if openFundDJ(conn,rs,sql) = true then
	showFlag = true
	MyCounter = 0
	dated=FormatYMD(trim(rs("本次日期")))
	do While Not rs.EOF 
		if ucase(trim(rs(1))) <> "EBTOTAL" and  ucase(trim(rs(1)))<> "EBZAA" then
			sumd = sumd +1
			MyCounter = MyCounter +1
			if left(ucase(trim(rs(1))),3)="EB0" then 
				T = T & " " & "上市" & replace(replace(trim(rs(5)),"��","")," ","_") 
			elseif left(ucase(trim(rs(1))),3)="EB1" then 
				T = T & " " & "上櫃" & replace(replace(trim(rs(5)),"��","")," ","_")
			else
				T = T & " " &  replace(replace(trim(rs(5)),"��","")," ","_")
			end if

			if isnumeric (trim(rs(2))) then
				sValue = cdbl(trim(rs(2))) * sitebase
			else
				sValue = rs(2)
			end if
			
			V = V & " " & Formatnumber(sValue,0,0,0,0)
			
			I = I & " " & formatnumber(rs(3),2)
			C = C & " " & GetColor(MyCounter)
			if MyCounter = 1 then
				S = "1"
			else
				S = S & " " & "0"
			end if
			      
			if isnull(rs(2)) = false then
				Mytotal = MyTotal + cdbl(rs(2))
			end if
						
			if isnull(rs(4))=false  then
				Mytotal2 = MyTotal2 + cdbl(rs(4))	
			end if
		end if
		rs.MoveNext
	loop
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
	    
	if MyTotal < 100 then
		MyTotal2 = MyTotal - MyTotal2
		MyTotal = 100 - MyTotal		
		
		T = mid(T,2) & " 其它"
		
		V = mid(V,2) & " " & formatnumber(MyTotal*sitebase,0,0,0,0)

		I = mid(I,2) & " " & formatnumber(MyTotal2) 
		C = mid(C,2) & " " & GetColor(MyCounter+1)
		S = S & " 0"
	end if
end if


'====================================================================================================================
Response.Write GetDocProlog("基金基本資料", "wr01", fid, "NA", "NA")
Response.Write "<script language=""javascript"" src=""/w/js/WtFundlistJS.djjs""></script>" & chr(13) & chr(10)

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
	var sURL = '/w/wr/wr01.djhtm?a=';
	var sFID = sObj.selTFund3.options[sObj.selTFund3.selectedIndex].value;
	if ( sFID != '0' )
		document.location = sURL + sFID;
}
//-->
</SCRIPT>
<%

xxxIDS = ""
xxxIDS = xxxIDS & "<select onchange=""selopn(this.options[this.selectedIndex].value )"" name=""IDS"" size=""1"">" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option selected " & "value=""/w/wr/wr01_" & fid & ".djhtm"">基金基本資料</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr02_" & fid & ".djhtm"">基金淨值表</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr03_" & fid & ".djhtm"">基金績效表</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr04_" & fid & ".djhtm"">基金持股狀況</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr10_" & fid & ".djhtm"">基金配息狀況</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr05_" & fid & ".djhtm"">基金相關新聞</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "</select>" & chr(13) & chr(10)


area = ""
sql = "exec tsp_get_fund_area '" & fid & "'"
if OpenFundDJ(conn, rs, sql) then
	while not rs.EOF
		if area <> "" then area = area & "，"
		area = area & trim(rs(2))
		rs.MoveNext
	wend
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
end if

sale = ""
sql = "exec tsp_get_fund_sale '" & fid & "'"
if OpenFundDJ(conn, rs, sql) then
	while not rs.EOF
		if sale <> "" then sale = sale & "，"
		sale = sale & trim(rs(2))
		rs.MoveNext
	wend
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
end if
  
sss = ""
sql = "exec tsp_get_fund_bank '" & fid & "'"
if OpenFundDJ(conn, rs, sql) then
	while not rs.EOF 
		if sss <> "" then sss = sss & "、"
		sss = sss  & trim(rs(1)) & "日"
		rs.movenext
	wend
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
	sss = sss & "&nbsp;(實際扣款日依申購人與銷售機構約定)"
end if
  

sql = "exec tsp_get_fund_info '" & fid & "','" & fundaspid & "'"
'Response.Write sql & "<BR>"
if OpenFundDJ(conn, rs, sql) then

	'Response.Write "<div class=""outbox"">" & chr(13) & chr(10)
  	Response.Write "<div class=""contentfield"">" & chr(13) & chr(10)
   'Response.Write "<div class=""article_block"">" & chr(13) & chr(10)
   Response.Write "<div class=""text_sqzer"">" & chr(13) & chr(10)
   'Response.Write "<div class=""companyselector"">" & chr(13) & chr(10)
   Response.Write "<div class=""wfb0c"">" & chr(13) & chr(10)
   Response.Write "<FORM method=POST name=wr01_frm align=center onsubmit=""return false;"">" & chr(13) & chr(10)
   Response.Write GenComboListTW(cid,fid,"wr01_frm")
'	Response.Write "<SELECT name=selFID onchange=selopn(this.options[this.selectedIndex].value)>" & chr(13) & chr(10)
'	for selcnt=1 to 5
'		Response.Write "<OPTION>ＸＸＸＸＸＸＸＸＸ</OPTION>" & chr(13) & chr(10)
'	next
'	Response.Write "</SELECT>" & xxxIDS & chr(13) & chr(10)
	Response.Write xxxIDS & chr(13) & chr(10)
	
	if g_Customer_ShowFundOrderBtn_Flag and fundaspid <> "" then
		if ucase(trim(rs("approved"))) = "Y" then
			Response.Write MakeButton(rno,trim(rs("bankfundid"))) & chr(13) & chr(10)
		else
			'Response.Write "<td class=" & cs & "c></td>" & chr(13) & chr(10)
		end if
	end if
	Response.Write "</form>" & chr(13) & chr(10)

	Response.Write "</div>" & chr(13) & chr(10)

	
	'============================================   基本資料 (wr01)   ====================================================

	Response.Write "<div class=""a_tab_block tab_btm_article"">" & chr(13) & chr(10)
	Response.Write "  <div class=""aj_block tab_top_article sbg""></div>" & chr(13) & chr(10)
	Response.Write "  <div class=""squeeze"">" & chr(13) & chr(10)
	Response.Write "    <div class=""a_tab_block_title"">" & chr(13) & chr(10)
	Response.Write "      <h5>" & fname & "</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	

	response.write "    <table>" & chr(13) & chr(10)
	
	response.write "      <tr class=""pink"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">基金公司</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%""><span class=""bear""><a href=""/w/wp/wp01_" & trim(rs(2)) & ".djhtm"">" & trim(rs(3)) & "</a></span></td>"  & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""17%"">基金類型</td>" & chr(13) & chr(10)
	response.write "        <td colspan=3 width=""46%""><span class=""bear""><a href=""/w/wr/wr00A_" & trim(rs(4)) & ".djhtm"">" & trim(rs(5)) & "</a></span></td>" & chr(13) & chr(10)
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""pink"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">成立日期</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%"">" & stdfmt(rs(8),5) & "</td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""17%"">成立時規模(億)</td>" & chr(13) & chr(10)
	response.write "        <td width=""13%"">" & stdfmt(rs(10),2) & "</td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""16%"">基金規模(億)</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%"">" & stdfmt(rs(9),2) & " (" & trim(rs("yb100311")) & ")</td>" & chr(13) & chr(10)
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""pink"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">基金經理人</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%""><a href=""/w/wt/wt03_" & trim(rs(6)) & ".djhtm"">" & trim(rs(7)) & "</a></td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""17%"">投資區域</td>" & chr(13) & chr(10)
	response.write "        <td width=""13%"">" & area & "&nbsp;</td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""16%"">計價幣別</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%"">" & stdfmt(rs(16),0) & "</td>" & chr(13) & chr(10)
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""gray"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">一個月報酬率</td>" & chr(13) & chr(10)

	if isnull(s801) or trim(s801) = "" or trim(s801) = "&nbsp;" then
		response.write "        <td width=""17%"">0.00%</span></td>" & chr(13) & chr(10)
	else
		if s801 < 0 then
			response.write "        <td width=""17%""><span class=""fall"">" & s801 & "%</span></td>" & chr(13) & chr(10)
		else
			response.write "        <td width=""17%"">" & s801 & "%</td>" & chr(13) & chr(10)
		end if
	end if
	response.write "        <td class=""col_head"" width=""17%"">最新淨值</td>" & chr(13) & chr(10)
	response.write "        <td colspan=3 width=""46%"">" & s300 & "</td>" & chr(13) & chr(10)
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""gray"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">成立日起報酬率</td>" & chr(13) & chr(10)
	if s330 < 0 then
		response.write "        <td width=""17%""><span class=""fall"">" & s330 & "%</span></td>" & chr(13) & chr(10)
	else
		response.write "        <td width=""17%"">" & s330 & "%</td>" & chr(13) & chr(10)
	end if
	response.write "        <td class=""col_head"" width=""17%"">與昨日比較</td>" & chr(13) & chr(10)
	if s310 > 0 then
		response.write "        <td width=""13%"">" & s310 & "↑</td>" & chr(13) & chr(10)
	elseif s310 < 0 then
		response.write "        <td width=""13%""><span class=""fall"">" & s310 & "↓</span></td>" & chr(13) & chr(10)
	else
		response.write "        <td width=""13%"">" & s310 & "</td>" & chr(13) & chr(10)
	end if
	response.write "        <td class=""col_head"" width=""16%"">漲跌幅</td>" & chr(13) & chr(10)
	if s320 < 0 then
		response.write "        <td width=""17%""><span class=""fall"">" & s320 & "%</span></td>" & chr(13) & chr(10)
	else
		response.write "        <td width=""17%"">" & s320 & "%</td>" & chr(13) & chr(10)
	end if
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""gray"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">今年以來報酬率</td>" & chr(13) & chr(10)
	if s340 < 0 then
		response.write "        <td width=""17%""><span class=""fall"">" & s340 & "%</span></td>" & chr(13) & chr(10)
	else
		response.write "        <td width=""17%"">" & s340 & "%</td>" & chr(13) & chr(10)
	end if
	response.write "        <td class=""col_head"" nowrap width=""17%""><a href=http://www.funddj.com/y/notes/rrnotes/rrnotes.htm target=_blank>風險收益等級</a></td>" & chr(13) & chr(10)
	response.write "        <td width=""13%""><span class=""bear"">" & GetRiskLevel(fid) & "</span></td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""16%"">最近配息</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%"">" & s120 & "</td>" & chr(13) & chr(10)	
	response.write "      </tr>" & chr(13) & chr(10)
	Response.Write "    </table>" & chr(13) & chr(10)
	
	
	'文件下載 : 2008/10/09 modified
	Response.Write "    <div align=""left"">" & vbcrlf
	Response.Write "    <ul class=""cate_addinfo_line"">" & vbcrlf
	Response.Write "      <li><em>表單下載：</em></li>" & vbcrlf
	Response.Write "      <li><a href=""" & GetTWFundInfoURL(fid,"2") & """ target=""_blank"" title=""財務報告書"">財務報告書</a></li>" & vbcrlf
	Response.Write "      <li><a href=""" & GetTWFundInfoURL(fid,"1") & """ target=""_blank"" title=""公開說明書"">公開說明書</a></li>" & vbcrlf
	Response.Write GetFundEasyReport(fid)
	Response.Write "      <li><a href=""#""  onclick=""javascript:return sOpenTradeRule('" & fid & "');"">短線交易規定</a></li>" & vbcrlf
    Response.Write "	  <li><a class=""end"" href="""& GetFundMonthReport(fid,int(Timer))&""" target=""_blank"">基金月報</a></li>" & vbcrlf

	Response.Write "    </ul>" & vbcrlf
	Response.Write "    </div>" & chr(13) & chr(10)
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)

	
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
	
		
	'商品DM
	'Response.Write "<ul class=""dmlink"">" & chr(13) & chr(10)
	'Response.Write "  <li class=""prod_dm""><a href=""#"" title=""商品DM"">商品DM</a></li>" & chr(13) & chr(10)
	'Response.Write "  <li class=""more""><a href=""#"" title=""MORE"">MORE</a></li>" & chr(13) & chr(10)
	'Response.Write "</ul>" & chr(13) & chr(10)
	
	Response.Write "    <h4>&nbsp;</h4>" & chr(13) & chr(10)
	'====================================================================================================================
	
	
	'============================================  基金淨值 (wr02)   ====================================================

	Response.Write "<div class=""article_inline_block pusher"">" & chr(13) & chr(10)
	Response.Write "  <div class=""squeezer"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h5>基金淨值走勢圖</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	
	sql = "exec tsp_get_fund_info '" & fid & "'"
	if OpenFundDJ(conn, rs, sql) then
		currencyType = trim(rs("currency"))
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if

	Response.Write "    <table>" & chr(13) & chr(10)
	Response.Write "      <tr><td colspan=8>" & chr(13) & chr(10)
	Response.write "        <applet   archive=""/w/jar/CURVE1.jar"" CODE=""CURVE1.class""  HEIGHT=""182"" WIDTH=""320"" VIEWASTEXT id=Applet1>" & chr(13) & chr(10)
	Response.write "          <param name=""BCD"" value=""/w/bcd/tBCDNavList_" & fid & "_1.djbcd"">" & chr(10)
	'Response.write "          <param name=""T"" value=""" & idname & "基金淨值走勢圖"">" & chr(10)
	Response.write "          <param name=""T"" value="""">" & chr(10)
	Response.write "          <param name=""U"" value=""元(" & currencyType & ")"">" & chr(10)
	Response.write "          <param name=""BC"" value=""fff8e5"">" & chr(10)
	Response.write "          <param name=""LC"" value=""0000FF"">" & chr(10)
	Response.write "        </applet>" & chr(10)
	Response.Write "      </td></tr>" & chr(13) & chr(10)
	
	rcnt = 0
	dim ary(30,2)
	sql = "exec tsp_get_nav_daily '" & fid & "'"
	if OpenFundDJ(conn, rs, sql) then
		idx = 0
		while not rs.EOF 
			ary(idx,0) = rs(0)
			ary(idx,1) = rs(1)
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
		Response.Write NavCol(ary, 0,  4, rcnt)
		Response.Write "    </table>" & chr(13) & chr(10)
	else
		Response.Write "<tr><td>無淨值資料</td></tr>" & chr(13) & chr(10)
		Response.Write "</table>" & chr(13) & chr(10)
	end if
	
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)
	'====================================================================================================================
	

	'============================================  基金績效 (wr03)   ====================================================
	
	Response.Write "<div class=""article_inline_block"">" & chr(13) & chr(10)
	Response.Write "  <div class=""squeezer"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h5>基金績效勢圖</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	
	Response.Write "    <table>" & chr(13) & chr(10)
	Response.Write "      <tr><td>" & chr(13) & chr(10)
	Response.Write "        <applet archive=""MCURVE5.jar"" codebase=""/w/jar"" CODE=""MCURVE5.class"" NAME=MCURVE5 HEIGHT=182 WIDTH=320 VIEWASTEXT id=MCURVE5>" & chr(13) & chr(10)
	Response.Write "          <param name=""BCD"" value=""/w/bcd/BCDROIList5_" & fid & "_NA_NA_NA_NA_NA_NA_1.djbcd"">" & chr(13) & chr(10)
	Response.Write "          <param name=""CAPTION"" value="""">" & chr(13) & chr(10)
	response.Write "          <param name=""BC"" value=""fff8e5"">" & chr(13) & chr(10)
	response.Write "          <param name=""LC"" value=""000000 00AAAA AAAA00 AA00AA 0000AA"">" & chr(13) & chr(10)
	response.Write "          <param name=""T"" value=""" & idname & """>" & chr(13) & chr(10)
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
	

	sql = "exec tsp_get_fund_info_800 '" & fid & "'"
	if OpenFundDJ(conn, rs, sql) then
		dot = 2
		dcg = 6
		getDecimal rs("yb800120"), dot, dcg
		
		Response.Write "      <tr>" & chr(13) & chr(10)
		ShowData(rs("yb800801"))
		ShowData(rs("yb800803"))
		ShowData(rs("yb800806"))
		ShowData(rs("yb800350"))
		ShowData(rs("yb800360"))
		ShowData(rs("yb800370"))
		Response.Write "      </tr>" & chr(13) & chr(10)

		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if

	Response.Write "    </table>" & chr(13) & chr(10)
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)


	'====================================================================================================================

	'============================================  持股比例 (wr04)   ====================================================

	Response.Write "<div class=""a_tab_block tab_btm_article holdpercent cleartitle"">" & chr(13) & chr(10)
	Response.Write "  <div class=""aj_block tab_top_article sbg""></div>" & chr(13) & chr(10)
	Response.Write "  <div class=""squeeze"">" & chr(13) & chr(10)
	Response.Write "    <div class=""a_tab_block_title"">" & chr(13) & chr(10)
	Response.Write "      <h5>持股比例</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	Response.Write "    <div class=""block_composer""> " & chr(13) & chr(10)
	
	
	if sitebase <> 0 and showFlag then
		Response.Write "      <table>" & chr(13) & chr(10)
		Response.Write "        <tr><td valign=""top"">" & chr(13) & chr(10)
		Response.Write "          <applet ARCHIVE=""PIE2DNoTable.JAR"" CODE=""PIE2DNoTable.class"" codebase=""/w/jar"" width=220 height=187px VIEWASTEXT id=Applet1>" & chr(13) & chr(10)
		Response.Write "            <param name=""T"" value=""基金持股分佈 " & T & """>" & chr(13) & chr(10)
		Response.Write "            <param name=""V"" value=""" & V & """>" & chr(13) & chr(10)
		Response.Write "            <param name=""C"" value=""" & C & """>" & chr(13) & chr(10)
		Response.Write "            <param name=""S"" value=""" & S & """>" & chr(13) & chr(10)
		Response.Write "            <param name=""F"" value=""0"">" & chr(13) & chr(10)
		Response.Write "            <param name=""BkColor"" value=""fff8e5"">" & chr(13) & chr(10)
		Response.Write "          </applet>" & chr(13) & chr(10)
		Response.Write "    </td>" & chr(13) & chr(10)
	'----------------------------------以表格顯示投資比例--------------------------------------------
		Response.Write "      <td valign=""top"">" & chr(13) & chr(10)
		Response.Write "      <table class=""rightside twincolour"" style=""width: 450px !important;"">" & chr(13) & chr(10)
		Response.Write "        <thead>" & chr(13) & chr(10)
		Response.Write "          <tr>" & chr(13) & chr(10)
		Response.Write "            <td>名稱</td>" & chr(13) & chr(10)
		Response.Write "            <td>值</td>" & chr(13) & chr(10)
		Response.Write "            <td>比例</td>" & chr(13) & chr(10)
		Response.Write "            <td>名稱</td>" & chr(13) & chr(10)
		Response.Write "            <td>值</td>" & chr(13) & chr(10)
		Response.Write "            <td>比例</td>" & chr(13) & chr(10)
		Response.Write "          </tr>" & chr(13) & chr(10)
		Response.Write "        </thead>" & chr(13) & chr(10)
		
		'顯示"依產業"的PIE圖
		sql = "exec tsp_get_fund_hold120 '" & fid & "'"
		dim scount,sTotal,sTotal2,sTotalValue,sPercent
		scount = 0
		sTotal = 0
		sTotal2 = 0
		sTotalValue = 0
		sPercent = 0
		rno = 0
		if openFundDJ(conn,rs,sql) = true then
			rno = 1
			do While Not rs.EOF 
				if ucase(trim(rs(1))) <> "EBTOTAL" and  ucase(trim(rs(1)))<> "EBZAA" then
					scount = scount +1
					
					
					if (scount mod 2 = 1) then
						if (rno mod 2 = 0) then
							response.write "        <tr>" & chr(13) & chr(10) 
						else
							response.write "        <tr class=""odd"">" & chr(13) & chr(10) 
						end if
					end if
					'名稱
					if left(ucase(trim(rs(1))),3)="EB0" then 
						response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount) & ";"">&nbsp;</span> 上市" & replace(replace(trim(rs(5)),"��","")," ","_") & "</div></td>" & chr(13) & chr(10) 
					elseif left(ucase(trim(rs(1))),3)="EB1" then 
						response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount) & ";"">&nbsp;</span> 上櫃" & replace(replace(trim(rs(5)),"��","")," ","_") & "</div></td>" & chr(13) & chr(10) 
					else
						response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount) & ";"">&nbsp;</span> " &  replace(replace(trim(rs(5)),"��","")," ","_") & "</div></td>" & chr(13) & chr(10) 
					end if
						
					'值
					if isnumeric (trim(rs(2))) then
						sValue = cdbl(trim(rs(2))) * sitebase
					else
						sValue = rs(2)
					end if
					response.write "          <td>" & Formatnumber(sValue,0,0,0,0) & "</td>" & chr(13) & chr(10) 
					if sValue >= 1 then
						sTotalValue = sTotalValue + Formatnumber(sValue,0,0,0,0)
					end if
					
					'比例
					response.write "          <td>" & formatnumber(rs(2)) & "%</td>" & chr(13) & chr(10) 
					sPercent = sPercent + formatnumber(rs(2))
					
					if isnull(rs(2)) = false then
						sTotal = sTotal + cdbl(rs(2))
					end if
								
					if isnull(rs(4))=false  then
						sTotal2 = sTotal2 + cdbl(rs(4))	
					end if
				
					if (scount mod 2 = 0) then
						response.write "        </tr>" & chr(13) & chr(10)
						rno = rno + 1
					end if
					
				end if
				
				rs.MoveNext
			loop
			rs.close
			conn.close
			set rs = nothing
			set conn = nothing
			
			'判斷是否為單數,若是,則要補齊商品右側空白的另一半
			if (scount mod 2 = 1) and (sTotal >= 100) then
				response.write "              <td>&nbsp;</td>" & vbcrlf
				response.write "              <td>&nbsp;</td>" & vbcrlf
				response.write "              <td>&nbsp;</td></tr>" & vbcrlf
			end if
			
			
			'取得其它的部份
			cs = "wfb1"
			if rno mod 2 = 0 then cs = "wfb2" 
			if (scount mod 2 = 1) then
				if sTotal < 100 then
					sTotal2 = sTotal - sTotal2
					sTotal = 100 - sTotal		
					response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount + 1) & ";"">&nbsp;</span> 其它</div></td>" & chr(13) & chr(10)
					response.write "          <td>" & formatnumber(sTotal*sitebase,0,0,0,0) & "</td>" & chr(13) & chr(10)
					response.write "          <td>" & formatnumber(sTotal) & "%</td>" & chr(13) & chr(10)
					response.write "        </tr>" & chr(13) & chr(10) 
					if isnumeric(formatnumber(sTotal*sitebase,0,0,0,0)) then
						sTotalValue = sTotalValue + Formatnumber(sTotal*sitebase,0,0,0,0)
					end if
					sPercent = sPercent + formatnumber(sTotal)
				end if
			else
				if sTotal < 100 then
					sTotal2 = sTotal - sTotal2
					sTotal = 100 - sTotal		
					if (rno mod 2 = 0) then
						response.write "        <tr>" & chr(13) & chr(10) 
					else
						response.write "        <tr class=""odd"">" & chr(13) & chr(10) 
					end if
					response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount + 1) & ";"">&nbsp;</span> 其它</div></td>" & chr(13) & chr(10)
					response.write "          <td>" & formatnumber(sTotal*sitebase,0,0,0,0) & "</td>" & chr(13) & chr(10)
					response.write "          <td>" & formatnumber(sTotal) & "%</td>" & chr(13) & chr(10)
					response.write "          <td>&nbsp;</td>" & chr(13) & chr(10)
					response.write "          <td>&nbsp;</td>" & chr(13) & chr(10)
					response.write "          <td>&nbsp;</td>" & chr(13) & chr(10)
					response.write "        </tr>" & chr(13) & chr(10) 
					sTotalValue = sTotalValue + Formatnumber(sTotal*sitebase,0,0,0,0)
					sPercent = sPercent + formatnumber(sTotal)
				end if
			end if
			
			
			'====================顯示合計的部份======================
	
			if (rno mod 2 = 1) then
				response.write "        <tr>" & chr(13) & chr(10) 
			else
				response.write "        <tr class=""odd"">" & chr(13) & chr(10) 
			end if

			response.write "          <td colspan=3>&nbsp;</td>" & chr(13) & chr(10)
			response.write "          <td> 合計</td>" & chr(13) & chr(10)
			response.write "          <td><span class=""bear"">" & sTotalValue & "</span></td>" & chr(13) & chr(10)
			response.write "          <td><span class=""bear"">" & sPercent & "%</span></td>" & chr(13) & chr(10)
			response.write "        </tr>" & chr(13) & chr(10) 
					
		end if
		Response.Write "      </table>" & chr(13) & chr(10)
	
		Response.Write "      </td></tr>" & chr(13) & chr(10)
		Response.Write "      </table>" & chr(13) & chr(10)

		'顯示"依地區"的PIE圖
		Response.Write GetData0(fid)
	
	else
		Response.Write "    <table>" & chr(13) & chr(10)
		Response.Write "      <tr><td>無持股資料</td></tr>" & chr(13) & chr(10)
		Response.Write "    </table>" & chr(13) & chr(10)
	end if
	
	Response.Write "      <div class=""cleartitle""></div>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)
	

	'====================================================================================================================
	

	'============================================  相關新聞 (wr05)   ====================================================

	Response.Write "<div class=""inline_block_out"">" & chr(13) & chr(10)
	Response.Write "  <div class=""inline_block"" style=""width: 340px; margin-right: 20px;"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h2 class=""tag_title tw_stock_curve"">相關新聞</h2>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	
	Response.Write "    <table class=""twincolour"">" & chr(13) & chr(10)

	Response.write sayNewsList(page)
	Response.Write "    </table>" & chr(13) & chr(10)
	Response.write "<a href=""/w/wr/wr05_" & fid & ".djhtm"" class=""more"">MORE</a> </div>" & chr(13) & chr(10)
		
	'============================================  基金配息 (wr10)   ====================================================
	
	Response.Write "  <div class=""inline_block"" style=""width: 340px;"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h2 class=""tag_title international_clock"">配息記錄</h2>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	
	Response.Write "    <table class=""twincolour"">" & chr(13) & chr(10)

	Response.Write "      <thead>" & chr(13) & chr(10)
	Response.Write "        <tr><td>日期</td>" & chr(13) & chr(10)
	Response.Write "          <td>狀態</td>" & chr(13) & chr(10)
	Response.Write "          <td>息值/比例</td>" & chr(13) & chr(10)
	Response.Write "          <td>幣別</td></tr>" & chr(13) & chr(10)
	Response.Write "      </thead>" & chr(13) & chr(10)
	
	sql = "exec YS130000 '" & fid & " '"
	if OpenFundDJ(conn, rs, sql) then
	
		rno = 0

		for icount = 1 to 4
			rno = rno + 1
			cs = "wfb1"
			if rno mod 2 = 0 then
				Response.Write "      <tr>" & chr(13) & chr(10)
			else
				Response.Write "      <tr class=""odd"">" & chr(13) & chr(10)
			end if
				
			Response.Write "        <td>" & trim(rs(0)) & "</td>" & chr(13) & chr(10)
			Response.Write "        <td>配息</td>" & chr(13) & chr(10)
			Response.Write "        <td>" & trim(rs(1)) & "</td>" & chr(13) & chr(10)
			        
			Select Case trim(rs(2))
				case "AX000010" '台幣
					currName = "台幣"
				case "AX000060" '歐元
					currName = "歐元"
				case "AX000070" '人民幣
					currName = "人民幣"
				case "AX000085" '馬來幣		
					currName = "馬來幣"
				case "AX000090" '印尼幣		
					currName = "印尼幣"			
				case else
					currName = "台幣"
			End Select
				            
			Response.Write "        <td class=" & cs & "c>" & currName & "</td>" & chr(13) & chr(10)
			Response.Write "      </tr>" & chr(13) & chr(10)
			
			rs.MoveNext
			if rs.Eof then
				exit for
			end if		
		next
	else 
		Response.Write "      <tr><td colspan=4>無配息資料</td></tr>" & chr(13) & chr(10)
	end if
	Response.Write "    </table>" & chr(13) & chr(10)
	Response.write "    <a href=""/w/wr/wr10_" & fid & ".djhtm"" class=""more"">MORE</a> </div>" & chr(13) & chr(10)
	
	'====================================================================================================================
	
	response.write "</div>" & chr(13) & chr(10)
	'response.write "</div>" & chr(13) & chr(10)
	response.write "</div>" & chr(13) & chr(10)
	'response.write "</div>" & chr(13) & chr(10)

else
	response.write "<table width=580>" & chr(13) & chr(10)
	Response.Write "<tr><td class=wfb0c>無基金基本資料</td></tr>" & chr(13) & chr(10)
	Response.Write "</table>" & chr(13) & chr(10)
end if
    
    
Response.Write "<script language=""JavaScript""><!--" & chr(13) & chr(10)
'== 2008/10/09 短線交易規定 open window function ==
Response.Write "function sOpenTradeRule(sFID)" & vbcrlf
Response.Write "{	" & vbcrlf
Response.Write "	var sURL = '/w/wr/wr01rule.djhtm?a='+sFID;	" & vbcrlf
Response.Write "	window.open(sURL,'newwindow',config='width=600,height=300,top=0,left=0,toolbar=0,menubar=0,scrollbars=yes,resizable=no,location=no,status=no');	" & vbcrlf
Response.Write "	return false;	" & vbcrlf
Response.Write "}	" & vbcrlf

'Response.Write "InitComboList(document.wr01_frm.selFID, '/w/wr/wr01_', '.djhtm', '" & fid & "', tfund_fund, '');" & chr(13) & chr(10)
Response.Write "// --></script>" & chr(13) & chr(10)

Response.Write GetDocEplog("Q")

'基金淨值資料
function NavCol(ary, first, last, limit)
	dim xxx,i
	xxx = ""
	rno = 0
	for i = 0 to 1
		
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
					fd1 = stdfmt(ary(idx,1),2)
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
			Response.Write "          <td nowrap><span class=""fall"">" & stdfmt(sData,2) & "</span></td>" & chr(13) & chr(10)
		else
			Response.Write "          <td nowrap>" & stdfmt(sData,2) & "</td>" & chr(13) & chr(10)
		end if
	else
		Response.Write "          <td nowrap>" & stdfmt(sData,2) & "</td>" & chr(13) & chr(10)
	end if

end Function


'基金持有類股圖
function GetData0(fid)
	Dim sSQL,aRs,sD,sIdx,sDTable
	dim MyCounter
	dim MyTotal
	dim MyTotal2
	dim sTotalValue
	dim sPercent
	dim T
	dim V
	dim I
	dim C
	dim S

	GetData0 = ""
	sD = ""
	sDTable = ""
	
	dot = 2
	dcg = 6

	sumd = 0
	scount = 0
	sSQL = "exec spj_mda72851 '" & fid & "'"
	'Response.Write ssql & "<BR>"
	set aRs = nothing
	aRs=OpenSQL_Fund(sSQL)
	sIdx = 0
	if  isEmpty(aRs) then
		GetData0 = sD
		exit function
	else 
		sIdx = sIdx + 1
		MyCounter = 0
		sTotalValue = 0
		sPercent = 0
		dated = FormatYMD(aRs(0,0))

		for forIdx = 0 to ubound(aRs,2)
			Tmp1 = ucase(chkdata(aRs(3,forIdx),sSetDefformat))
			if Tmp1 <> "合計" then
				cs = "wfb1"
				if rno mod 2 = 0 then cs = "wfb2" 
				
				sumd = sumd +1
				MyCounter = MyCounter +1
				scount = scount + 1
				Tmp1 = replace(replace(Tmp1,"��","")," ","_")  '去除全形空白			
				T = T & " " &  Tmp1 
				
					
				sValue = aRs(6,forIdx) 
				if isnumeric (trim(aRs(6,forIdx))) then
					sValue = cdbl(trim(aRs(6,forIdx)))/10
				end if
				'====計算合計的值及比例====
				if isnumeric(formatnumber(sValue,0,0,0,0)) then
					sTotalValue = sTotalValue + Formatnumber(sValue,0,0,0,0)
				end if
				sPercent = sPercent + formatnumber(aRs(4,forIdx))
				'==========================
				
				V = V & " " & Formatnumber(sValue,0,0,0,0)
				I = I & " " & "0.00"
				
				C = C & " " & GetColor(MyCounter)
				if MyCounter = 1 then
					S = "1"
				else
					S = S & " " & "0"
				end if
				
				'================以表格顯示持股地區分佈================
				if scount = 1 then
					sDTable = sDTable & "        <table class=""rightside twincolour"" style=""width: 450px !important;"">" & chr(13) & chr(10)
					sDTable = sDTable & "          <thead>" & chr(13) & chr(10)
				 	sDTable = sDTable & "            <tr>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>名稱</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>值</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>比例</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>名稱</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>值</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>比例</td>" & chr(13) & chr(10)
					sDTable = sDTable & "            </tr>" & chr(13) & chr(10)
					sDTable = sDTable & "          </thead>" & chr(13) & chr(10)
				end if
				if (scount mod 2 = 1) then
					if rno mod 2 = 0 then
						sDTable = sDTable & "            <tr>" & vbcrlf
					else
						sDTable = sDTable & "            <tr class=""odd"">" & vbcrlf
					end if
				end if
				sDTable = sDTable & "              <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount) & ";"">&nbsp;</span> " & Tmp1 & "</div></td>" & vbcrlf
				sDTable = sDTable & "              <td>" & Formatnumber(sValue,0,0,0,0) & "</td>" & vbcrlf
				sDTable = sDTable & "              <td>" & formatnumber(aRs(4,forIdx)) & "%</td>" & vbcrlf
				if (scount mod 2 = 0) then
					sDTable = sDTable & "            </tr>" & vbcrlf
					rno = rno + 1
				end if
				'========================================================

			end if			
		next
	
		T = mid(T,2)
		V = mid(V,2)
		I = mid(I,2)
		C = mid(C,2)
		
		'===========================================================
		'判斷是否為單數,若是,則要補齊右側的另一半
		if (scount mod 2 = 1) then
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
		sDTable = sDTable & "              <td> 合計</td>" & vbcrlf
		sDTable = sDTable & "              <td><span class=""bear"">" & sTotalValue & "</span></td>" & vbcrlf
		sDTable = sDTable & "              <td><span class=""bear"">" & sPercent & "%</span></td>" & vbcrlf
		sDTable = sDTable & "            </tr>" & vbcrlf
		sDTable = sDTable & "          </table>" & vbcrlf
		'===========================================================
	end if
		
	set aRs = nothing	


	sD = sD & "      <table>" & chr(13) & chr(10)
	sD = sD & "        <tr><td valign=""top"">" & chr(13) & chr(10)
	sD = sD & "            <applet ARCHIVE=""PIE2DNoTable.JAR"" CODE=""PIE2DNoTable.class"" codebase=""/w/jar"" width=220 height=187 VIEWASTEXT id=Applet1>" & chr(13) & chr(10)
	sD = sD & "              <param name=""T"" value=""基金持股分佈 " & T & """>" & chr(13) & chr(10)
	sD = sD & "              <param name=""V"" value=""" & V & """>" & chr(13) & chr(10)
	sD = sD & "              <param name=""C"" value=""" & C & """>" & chr(13) & chr(10)
	sD = sD & "              <param name=""S"" value=""" & S & """>" & chr(13) & chr(10)
	sD = sD & "              <param name=""F"" value=""1"">" & chr(13) & chr(10)
	sD = sD & "              <param name=""BkColor"" value=""fff8e5"">" & chr(13) & chr(10)
	sD = sD & "            </applet>" & chr(13) & chr(10)
	sD = sD & "          </td>" & chr(13) & chr(10)
	'-----------  產業別----------
	sD = sD & "          <td valign=""top"">" & chr(13) & chr(10)
	sD = sD & sDTable 
	sD = sD & "          </td></tr></table>" & chr(13) & chr(10)


	rcnt = 0
	set aRs = nothing
	GetData0 = sD
end function		

'基金新聞
function sayNewsList(page)
	dim xxx 
	xxx = ""
	sql = "exec tsp_get_news_list null,'" & fid & "',null"
	
	if OpenFundDJ(conn, rs, sql) then
	  
		rno = 0
		
		for iPage = 1 to 3
			rno = rno + 1
			cs = "wfb1"
			if rno mod 2 = 0 then
				xxx = xxx & "      <tr>" & chr(13) & chr(10)
			else
				xxx = xxx & "      <tr class=""odd"">" & chr(13) & chr(10)
			end if
			xxx = xxx & "        <td>" & formatYYMD(rs(0)) & "</td>" & chr(13) & chr(10)
			xxx = xxx & "        <td><div align=""left""><a href=""/w/wp/wp05A_" & trim(rs(4)) & ".djhtm"">" & trim(rs(1)) & "</a></div></td>" & chr(13) & chr(10)
			xxx = xxx & "      </tr>" & chr(13) & chr(10)
			rs.MoveNext
			if rs.Eof then
				exit for
			end if		
		next
			
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	else
		xxx = xxx & "      <tr><td>無新聞資料</td</tr>" & chr(13) & chr(10)
	end if
	sayNewsList = xxx
end function

%>

