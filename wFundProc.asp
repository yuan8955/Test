<!--#include file="..\AspDB.asp" -->
<!--#include file="..\RTHost.asp" -->
<!--#include file="BFProc.asp" -->
<!--#include Virtual="/Lib/FilterCSFun.asp" -->
<%
Response.Expires = 0
Response.CacheControl = "private23"

Function OpenWFundDB(conn, rs, sql)
	if SQLInjectionFilter(sql) then 
		exit function
	end if
	
	on error resume next
	Set conn = Server.CreateObject("ADODB.Connection")
	If err.number Then
		Set conn = Nothing
		OpenWFundDB = False
		Exit Function
	End if
	If conn is Nothing Then
		OpenWFundDB = False
		Exit function
	End if
	
	conn.cursorlocation = 3
	conn.Open GetDBconnStrWFund()
	If err.number Then
		Set conn = Nothing
		OpenWFundDB = False
		Exit Function
	End if
	  
	Set rs = Server.CreateObject("ADODB.Recordset")
	If err.number Then
		conn.Close
		Set conn = Nothing
		Set rs = Nothing
		OpenWFundDB = False
		Exit Function
	End if
	If rs is Nothing Then
		conn.Close
		Set Conn = Nothing
		Set rs = Nothing
		OpenWFundDB = False
		Exit Function
	End if
	  
	conn.CommandTimeout = 360
	rs.Open sql, conn, 3
	If err.number Then
		conn.Close
		Set rs = Nothing
		Set conn = Nothing
		OpenWFundDB = False
		Exit Function
	End if
	If rs.State <> 1 or rs.BOF or rs.EOF then
		rs.Close
		conn.Close
		Set rs = Nothing
		Set conn = Nothing
		OpenWFundDB = False
		Exit Function
	End if
	OpenWFundDB = True
End Function

'EOF will be reported as error
Function OpenJust(conn, rs, sql)
	if SQLInjectionFilter(sql) then 
		exit function
	end if
	
	on error resume next
	Set conn = Server.CreateObject("ADODB.Connection")
	If err.number Then
		Set conn = Nothing
		OpenJust = False
		Exit Function
	End if
	If conn is Nothing Then
		OpenJust = False
		Exit function
	End if
	  
	conn.cursorlocation = 3  
	conn.Open GetDBConnStr()
	If err.number Then
		Set conn = Nothing
		OpenJust = False
		Exit Function
	End if
	  
	Set rs = Server.CreateObject("ADODB.Recordset")
	If err.number Then
		conn.Close
		Set conn = Nothing
		Set rs = Nothing
		OpenJust = False
		Exit Function
	End if
	If rs is Nothing Then
		conn.Close
		Set Conn = Nothing
		Set rs = Nothing
		OpenJust = False
		Exit Function
	End if
	  
	conn.CommandTimeout = 360
	rs.CursorLocation = 3
	rs.Open sql, conn, 3
	If err.number Then
		conn.Close
		Set rs = Nothing
		Set conn = Nothing
		OpenJust = False
		Exit Function
	End if
	If rs.State <> 1 or rs.BOF or rs.EOF then
		rs.Close
		conn.Close
		Set rs = Nothing
		Set conn = Nothing
		OpenJust = False
		Exit Function
	End if
	OpenJust = True
End Function

'EOF will be reported as error
Function OpenFundDJ(conn, rs, sql)
	if SQLInjectionFilter(sql) then 
		exit function
	end if
	
	on error resume next
	Set conn = Server.CreateObject("ADODB.Connection")
	If err.number Then
		Set conn = Nothing
		OpenFundDJ = False
		Exit Function
	End if
	If conn is Nothing Then
		OpenFundDJ = False
		Exit function
	End if
	  
	conn.cursorlocation = 3  
	conn.Open GetDBconnStrFund()
	If err.number Then
		Set conn = Nothing
		OpenFundDJ = False
		Exit Function
	End if
	  
	Set rs = Server.CreateObject("ADODB.Recordset")
	If err.number Then
		conn.Close
		Set conn = Nothing
		Set rs = Nothing
		OpenFundDJ = False
		Exit Function
	End if
	If rs is Nothing Then
		conn.Close
		Set Conn = Nothing
		Set rs = Nothing
		OpenFundDJ = False
		Exit Function
	End if
	  
	conn.CommandTimeout = 360
	rs.Open sql, conn, 3
	If err.number Then
		conn.Close
		Set rs = Nothing
		Set conn = Nothing
		OpenFundDJ = False
		Exit Function
	End if
	If rs.State <> 1 or rs.BOF or rs.EOF then
		rs.Close
		conn.Close
		Set rs = Nothing
		Set conn = Nothing
		OpenFundDJ = False
		Exit Function
	End if
	OpenFundDJ = True
End Function

Function OpenCFODB(conn, rs, sql)
	if SQLInjectionFilter(sql) then 
		exit function
	end if
	
	on error resume next
	Set conn = Server.CreateObject("ADODB.Connection")
	If err.number Then
		Set conn = Nothing
		OpenCFODB = False
		Exit Function
	End if
	If conn is Nothing Then
		OpenCFODB = False
		Exit function
	End if
	conn.cursorlocation = 3  
	conn.Open GetDBconnStrCFO()
	If err.number Then
		Set conn = Nothing
		OpenCFODB = False
		Exit Function
	End if
	Set rs = Server.CreateObject("ADODB.Recordset")
	If err.number Then
		conn.Close
		Set conn = Nothing
		Set rs = Nothing
		OpenCFODB = False
		Exit Function
	End if
	If rs is Nothing Then
		conn.Close
		Set Conn = Nothing
		Set rs = Nothing
		OpenCFODB = False
		Exit Function
	End if
	conn.CommandTimeout = 360
	'Response.Write SQL
	rs.Open sql, conn, adOpenStatic
	If err.number Then
		conn.Close
		Set rs = Nothing
		Set conn = Nothing
		OpenCFODB = False
		Exit Function
	End if
	'If rs.State <> adStateOpen or rs.BOF or rs.EOF then
	if rs.bof=true and rs.eof=true then
		rs.Close
		conn.Close
		Set rs = Nothing
		Set conn = Nothing
		OpenCFODB = False
		Exit Function
	End if
	OpenCFODB = True
End Function

function stdfmt(p_val, p_fmt)
	dim rc
	if isnull(p_val) then
		rc = "&nbsp;"
	else
		select case p_fmt
			case 0
				rc = trim(p_val)
			case 2
				rc = formatnumber(p_val,2) 
			case 3
				rc = month(p_val) & "/" & day(p_val)
			case 4
				rc = formatnumber(p_val,4) 
			case 5
				rc = year(p_val) & "年" & month(p_val) & "月"
			case 8
				rc = right("0" & month(p_val),2) & "/" & right("0" & day(p_val), 2)
			case 6
				if cdbl(p_val) > 0 then
					rc = "<font color=#ff0000>" & formatnumber(p_val,2) & "</font>"
				elseif cdbl(p_val) < 0 then
					rc = "<font color=#006600>" & formatnumber(p_val,2) & "</font>" 
				else
					rc = formatnumber(p_val,2) 
				end if
			case 7
				if cdbl(p_val) > 0 then
					rc = "<font color=#ff0000>" & formatnumber(p_val,2) & "%</font>"
				elseif cdbl(p_val) < 0 then
					rc = "<font color=#006600>" & formatnumber(p_val,2) & "%</font>" 
				else
					rc = formatnumber(p_val,2) 
				end if
		end select
	end if
	stdfmt = rc
end function

function FundMsgText(p_tag)
	dim xxx : xxx = ""
	
	xxx = xxx & "<table border=""0"" class=""wfb4l"">" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註1:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">各基金經金管會核准或同意生效，惟不表示絕無風險，基金經理公司以往之經理績效不保證基金之最低投資收益；基金經理公司除盡善良管理人之注意義務外，不負責各基金之盈虧，亦不保證最低之收益，投資人申購前應詳閱基金公開說明書。投資基金所應承擔之相關風險及應負擔之費用(境外基金含分銷費用)已揭露於基金公開說明書或投資人須知中，投資人可至公開資訊觀測站或境外基金資訊觀測站查閱。基金並非存款，基金投資非屬存款保險承保範圍投資人需自負盈虧。基金投資具投資風險，此一風險可能使本金發生虧損，其中可能之最大損失為全部信託本金。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註2:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">績效計算為原幣別報酬，且皆有考慮配息情況。基金配息率不代表基金報酬率，且過去配息率不代表未來配息率。所有基金績效，均為過去績效，不代表未來之績效表現，亦不保證基金之最低投資收益。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註3:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">基金淨值可能因市場因素而上下波動，基金淨值僅供參考，實際以基金公司公告之淨值為準；海外市場指數類型基金，以交易日當天收盤價為淨值參考價；部份基金採雙軌報價，實際交易以基金公司所公告的買回價／賣出價為計算基礎。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註4:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">有關基金應負擔之費用（含分銷費用）已揭露於基金之公開說明書及投資人須知中，投資人可至境外基金資訊觀測站(http://www.fundclear.com.tw)下載，或逕向總代理人網站查閱。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註5:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">上述銷售費用僅供參考，實際費率以各銷售機構為主。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註6:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">上述短線交易規定資料僅供參考，實際規定應以基金公開說明書為主。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註7</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">標示""(已撤銷核備)""及""(未申報生效)""之基金資料僅提供原有投資者參考，以供其做買回、轉換或繼續持有之決定；「已撤銷核備」為在台灣已下架(停售)的基金；「未申報生效」為未於法規規定之時效內向金管會辦理申報生效作業之境外基金。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註8:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">基金B股在贖回時，基金公司將依持有期間長短收取不同比率之遞延申購手續費，該費用將自贖回總額中扣除；另各基金公司需收取一定比率之分銷費用，將反映於每日基金淨值中，投資人無需額外支付。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註9:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">依金管會規定基金投資大陸證券市場之有價證券不得超過本基金資產淨值之10%，當該基金投資地區包含中國大陸及香港，基金淨值可能因為大陸地區之法令、政治或經濟環境改變而受不同程度之影響。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註10:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">高收益債券基金之投資標的涵蓋低於投資等級之非投資等級債券，故需承受較大之價格波動，而利率風險、信用違約風險、外匯波動風險也將高於一般投資等級之債券。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註11:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">投資人投資高收益債券基金不宜占其投資組合過高之比重，本基金經行政院金融監督管理委員會核准，惟不表示絕無風險。由於高收益債券之信用評等未達投資等級或未經信用評等，且對利率變動的敏感度甚高，故此類基金可能會因利率上升、市場流動性下降，或債券發行機構違約不支付本金、利息或破產而蒙受虧損，此類基金不適合無法承擔相關風險之投資人。基金經理公司以往之經理績效不保證基金之最低投資收益；基金經理公司除盡善良管理人之注意義務外，不負責各基金之盈虧，亦不保證最低之收益，投資人申購前應詳閱基金公開說明書。此外，部分高收益債券基金可能投資於符合美國Rule 144A規定具有私募性質之債券，由於美國Rule144 A債券僅限機構投資人購買，資訊揭露要求較一般債券寬鬆，於次級市場交易時可能因參與者較少，或交易對手出價意願較低，導致產生較大的買賣價差，進而影響基金淨值。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註12:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">委託人委託投資之基金若歸類為高收益債券型基金，委託人暸解此等類型基金係主要投資於非投資等級之高風險債券，其投資風險主要來自於所投資債券標的之利率及信用風險。債券價格與利率係為反向關係，當市場利率上調時將導致債券價格下跌，產生利率風險；此外，投資於高收益債券亦可能隱含債券發行主體無法償付本息之信用風險。故當高收益債券型基金之投資標的發生上開利率或信用風險事件時，其淨資產價值亦將因此而產生波動。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註13:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">新興市場債券基金之投資標的包含政治、經濟相對較不穩定之新興市場國家之債券，因此將面臨較高的政治、經濟變動風險、利率風險、債信風險與外匯波動風險。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註14:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">基金之主要投資風險除包含一般固定收益產品之利率風險、流動風險、匯率風險、信用或違約風險外，由於此類基金有投資部份的新興國家債券，而新興國家的債信等級普遍較已開發國家為低，所以承受的信用風險也相對較高，尤其當新興國家經濟基本面與政治狀況變動時，均可能影響其償債能力與債券信用品質。基金投資均涉及風險且不負任何抵抗投資虧損之擔保。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註15:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">基金的配息可能由基金的收益或本金中支付。任何涉及由本金支出的部份，可能導致原始投資金額減損。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">註16:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">上述資料只供參考用途，嘉實資訊自當盡力提供正確訊息，但如有錯漏或疏忽，本公司或關係企業與其任何董事或任何受僱人，恕不負任何法律責任。</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "</table>" & vbcrlf
	
	FundMsgText = xxx
end function

function printf(p_val)
	Response.Write "#" & p_val & "#"
end function

Function FormatYMD(d)
	Dim xxx, mm, dd
	
	On Error Resume Next
	mm = Month(CDate(d))
	If CInt(mm) < 10 Then
		mm = "0" & mm
	End If
	dd = Day(CDate(d))
	If CInt(dd) < 10 Then
		dd = "0" & dd
	End If		
	if CInt(Year(CDate(d)) - 1911) <= 0 then
		xxx = "N/A"
	else
		xxx = (Year(CDate(d)) - 1911) & "/" & CStr(mm) & "/" & CStr(dd)		
	end if		
	If err.number Then
		xxx = "N/A"
	End If
	FormatYMD = CStr(xxx)
End Function

Function FormatYYMD(d)
	Dim xxx, mm, dd
	
	On Error Resume Next
	mm = Month(CDate(d))
	If CInt(mm) < 10 Then
		mm = "0" & mm
	End If
	dd = Day(CDate(d))
	If CInt(dd) < 10 Then
		dd = "0" & dd
	End If		
	if CInt(Year(CDate(d))) <= 0 then
		xxx = "N/A"
	else
		xxx = (Year(CDate(d))) & "/" & CStr(mm) & "/" & CStr(dd)		
	end if		
	If err.number Then
		xxx = "N/A"
	End If
	FormatYYMD = CStr(xxx)
End Function
Function FormatYMDT(d)
	dim xxx, hh, mm, ss
	
	xxx = FormatYMD(d)
	if xxx = "N/A" then
		FormatYMDT = xxx
		exit function
	end if
	hh = Hour(CDate(d))
	if hh < 10 then
		hh = "0" & CStr(hh)
	end if
	mm = Minute(CDate(d))
	if mm < 10 then
		mm = "0" & CStr(mm)
	end if
	ss = Second(CDate(d))
	if ss < 10 then
		ss = "0" & CStr(ss)
	end if
	FormatYMDT = xxx & " " & CStr(hh) & ":" & CStr(mm) & ":" & CStr(ss)
End Function

Function FormatExpT(d)
	Dim xxx, mm, dd, h, m, s
	
	On Error Resume Next
	mm = Month(CDate(d))
	If CInt(mm) < 10 Then
		mm = "0" & mm
	End If
	dd = Day(CDate(d))
	If CInt(dd) < 10 Then
		dd = "0" & dd
	End If
	xxx = Year(CDate(d)) & "/" & CStr(mm) & "/" & CStr(dd)
	h = Hour(CDate(d))
	if h < 10 then
		h = "0" & CStr(h)
	end if
	m = Minute(CDate(d))
	if m < 10 then
		m = "0" & CStr(m)
	end if
	s = Second(CDate(d))
	if s < 10 then
		s = "0" & CStr(s)
	end if
	xxx = xxx & " " & CStr(h) & ":" & CStr(m) & ":" & CStr(s)
	If err.number Then
		xxx = ""
	End If
	FormatExpT = CStr(xxx)
End Function

Function FormatExpT2(d)
Dim xxx, mm, dd, h, m, s

	On Error Resume Next
	mm = Month(CDate(d))
	If CInt(mm) < 10 Then
		mm = "0" & mm
	End If
	dd = Day(CDate(d))
	If CInt(dd) < 10 Then
		dd = "0" & dd
	End If
	xxx = Year(CDate(d)) & "/" & CStr(mm) & "/" & CStr(dd)
	h = Hour(CDate(d))
	if h < 10 then
		h = "0" & CStr(h)
	end if
	m = Minute(CDate(d))
	if m < 10 then
		m = "0" & CStr(m)
	end if
	'不顯示秒數
	's = Second(CDate(d))
	'if s < 10 then
	's = "0" & CStr(s)
	'end if
	'xxx = xxx & " " & CStr(h) & ":" & CStr(m) & ":" & CStr(s)
	xxx = xxx & " " & CStr(h) & ":" & CStr(m)
	If err.number Then
		xxx = ""
	End If
	FormatExpT2 = CStr(xxx)
End Function


function getWFund_CIDName(cid)
	dim sql, fname
	fname = ""  
	sql = "exec wsp_get_cid_info '" & cid & "'"
	if OpenWFundDB(conn, rs, sql) then
		fname = replace(trim(rs(1))," ","")
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if
	getWFund_CIDName = fname
end function

function getWFund_FIDName(fid, cid)
	dim sql, fname 
	cid = "NA"
	fname = ""
	sql = "exec wsp_get_fid_name '" & fid & "'"
	if OpenWFundDB(conn, rs, sql) then
		cid = trim(rs(2))
		fname = replace(trim(rs(1))," ","")
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if
	getWFund_FIDName = fname
end function

function GetDocProlog(title,pid,fid,cid)
	dim xxx
	xxx = ""
	xxx = xxx & "<html>" & chr(13) & chr(13)
	xxx = xxx & "<head>" & chr(13) & chr(13)
	xxx = xxx & "<script language=""JavaScript"" src=""/w/js/wfundjs.djjs""></script>" & chr(13) & chr(10)
	xxx = xxx & "<script language=""JavaScript"" src=""/w/js/wfund.js""></script>" & chr(13) & chr(10)
	if pid = "wb01" then
		xxx = xxx & "<script type=""text/JavaScript"" src=""/include/js/publicFunction.js""></script>" & chr(13) & chr(10)
		xxx = xxx & "<link href=""/include/css/global.css"" rel=""stylesheet"" type=""text/css""/>" & chr(13) & chr(10)
		xxx = xxx & "<link href=""/include/css/nav.css"" rel=""stylesheet"" type=""text/css""/>" & chr(13) & chr(10)
		xxx = xxx & "<link href=""/include/css/fund/fund.css"" rel=""stylesheet"" type=""text/css""/>" & chr(13) & chr(10)
		xxx = xxx & "<link href=""/include/css/fund/market_info/basic_info/basic_info.css"" rel=""stylesheet"" type=""text/css""/>" & chr(13) & chr(10)
		xxx = xxx & "<link href=""/include/css/article.css"" rel=""stylesheet"" type=""text/css""/>" & chr(13) & chr(10)
		xxx = xxx & "<link href=""/include/css/article_colour.css"" rel=""stylesheet"" type=""text/css"" />" & chr(13) & chr(10)
		'xxx = xxx & "<link href=""/include/css/tab_table.css"" rel=""stylesheet"" type=""text/css""/>" & chr(13) & chr(10)
	end if
	xxx = xxx & "<script language=""JavaScript""><!--" & chr(13) & chr(10)
	xxx = xxx & "CheckMenu('" & lcase(trim(pid)) & "','" & ucase(trim(fid)) & "','" & ucase(trim(cid)) & "');" & chr(13) & chr(10)
	xxx = xxx & "if (navigator.appName.indexOf('Netscape') != -1) {" & chr(13) & chr(10)
	xxx = xxx & "document.writeln('<link rel=stylesheet href=""/w/js/wFundNS.css"" type=""text/css"">\n');" & chr(13) & chr(10)
	xxx = xxx & "} else {" & chr(13) & chr(10)
	xxx = xxx & "document.writeln('<link rel=stylesheet href=""/w/js/wFund.css"" type=""text/css"">\n');" & chr(13) & chr(10)
	xxx = xxx & "}" & chr(13) & chr(10)
	xxx = xxx & "// --></script>" & chr(13) & chr(10)
	xxx = xxx & "<meta HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=big5"">" & chr(13) & chr(10)
	xxx = xxx & "<meta HTTP-EQUIV=""PRAGMA"" CONTENT=""NO-CACHE"">" & chr(13) & chr(10)
	xxx = xxx & "<title>海外基金" & title & "</title>" & chr(13) & chr(10)
	xxx = xxx & "</head>" & chr(13) & chr(10)
	xxx = xxx & "<body>" & chr(13) & chr(10)
	xxx = xxx & "<div align=left id=""SysJustIFRAMEDIV"">" & chr(13) & chr(10)
	xxx = xxx & "<script language=""javascript"" src=""/w/js/SU.js""></script>" & chr(13) & chr(10)
	'xxx = xxx & "<div align=center>" & chr(13) & chr(10)
	GetDocProlog = xxx
end function

function GetDocEplog(p_dataflag)
	dim xxx,x
	xxx = ""
	xxx = xxx & "<script language=""javascript"" src=""/w/js/SD.js""></script>" & chr(13) & chr(10)
	xxx = xxx & "</div></body></html>" & chr(13) & chr(10)
	xxx = xxx & "<!--"  & FormatExpT(CalcExpTime(p_dataflag)) & "-->"
	GetDocEplog = xxx
	
	x = "<!--" & FormatExpT(CalcExpTime(p_dataflag)) & "-->"
	Response.AddHeader "DJ_Expired", x	
end function

function datechg(tt)
	if tt="NA" then
		tt=""
	else
		tt=replace(tt,"-","/")
		tt=cdate(tt)
		'tt=year(tt)-1911 & "/" & month(tt) & "/" & day(tt)
		tt=year(tt) & "/" & month(tt) & "/" & day(tt)
	end if
	datechg=tt
end function

'取得顏色區塊
Function GetColor(CIndex)
	Select case CIndex
	  case 0
	    GetColor="ccffff"
	  case 1
	    GetColor="ccccff"
	  case 2
	    GetColor="ffccff"
	  case 3
	    GetColor="ccffcc"
	  case 4
	    GetColor="33cccc"
	  case 5
	    GetColor="ffcc99"
	  case 6
	    GetColor="9999ff"
	  case 7
	    GetColor="ff6666"
	  case 8
	    GetColor="66ff66"
	  case 9
	    GetColor="cc66cc"
	  case 10
	    GetColor="00cc99"
	  case 11
	    GetColor="00ffff"
	  case 12
	    GetColor="0000ff"
	  case 13
	    GetColor="ff0033"
	  case 14
	    GetColor="ffff00"
	  case 15
	    GetColor="bb3399"
	  case 16
	    GetColor="bb9933"
	  case 17
	    GetColor="990099"
	  case 18
	    GetColor="006600"
	  case 19
	    GetColor="993333"
	  case 20
	    GetColor="8fbc8b"
	  case 21
	    GetColor="9966ff"
	  case 100
	    GetColor="b0e0e6"
	  case 101
	    GetColor="ff0000"
	  case 102
	    GetColor="000000"
	  case else
	    GetColor="000000"
	End Select
End Function

'新版的converJS 有改js中的link
Function ConverJS( strContent)
	Response.Write "document.writeln('" & Replace(strContent, "'", "\'") & "\n');" & chr(13) & chr(10)
end Function

Function MakeButton(rno,bankFundID)
	dim bbb
'	if make_btn = "y" then
		bbb = "<input type=""button"" value=""申購"" name=B" & rno & " onclick=""goBuyFund('" & bankfundid & "')"">"
'	else
'		bbb = ""
'	end if
	MakeButton = bbb
End Function

Function Wf09_Msg(idx)
	Dim xx : xx = ""
	
	select case idx
		case 1	'wf09a
			xx = xx & "註1:績效計算為原幣別報酬，且皆有考慮配息情況。基金配息率不代表基金報酬率，且過去配息率不代表未來配息率。所有基金績效，均為過去績效，不代表未來之績效表現，亦不保證基金之最低投資收益。<BR>" & vbcrlf
			xx = xx & "註2:基金淨值可能因市場因素而上下波動，基金淨值僅供參考，實際以基金公司公告之淨值為準；海外市場指數類型基金，以交易日當天收盤價為淨值參考價。<BR>" & vbcrlf
			xx = xx & "註3:有關基金應負擔之費用（含分銷費用）已揭露於基金之公開說明書及投資人須知中，投資人可至境外基金資訊觀測站(http://www.fundclear.com.tw)下載，或逕向總代理人網站查閱。<BR>" & vbcrlf
			xx = xx & "註4:標示""(已撤銷核備)""及""(未申報生效)""之基金資料僅提供原有投資者參考，以供其做買回、轉換或繼續持有之決定；「已撤銷核備」為在台灣已下架 (停售) 的基金；「未申報生效」為未於法規規定之時效內向金管會辦理申報生效作業之境外基金。<BR>" & vbcrlf
			xx = xx & "註5:境外基金經行政院金融監督管理委員會核准或申報生效在國內募集及銷售，惟不表示絕無風險。基金經理公司以往之經理績效不保證基金之最低投資收益；基金經理公司除盡善良管理人之注意義務外，不負責本基金之盈虧，亦不保證最低之收益，投資人申購前應詳閱基金公開說明書。<BR>" & vbcrlf
			xx = xx & "註6:上述資料只供參考用途，基智網自當盡力提供正確訊息，但如有錯漏或疏忽，本公司或關係企業與其任何董事或任何受僱人，恕不負任何法律責任。<BR>" & vbcrlf
			xx = xx & "註7:在國內經證期局核備之海外基金，挑選出全球型基金、區域型(含單一國家)和特殊類(包括能源、貴金屬、醫療)三類基金。<BR>" & vbcrlf
			xx = xx & "註8:依一年報酬率，找出全球型前10檔基金、區域型前5檔基金，和特殊類前5檔基金，將這20支基金依報酬率及標準差散佈在X Y圖上。<BR>" & vbcrlf
		case 2	'wf09b
			xx = xx & "註1:基金績效計算皆有考慮配息，基金配息率不代表基金報酬率，且過去配息率不代表未來配息率。所有基金績效，均為過去績效，不代表未來之績效表現，亦不保證基金之最低投資收益。<BR>" & vbcrlf
			xx = xx & "註2:基金淨值僅供參考，實際以基金公司公告之淨值為準。<BR>" & vbcrlf
			xx = xx & "註3:境內基金經行政院金融監督管理委員會核准在國內募集及銷售，惟不表示絕無風險。基金經理公司以往之經理績效不保證基金之最低投資收益；基金經理公司除盡善良管理人之注意義務外，不負責本基金之盈虧，亦不保證最低之收益，投資人申購前應詳閱基金公開說明書。<BR>" & vbcrlf
			xx = xx & "註4:上述資料只供參考用途，基智網(FundDJ)自當盡力提供正確訊息，但如有錯漏或疏忽，本公司或關係企業與其任何董事或任何受僱人，恕不負任何法律責任。<BR>" & vbcrlf
			xx = xx & "註5:所謂國內股票型基金意指投信發行，投資標的為國內股票之基金(不包含國外募集投資國內)。<BR>" & vbcrlf
			xx = xx & "註6:在所有基金中挑選SHARPE值前20名，並將這20支基金依報酬率及標準差散佈在X Y圖上，在X Y圖上畫出國內股票型基金之平均報酬率與平均標準差，以兩平均線之交點為原點，第二與第三象限適合保守型投資人；第一與第二象限適合積極型投資人。若有基金不在上述規則之內，將此類基金放入其他推薦基金。<BR>" & vbcrlf
		case 3	'wf09c
			xx = xx & "註1:基金績效計算皆有考慮配息，基金配息率不代表基金報酬率，且過去配息率不代表未來配息率。所有基金績效，均為過去績效，不代表未來之績效表現，亦不保證基金之最低投資收益。<BR>" & vbcrlf
			xx = xx & "註2:基金淨值僅供參考，實際以基金公司公告之淨值為準。<BR>" & vbcrlf
			xx = xx & "註3:境內基金經行政院金融監督管理委員會核准在國內募集及銷售，惟不表示絕無風險。基金經理公司以往之經理績效不保證基金之最低投資收益；基金經理公司除盡善良管理人之注意義務外，不負責本基金之盈虧，亦不保證最低之收益，投資人申購前應詳閱基金公開說明書。<BR>" & vbcrlf
			xx = xx & "註4:上述資料只供參考用途，基智網(FundDJ)自當盡力提供正確訊息，但如有錯漏或疏忽，本公司或關係企業與其任何董事或任何受僱人，恕不負任何法律責任。<BR>" & vbcrlf
			xx = xx & "註5:所謂國內債券型基金包含有買回期限限制類與無買回期限限制類。<BR>" & vbcrlf
			xx = xx & "註6:在所有基金中挑選SHARPE值前20名，並將這20支基金依報酬率及標準差散佈在X Y圖上，在X Y圖上畫出國內股票型基金之平均報酬率與平均標準差，以兩平均線之交點為原點，第二與第三象限適合保守型投資人；第一與第二象限適合積極型投資人。若有基金不在上述規則之內，將此類基金放入其他推薦基金。<BR>" & vbcrlf
		case 4	'wf09d
			xx = xx & "註1:績效計算為原幣別報酬，且皆有考慮配息情況。基金配息率不代表基金報酬率，且過去配息率不代表未來配息率。所有基金績效，均為過去績效，不代表未來之績效表現，亦不保證基金之最低投資收益。<BR>" & vbcrlf
			xx = xx & "註2:基金淨值可能因市場因素而上下波動，基金淨值僅供參考，實際以基金公司公告之淨值為準；海外市場指數類型基金，以交易日當天收盤價為淨值參考價。<BR>" & vbcrlf
			xx = xx & "註3:有關基金應負擔之費用（含分銷費用）已揭露於基金之公開說明書及投資人須知中，投資人可至境外基金資訊觀測站(http://www.fundclear.com.tw)下載，或逕向總代理人網站查閱。<BR>" & vbcrlf
			xx = xx & "註4:標示""(已撤銷核備)""及""(未申報生效)""之基金資料僅提供原有投資者參考，以供其做買回、轉換或繼續持有之決定；「已撤銷核備」為在台灣已下架 (停售) 的基金；「未申報生效」為未於法規規定之時效內向金管會辦理申報生效作業之境外基金。<BR>" & vbcrlf
			xx = xx & "註5:境外基金經行政院金融監督管理委員會核准或申報生效在國內募集及銷售，惟不表示絕無風險。基金經理公司以往之經理績效不保證基金之最低投資收益；基金經理公司除盡善良管理人之注意義務外，不負責本基金之盈虧，亦不保證最低之收益，投資人申購前應詳閱基金公開說明書。<BR>" & vbcrlf
			xx = xx & "註6:上述資料只供參考用途，基智網自當盡力提供正確訊息，但如有錯漏或疏忽，本公司或關係企業與其任何董事或任何受僱人，恕不負任何法律責任。<BR>" & vbcrlf
			xx = xx & "註7:海外債券為在國內經證期局核備之海外債券型基金，包括債券型、高收益債、短期債券及可轉換債券等基金。<BR>" & vbcrlf
			xx = xx & "註8:在所有基金中挑選SHARPE值前20名，並將這20支基金依報酬率及標準差散佈在X Y圖上，在X Y圖上畫出海外債券型基金之平均報酬率與平均標準差，以兩平均線之交點為原點，第二與第三象限適合保守型投資人；第一與第二象限適合積極型投資人。若有基金不在上述規則之內，將此類基金放入其他推薦基金。<BR>" & vbcrlf
	end select
	
	Wf09_Msg = xx
end Function


'傳回基金風險收益等級.
function GetRiskLevel(sFundID)
	Dim sSQL ,rs,conn
	GetRiskLevel = "N/A"
	sSQL = "exec spj_mda70164 '" & sFundID & "'"
	
	if OpenFundDJ(conn,rs,sSQL) then
		if Not isnull(trim(rs(1))) Then
			if trim(rs(1)) & "" = "" Then
				GetRiskLevel = "N/A"
			else
				GetRiskLevel = trim(rs(1))
			end if
		end if
		
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if
end function

'取得該銀行銷售的基金是否可以申購 : approved field
Function GetFundApproved(sBankFundID)
	Dim SQLstr : SQLstr = ""
	Dim conn,rs
	
	SQLstr = SQLstr & " select approved from waspprofile "
	SQLstr = SQLstr & " where aspid='" & FUNDASPID & "' and bankfundid='" & trim(sBankFundID) & "'"
	'Response.Write SQLstr & "<BR>"
	if OpenWFundDB(conn,rs,SQLstr) then
		tmp = trim(rs(0))
		if isnull(tmp) or tmp = "" then
			tmp = ""
		else
			tmp = ucase(tmp)
		end if
	
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if	
	
	GetFundApproved = tmp
end Function


'-- 海外單一基金的 二層式基金下拉選單 相關 Script --
Function GenComboList(CID,FID,sFormName)
	Dim xx : xx = ""
	xx = xx & "<SCRIPT LANGUAGE=javascript><!--" & vbCrLf
	'xx = xx & " alert(cuteduck); " & vbcrlf
	xx = xx & " var sObj = eval('document.' + '" & sFormName & "'); " & vbcrlf
	xx = xx & "	iID = '" & FID & "';" & vbCrLf
	xx = xx & "	GenFundCorpCombo('" & CID & "','" & FID & "','" & sFormName & "');" & vbCrLf

	xx = xx & "	for (i=0;i<sObj.selFund_corp.options.length;i++)" & vbCrLf
	xx = xx & "	{" & vbCrLf
	xx = xx & "		var tmpID1 = sObj.selFund_corp.options[i].value.toUpperCase();" & vbCrLf
	xx = xx & "		if (tmpID1 == '" & CID & "') " & vbCrLf
	xx = xx & "		{" & vbCrLf
	xx = xx & "			sObj.selFund_corp.selectedIndex = i;" & vbCrLf
	xx = xx & "			break;" & vbCrLf
	xx = xx & "		}" & vbCrLf
	xx = xx & "	}" & vbCrLf
	
	xx = xx & "	for (i=0;i<sObj.selFund3.options.length;i++)" & vbCrLf
	xx = xx & "	{" & vbCrLf
	xx = xx & "		var tmpID2 = sObj.selFund3.options[i].value.toUpperCase();" & vbCrLf
	xx = xx & "		if (iID != '')" & vbCrLf
	xx = xx & "		{" & vbCrLf
	xx = xx & "			if (tmpID2 == iID )" & vbCrLf
	xx = xx & "			{" & vbCrLf
	xx = xx & "				sObj.selFund3.selectedIndex = i;" & vbCrLf
	xx = xx & "				break;" & vbCrLf
	xx = xx & "			}" & vbCrLf
	xx = xx & "		}" & vbCrLf
	xx = xx & "		else" & vbCrLf
	xx = xx & "			sObj.selFund3.selectedIndex = 0;" & vbCrLf
	xx = xx & "	}" & vbCrLf
	xx = xx & "//--></SCRIPT>" & vbCrLf
	
	GenComboList = xx
end Function

%>