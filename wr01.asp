<!-- #include file="wtFundProc.asp" -->
<%
fid = ucase(trim(request("A")))
if fid="" or fid="NA" then fid = deftFID

idname = getFundDJ_IDName(fid)

'=====================================  �򥻸�Ƴ��S����ƨ��o (wr01)   ================================================
Dim s801,s300,s310,s320,s330,s340,fname,s120

sql = "exec tsp_get_fund_info_800 '" & fid & "'"

if OpenFundDJ(conn, rs, sql) then
	dot = 2
	dcg = 6
	getDecimal rs("yb800120"), dot, dcg
	
	s300 = stdfmt(rs("yb800300"),dot) & " (" & trim(rs("yb800010")) & ")"     '�̪�b��
	s330 = stdfmt(rs("yb800330"),2)       '���ߤ�_���S�v
	s340 = stdfmt(rs("yb800340"),2)       '���~�H�ӳ��S�v
	s801 = stdfmt(rs("yb800801"),2)       '�@�Ӥ���S�v
	s310 = stdfmt(rs("yb800310"),2)       '�P�Q����
	s320 = stdfmt(rs("yb800320"),2)       '���^�T
	fname = trim(rs("fundname"))
	cid =  trim(rs("yb800140"))
	
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
end if

'�̪�@���t�����
sql = "exec YS130000 '" & fid & " '"
if OpenFundDJ(conn, rs, sql) then
	s120 = trim(rs(1)) & " (" & trim(rs(0)) & ")"
end if
if s120 = "" then
	's120 = "�S���t�����"
	s120 = "N/A"
end if

'============================================  ���Ѥ�ҭp�� (wr04)   ====================================================
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
	dated=FormatYMD(trim(rs("�������")))
	do While Not rs.EOF 
		if ucase(trim(rs(1))) <> "EBTOTAL" and  ucase(trim(rs(1)))<> "EBZAA" then
			sumd = sumd +1
			MyCounter = MyCounter +1
			if left(ucase(trim(rs(1))),3)="EB0" then 
				T = T & " " & "�W��" & replace(replace(trim(rs(5)),"��","")," ","_") 
			elseif left(ucase(trim(rs(1))),3)="EB1" then 
				T = T & " " & "�W�d" & replace(replace(trim(rs(5)),"��","")," ","_")
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
		
		T = mid(T,2) & " �䥦"
		
		V = mid(V,2) & " " & formatnumber(MyTotal*sitebase,0,0,0,0)

		I = mid(I,2) & " " & formatnumber(MyTotal2) 
		C = mid(C,2) & " " & GetColor(MyCounter+1)
		S = S & " 0"
	end if
end if


'====================================================================================================================
Response.Write GetDocProlog("����򥻸��", "wr01", fid, "NA", "NA")
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
xxxIDS = xxxIDS & "<option selected " & "value=""/w/wr/wr01_" & fid & ".djhtm"">����򥻸��</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr02_" & fid & ".djhtm"">����b�Ȫ�</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr03_" & fid & ".djhtm"">����Z�Ī�</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr04_" & fid & ".djhtm"">������Ѫ��p</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr10_" & fid & ".djhtm"">����t�����p</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "<option " & "value=""/w/wr/wr05_" & fid & ".djhtm"">��������s�D</option>" & chr(13) & chr(10)
xxxIDS = xxxIDS & "</select>" & chr(13) & chr(10)


area = ""
sql = "exec tsp_get_fund_area '" & fid & "'"
if OpenFundDJ(conn, rs, sql) then
	while not rs.EOF
		if area <> "" then area = area & "�A"
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
		if sale <> "" then sale = sale & "�A"
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
		if sss <> "" then sss = sss & "�B"
		sss = sss  & trim(rs(1)) & "��"
		rs.movenext
	wend
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
	sss = sss & "&nbsp;(��ڦ��ڤ�̥��ʤH�P�P����c���w)"
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
'		Response.Write "<OPTION>����������</OPTION>" & chr(13) & chr(10)
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

	
	'============================================   �򥻸�� (wr01)   ====================================================

	Response.Write "<div class=""a_tab_block tab_btm_article"">" & chr(13) & chr(10)
	Response.Write "  <div class=""aj_block tab_top_article sbg""></div>" & chr(13) & chr(10)
	Response.Write "  <div class=""squeeze"">" & chr(13) & chr(10)
	Response.Write "    <div class=""a_tab_block_title"">" & chr(13) & chr(10)
	Response.Write "      <h5>" & fname & "</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	

	response.write "    <table>" & chr(13) & chr(10)
	
	response.write "      <tr class=""pink"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">������q</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%""><span class=""bear""><a href=""/w/wp/wp01_" & trim(rs(2)) & ".djhtm"">" & trim(rs(3)) & "</a></span></td>"  & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""17%"">�������</td>" & chr(13) & chr(10)
	response.write "        <td colspan=3 width=""46%""><span class=""bear""><a href=""/w/wr/wr00A_" & trim(rs(4)) & ".djhtm"">" & trim(rs(5)) & "</a></span></td>" & chr(13) & chr(10)
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""pink"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">���ߤ��</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%"">" & stdfmt(rs(8),5) & "</td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""17%"">���߮ɳW��(��)</td>" & chr(13) & chr(10)
	response.write "        <td width=""13%"">" & stdfmt(rs(10),2) & "</td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""16%"">����W��(��)</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%"">" & stdfmt(rs(9),2) & " (" & trim(rs("yb100311")) & ")</td>" & chr(13) & chr(10)
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""pink"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">����g�z�H</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%""><a href=""/w/wt/wt03_" & trim(rs(6)) & ".djhtm"">" & trim(rs(7)) & "</a></td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""17%"">���ϰ�</td>" & chr(13) & chr(10)
	response.write "        <td width=""13%"">" & area & "&nbsp;</td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""16%"">�p�����O</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%"">" & stdfmt(rs(16),0) & "</td>" & chr(13) & chr(10)
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""gray"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">�@�Ӥ���S�v</td>" & chr(13) & chr(10)

	if isnull(s801) or trim(s801) = "" or trim(s801) = "&nbsp;" then
		response.write "        <td width=""17%"">0.00%</span></td>" & chr(13) & chr(10)
	else
		if s801 < 0 then
			response.write "        <td width=""17%""><span class=""fall"">" & s801 & "%</span></td>" & chr(13) & chr(10)
		else
			response.write "        <td width=""17%"">" & s801 & "%</td>" & chr(13) & chr(10)
		end if
	end if
	response.write "        <td class=""col_head"" width=""17%"">�̷s�b��</td>" & chr(13) & chr(10)
	response.write "        <td colspan=3 width=""46%"">" & s300 & "</td>" & chr(13) & chr(10)
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""gray"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">���ߤ�_���S�v</td>" & chr(13) & chr(10)
	if s330 < 0 then
		response.write "        <td width=""17%""><span class=""fall"">" & s330 & "%</span></td>" & chr(13) & chr(10)
	else
		response.write "        <td width=""17%"">" & s330 & "%</td>" & chr(13) & chr(10)
	end if
	response.write "        <td class=""col_head"" width=""17%"">�P�Q����</td>" & chr(13) & chr(10)
	if s310 > 0 then
		response.write "        <td width=""13%"">" & s310 & "��</td>" & chr(13) & chr(10)
	elseif s310 < 0 then
		response.write "        <td width=""13%""><span class=""fall"">" & s310 & "��</span></td>" & chr(13) & chr(10)
	else
		response.write "        <td width=""13%"">" & s310 & "</td>" & chr(13) & chr(10)
	end if
	response.write "        <td class=""col_head"" width=""16%"">���^�T</td>" & chr(13) & chr(10)
	if s320 < 0 then
		response.write "        <td width=""17%""><span class=""fall"">" & s320 & "%</span></td>" & chr(13) & chr(10)
	else
		response.write "        <td width=""17%"">" & s320 & "%</td>" & chr(13) & chr(10)
	end if
	response.write "      </tr>" & chr(13) & chr(10)
	
	response.write "      <tr class=""gray"">" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""20%"">���~�H�ӳ��S�v</td>" & chr(13) & chr(10)
	if s340 < 0 then
		response.write "        <td width=""17%""><span class=""fall"">" & s340 & "%</span></td>" & chr(13) & chr(10)
	else
		response.write "        <td width=""17%"">" & s340 & "%</td>" & chr(13) & chr(10)
	end if
	response.write "        <td class=""col_head"" nowrap width=""17%""><a href=http://www.funddj.com/y/notes/rrnotes/rrnotes.htm target=_blank>���I���q����</a></td>" & chr(13) & chr(10)
	response.write "        <td width=""13%""><span class=""bear"">" & GetRiskLevel(fid) & "</span></td>" & chr(13) & chr(10)
	response.write "        <td class=""col_head"" width=""16%"">�̪�t��</td>" & chr(13) & chr(10)
	response.write "        <td width=""17%"">" & s120 & "</td>" & chr(13) & chr(10)	
	response.write "      </tr>" & chr(13) & chr(10)
	Response.Write "    </table>" & chr(13) & chr(10)
	
	
	'���U�� : 2008/10/09 modified
	Response.Write "    <div align=""left"">" & vbcrlf
	Response.Write "    <ul class=""cate_addinfo_line"">" & vbcrlf
	Response.Write "      <li><em>���U���G</em></li>" & vbcrlf
	Response.Write "      <li><a href=""" & GetTWFundInfoURL(fid,"2") & """ target=""_blank"" title=""�]�ȳ��i��"">�]�ȳ��i��</a></li>" & vbcrlf
	Response.Write "      <li><a href=""" & GetTWFundInfoURL(fid,"1") & """ target=""_blank"" title=""���}������"">���}������</a></li>" & vbcrlf
	Response.Write GetFundEasyReport(fid)
	Response.Write "      <li><a href=""#""  onclick=""javascript:return sOpenTradeRule('" & fid & "');"">�u�u����W�w</a></li>" & vbcrlf
    Response.Write "	  <li><a class=""end"" href="""& GetFundMonthReport(fid,int(Timer))&""" target=""_blank"">������</a></li>" & vbcrlf

	Response.Write "    </ul>" & vbcrlf
	Response.Write "    </div>" & chr(13) & chr(10)
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)

	
	rs.close
	conn.close
	set rs = nothing
	set conn = nothing
	
		
	'�ӫ~DM
	'Response.Write "<ul class=""dmlink"">" & chr(13) & chr(10)
	'Response.Write "  <li class=""prod_dm""><a href=""#"" title=""�ӫ~DM"">�ӫ~DM</a></li>" & chr(13) & chr(10)
	'Response.Write "  <li class=""more""><a href=""#"" title=""MORE"">MORE</a></li>" & chr(13) & chr(10)
	'Response.Write "</ul>" & chr(13) & chr(10)
	
	Response.Write "    <h4>&nbsp;</h4>" & chr(13) & chr(10)
	'====================================================================================================================
	
	
	'============================================  ����b�� (wr02)   ====================================================

	Response.Write "<div class=""article_inline_block pusher"">" & chr(13) & chr(10)
	Response.Write "  <div class=""squeezer"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h5>����b�Ȩ��չ�</h5>" & chr(13) & chr(10)
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
	'Response.write "          <param name=""T"" value=""" & idname & "����b�Ȩ��չ�"">" & chr(10)
	Response.write "          <param name=""T"" value="""">" & chr(10)
	Response.write "          <param name=""U"" value=""��(" & currencyType & ")"">" & chr(10)
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
		Response.Write "<tr><td>�L�b�ȸ��</td></tr>" & chr(13) & chr(10)
		Response.Write "</table>" & chr(13) & chr(10)
	end if
	
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)
	'====================================================================================================================
	

	'============================================  ����Z�� (wr03)   ====================================================
	
	Response.Write "<div class=""article_inline_block"">" & chr(13) & chr(10)
	Response.Write "  <div class=""squeezer"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h5>����Z�Ķչ�</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	
	Response.Write "    <table>" & chr(13) & chr(10)
	Response.Write "      <tr><td>" & chr(13) & chr(10)
	Response.Write "        <applet archive=""MCURVE5.jar"" codebase=""/w/jar"" CODE=""MCURVE5.class"" NAME=MCURVE5 HEIGHT=182 WIDTH=320 VIEWASTEXT id=MCURVE5>" & chr(13) & chr(10)
	Response.Write "          <param name=""BCD"" value=""/w/bcd/BCDROIList5_" & fid & "_NA_NA_NA_NA_NA_NA_1.djbcd"">" & chr(13) & chr(10)
	Response.Write "          <param name=""CAPTION"" value="""">" & chr(13) & chr(10)
	response.Write "          <param name=""BC"" value=""fff8e5"">" & chr(13) & chr(10)
	response.Write "          <param name=""LC"" value=""000000 00AAAA AAAA00 AA00AA 0000AA"">" & chr(13) & chr(10)
	response.Write "          <param name=""T"" value=""" & idname & """>" & chr(13) & chr(10)
	response.Write "          <param name=""U"" value=""�@ �@ �@ �@ �@"">" & chr(13) & chr(10)
	response.Write "        </applet>" & chr(13) & chr(10)
	Response.Write "      </td></tr>" & chr(13) & chr(10)
	Response.Write "    </table>" & chr(13) & chr(10)
	
	Response.Write "    <table>" & chr(13) & chr(10)
	Response.Write "      <thead>" & chr(13) & chr(10)
	Response.Write "        <tr>" & chr(13) & chr(10)
	Response.Write "          <td>�@�Ӥ�</td>" & chr(13) & chr(10)
	Response.Write "          <td>�T�Ӥ�</td>" & chr(13) & chr(10)
	Response.Write "          <td>���Ӥ�</td>" & chr(13) & chr(10)
	Response.Write "          <td>�~�ƼзǮt</td>" & chr(13) & chr(10)
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

	'============================================  ���Ѥ�� (wr04)   ====================================================

	Response.Write "<div class=""a_tab_block tab_btm_article holdpercent cleartitle"">" & chr(13) & chr(10)
	Response.Write "  <div class=""aj_block tab_top_article sbg""></div>" & chr(13) & chr(10)
	Response.Write "  <div class=""squeeze"">" & chr(13) & chr(10)
	Response.Write "    <div class=""a_tab_block_title"">" & chr(13) & chr(10)
	Response.Write "      <h5>���Ѥ��</h5>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	Response.Write "    <div class=""block_composer""> " & chr(13) & chr(10)
	
	
	if sitebase <> 0 and showFlag then
		Response.Write "      <table>" & chr(13) & chr(10)
		Response.Write "        <tr><td valign=""top"">" & chr(13) & chr(10)
		Response.Write "          <applet ARCHIVE=""PIE2DNoTable.JAR"" CODE=""PIE2DNoTable.class"" codebase=""/w/jar"" width=220 height=187px VIEWASTEXT id=Applet1>" & chr(13) & chr(10)
		Response.Write "            <param name=""T"" value=""������Ѥ��G " & T & """>" & chr(13) & chr(10)
		Response.Write "            <param name=""V"" value=""" & V & """>" & chr(13) & chr(10)
		Response.Write "            <param name=""C"" value=""" & C & """>" & chr(13) & chr(10)
		Response.Write "            <param name=""S"" value=""" & S & """>" & chr(13) & chr(10)
		Response.Write "            <param name=""F"" value=""0"">" & chr(13) & chr(10)
		Response.Write "            <param name=""BkColor"" value=""fff8e5"">" & chr(13) & chr(10)
		Response.Write "          </applet>" & chr(13) & chr(10)
		Response.Write "    </td>" & chr(13) & chr(10)
	'----------------------------------�H�����ܧ����--------------------------------------------
		Response.Write "      <td valign=""top"">" & chr(13) & chr(10)
		Response.Write "      <table class=""rightside twincolour"" style=""width: 450px !important;"">" & chr(13) & chr(10)
		Response.Write "        <thead>" & chr(13) & chr(10)
		Response.Write "          <tr>" & chr(13) & chr(10)
		Response.Write "            <td>�W��</td>" & chr(13) & chr(10)
		Response.Write "            <td>��</td>" & chr(13) & chr(10)
		Response.Write "            <td>���</td>" & chr(13) & chr(10)
		Response.Write "            <td>�W��</td>" & chr(13) & chr(10)
		Response.Write "            <td>��</td>" & chr(13) & chr(10)
		Response.Write "            <td>���</td>" & chr(13) & chr(10)
		Response.Write "          </tr>" & chr(13) & chr(10)
		Response.Write "        </thead>" & chr(13) & chr(10)
		
		'���"�̲��~"��PIE��
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
					'�W��
					if left(ucase(trim(rs(1))),3)="EB0" then 
						response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount) & ";"">&nbsp;</span> �W��" & replace(replace(trim(rs(5)),"��","")," ","_") & "</div></td>" & chr(13) & chr(10) 
					elseif left(ucase(trim(rs(1))),3)="EB1" then 
						response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount) & ";"">&nbsp;</span> �W�d" & replace(replace(trim(rs(5)),"��","")," ","_") & "</div></td>" & chr(13) & chr(10) 
					else
						response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount) & ";"">&nbsp;</span> " &  replace(replace(trim(rs(5)),"��","")," ","_") & "</div></td>" & chr(13) & chr(10) 
					end if
						
					'��
					if isnumeric (trim(rs(2))) then
						sValue = cdbl(trim(rs(2))) * sitebase
					else
						sValue = rs(2)
					end if
					response.write "          <td>" & Formatnumber(sValue,0,0,0,0) & "</td>" & chr(13) & chr(10) 
					if sValue >= 1 then
						sTotalValue = sTotalValue + Formatnumber(sValue,0,0,0,0)
					end if
					
					'���
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
			
			'�P�_�O�_�����,�Y�O,�h�n�ɻ��ӫ~�k���ťժ��t�@�b
			if (scount mod 2 = 1) and (sTotal >= 100) then
				response.write "              <td>&nbsp;</td>" & vbcrlf
				response.write "              <td>&nbsp;</td>" & vbcrlf
				response.write "              <td>&nbsp;</td></tr>" & vbcrlf
			end if
			
			
			'���o�䥦������
			cs = "wfb1"
			if rno mod 2 = 0 then cs = "wfb2" 
			if (scount mod 2 = 1) then
				if sTotal < 100 then
					sTotal2 = sTotal - sTotal2
					sTotal = 100 - sTotal		
					response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount + 1) & ";"">&nbsp;</span> �䥦</div></td>" & chr(13) & chr(10)
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
					response.write "          <td><div align=""left""><span style=""height:15px; width:15px; background:#" & GetColor(scount + 1) & ";"">&nbsp;</span> �䥦</div></td>" & chr(13) & chr(10)
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
			
			
			'====================��ܦX�p������======================
	
			if (rno mod 2 = 1) then
				response.write "        <tr>" & chr(13) & chr(10) 
			else
				response.write "        <tr class=""odd"">" & chr(13) & chr(10) 
			end if

			response.write "          <td colspan=3>&nbsp;</td>" & chr(13) & chr(10)
			response.write "          <td> �X�p</td>" & chr(13) & chr(10)
			response.write "          <td><span class=""bear"">" & sTotalValue & "</span></td>" & chr(13) & chr(10)
			response.write "          <td><span class=""bear"">" & sPercent & "%</span></td>" & chr(13) & chr(10)
			response.write "        </tr>" & chr(13) & chr(10) 
					
		end if
		Response.Write "      </table>" & chr(13) & chr(10)
	
		Response.Write "      </td></tr>" & chr(13) & chr(10)
		Response.Write "      </table>" & chr(13) & chr(10)

		'���"�̦a��"��PIE��
		Response.Write GetData0(fid)
	
	else
		Response.Write "    <table>" & chr(13) & chr(10)
		Response.Write "      <tr><td>�L���Ѹ��</td></tr>" & chr(13) & chr(10)
		Response.Write "    </table>" & chr(13) & chr(10)
	end if
	
	Response.Write "      <div class=""cleartitle""></div>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	Response.Write "  </div>" & chr(13) & chr(10)
	Response.Write "</div>" & chr(13) & chr(10)
	

	'====================================================================================================================
	

	'============================================  �����s�D (wr05)   ====================================================

	Response.Write "<div class=""inline_block_out"">" & chr(13) & chr(10)
	Response.Write "  <div class=""inline_block"" style=""width: 340px; margin-right: 20px;"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h2 class=""tag_title tw_stock_curve"">�����s�D</h2>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	
	Response.Write "    <table class=""twincolour"">" & chr(13) & chr(10)

	Response.write sayNewsList(page)
	Response.Write "    </table>" & chr(13) & chr(10)
	Response.write "<a href=""/w/wr/wr05_" & fid & ".djhtm"" class=""more"">MORE</a> </div>" & chr(13) & chr(10)
		
	'============================================  ����t�� (wr10)   ====================================================
	
	Response.Write "  <div class=""inline_block"" style=""width: 340px;"">" & chr(13) & chr(10)
	Response.Write "    <div align=""left"">" & chr(13) & chr(10)
	Response.Write "    <h2 class=""tag_title international_clock"">�t���O��</h2>" & chr(13) & chr(10)
	Response.Write "    </div>" & chr(13) & chr(10)
	
	Response.Write "    <table class=""twincolour"">" & chr(13) & chr(10)

	Response.Write "      <thead>" & chr(13) & chr(10)
	Response.Write "        <tr><td>���</td>" & chr(13) & chr(10)
	Response.Write "          <td>���A</td>" & chr(13) & chr(10)
	Response.Write "          <td>����/���</td>" & chr(13) & chr(10)
	Response.Write "          <td>���O</td></tr>" & chr(13) & chr(10)
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
			Response.Write "        <td>�t��</td>" & chr(13) & chr(10)
			Response.Write "        <td>" & trim(rs(1)) & "</td>" & chr(13) & chr(10)
			        
			Select Case trim(rs(2))
				case "AX000010" '�x��
					currName = "�x��"
				case "AX000060" '�ڤ�
					currName = "�ڤ�"
				case "AX000070" '�H����
					currName = "�H����"
				case "AX000085" '���ӹ�		
					currName = "���ӹ�"
				case "AX000090" '�L����		
					currName = "�L����"			
				case else
					currName = "�x��"
			End Select
				            
			Response.Write "        <td class=" & cs & "c>" & currName & "</td>" & chr(13) & chr(10)
			Response.Write "      </tr>" & chr(13) & chr(10)
			
			rs.MoveNext
			if rs.Eof then
				exit for
			end if		
		next
	else 
		Response.Write "      <tr><td colspan=4>�L�t�����</td></tr>" & chr(13) & chr(10)
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
	Response.Write "<tr><td class=wfb0c>�L����򥻸��</td></tr>" & chr(13) & chr(10)
	Response.Write "</table>" & chr(13) & chr(10)
end if
    
    
Response.Write "<script language=""JavaScript""><!--" & chr(13) & chr(10)
'== 2008/10/09 �u�u����W�w open window function ==
Response.Write "function sOpenTradeRule(sFID)" & vbcrlf
Response.Write "{	" & vbcrlf
Response.Write "	var sURL = '/w/wr/wr01rule.djhtm?a='+sFID;	" & vbcrlf
Response.Write "	window.open(sURL,'newwindow',config='width=600,height=300,top=0,left=0,toolbar=0,menubar=0,scrollbars=yes,resizable=no,location=no,status=no');	" & vbcrlf
Response.Write "	return false;	" & vbcrlf
Response.Write "}	" & vbcrlf

'Response.Write "InitComboList(document.wr01_frm.selFID, '/w/wr/wr01_', '.djhtm', '" & fid & "', tfund_fund, '');" & chr(13) & chr(10)
Response.Write "// --></script>" & chr(13) & chr(10)

Response.Write GetDocEplog("Q")

'����b�ȸ��
function NavCol(ary, first, last, limit)
	dim xxx,i
	xxx = ""
	rno = 0
	for i = 0 to 1
		
		if i = 0 then
			xxx = xxx & "      <tr><td nowrap>���</td>" & chr(13) & chr(10)
		else
			xxx = xxx & "      <tr><td nowrap>�b��</td>" & chr(13) & chr(10)
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

'����Z�ļƭ�(�P�_���t)
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


'����������ѹ�
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
			if Tmp1 <> "�X�p" then
				cs = "wfb1"
				if rno mod 2 = 0 then cs = "wfb2" 
				
				sumd = sumd +1
				MyCounter = MyCounter +1
				scount = scount + 1
				Tmp1 = replace(replace(Tmp1,"��","")," ","_")  '�h�����Ϊť�			
				T = T & " " &  Tmp1 
				
					
				sValue = aRs(6,forIdx) 
				if isnumeric (trim(aRs(6,forIdx))) then
					sValue = cdbl(trim(aRs(6,forIdx)))/10
				end if
				'====�p��X�p���ȤΤ��====
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
				
				'================�H�����ܫ��Ѧa�Ϥ��G================
				if scount = 1 then
					sDTable = sDTable & "        <table class=""rightside twincolour"" style=""width: 450px !important;"">" & chr(13) & chr(10)
					sDTable = sDTable & "          <thead>" & chr(13) & chr(10)
				 	sDTable = sDTable & "            <tr>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>�W��</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>��</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>���</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>�W��</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>��</td>" & chr(13) & chr(10)
					sDTable = sDTable & "              <td>���</td>" & chr(13) & chr(10)
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
		'�P�_�O�_�����,�Y�O,�h�n�ɻ��k�����t�@�b
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
		sDTable = sDTable & "              <td> �X�p</td>" & vbcrlf
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
	sD = sD & "              <param name=""T"" value=""������Ѥ��G " & T & """>" & chr(13) & chr(10)
	sD = sD & "              <param name=""V"" value=""" & V & """>" & chr(13) & chr(10)
	sD = sD & "              <param name=""C"" value=""" & C & """>" & chr(13) & chr(10)
	sD = sD & "              <param name=""S"" value=""" & S & """>" & chr(13) & chr(10)
	sD = sD & "              <param name=""F"" value=""1"">" & chr(13) & chr(10)
	sD = sD & "              <param name=""BkColor"" value=""fff8e5"">" & chr(13) & chr(10)
	sD = sD & "            </applet>" & chr(13) & chr(10)
	sD = sD & "          </td>" & chr(13) & chr(10)
	'-----------  ���~�O----------
	sD = sD & "          <td valign=""top"">" & chr(13) & chr(10)
	sD = sD & sDTable 
	sD = sD & "          </td></tr></table>" & chr(13) & chr(10)


	rcnt = 0
	set aRs = nothing
	GetData0 = sD
end function		

'����s�D
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
		xxx = xxx & "      <tr><td>�L�s�D���</td</tr>" & chr(13) & chr(10)
	end if
	sayNewsList = xxx
end function

%>

