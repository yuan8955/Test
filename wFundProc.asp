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
				rc = year(p_val) & "�~" & month(p_val) & "��"
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
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��1:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�U����g���޷|�֭�ΦP�N�ͮġA������ܵ��L���I�A����g�z���q�H�����g�z�Z�Ĥ��O�Ұ�����̧C��ꦬ�q�F����g�z���q���ɵ��}�޲z�H���`�N�q�ȥ~�A���t�d�U����������A�礣�O�ҳ̧C�����q�A���H���ʫe���Ծ\������}�����ѡC����������Ӿᤧ�������I�����t�ᤧ�O��(�ҥ~����t���P�O��)�w���S�������}�����ѩΧ��H�������A���H�i�ܤ��}��T�[�����ιҥ~�����T�[�����d�\�C����ëD�s�ڡA������D�ݦs�ګO�I�ӫO�d����H�ݦۭt�����C��������ꭷ�I�A���@���I�i��ϥ����o�����l�A�䤤�i�ध�̤j�l���������H�U�����C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��2:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�Z�ĭp�⬰����O���S�A�B�Ҧ��Ҽ{�t�����p�C����t���v���N�������S�v�A�B�L�h�t���v���N���Ӱt���v�C�Ҧ�����Z�ġA�����L�h�Z�ġA���N���Ӥ��Z�Ī�{�A�礣�O�Ұ�����̧C��ꦬ�q�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��3:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">����b�ȥi��]�����]���ӤW�U�i�ʡA����b�ȶȨѰѦҡA��ڥH������q���i���b�Ȭ��ǡF���~����������������A�H������Ѧ��L�����b�ȰѦһ��F������������y�����A��ڥ���H������q�Ҥ��i���R�^������X�����p���¦�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��4:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">����������t�ᤧ�O�Ρ]�t���P�O�Ρ^�w���S���������}�����ѤΧ��H�������A���H�i�ܹҥ~�����T�[����(http://www.fundclear.com.tw)�U���A�γw�V�`�N�z�H�����d�\�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��5:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�W�z�P��O�ζȨѰѦҡA��ڶO�v�H�U�P����c���D�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��6:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�W�z�u�u����W�w��ƶȨѰѦҡA��ڳW�w���H������}�����Ѭ��D�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��7</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�Х�""(�w�M�P�ֳ�)""��""(���ӳ��ͮ�)""�������ƶȴ��ѭ즳���̰ѦҡA�H�Ѩ䰵�R�^�B�ഫ���~��������M�w�F�u�w�M�P�ֳơv���b�x�W�w�U�[(����)������F�u���ӳ��ͮġv������k�W�W�w���ɮĤ��V���޷|��z�ӳ��ͮħ@�~���ҥ~����C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��8:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">���B�Ѧbū�^�ɡA������q�N�̫����������u�������P��v���������ʤ���O�A�ӶO�αN��ū�^�`�B�������F�t�U������q�ݦ����@�w��v�����P�O�ΡA�N�ϬM��C�����b�Ȥ��A���H�L���B�~��I�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��9:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�̪��޷|�W�w������j���Ҩ饫���������Ҩ餣�o�W�L������겣�b�Ȥ�10%�A��Ӱ�����a�ϥ]�t����j���έ���A����b�ȥi��]���j���a�Ϥ��k�O�B�F�v�θg�����ҧ��ܦӨ����P�{�פ��v�T�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��10:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�����q�Ũ��������Ъ��[�\�C���굥�Ť��D��굥�ŶŨ�A�G�ݩӨ����j������i�ʡA�ӧQ�v���I�B�H�ιH�����I�B�~�תi�ʭ��I�]�N����@���굥�Ť��Ũ�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��11:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">���H��갪���q�Ũ������y�e����զX�L�����񭫡A������g��F�|���ĺʷ��޲z�e���|�֭�A������ܵ��L���I�C�ѩ󰪦��q�Ũ餧�H�ε������F��굥�ũΥ��g�H�ε����A�B��Q�v�ܰʪ��ӷP�׬ư��A�G��������i��|�]�Q�v�W�ɡB�����y�ʩʤU���A�ζŨ�o����c�H������I�����B�Q���ί}���ӻX�����l�A����������A�X�L�k�Ӿ�������I�����H�C����g�z���q�H�����g�z�Z�Ĥ��O�Ұ�����̧C��ꦬ�q�F����g�z���q���ɵ��}�޲z�H���`�N�q�ȥ~�A���t�d�U����������A�礣�O�ҳ̧C�����q�A���H���ʫe���Ծ\������}�����ѡC���~�A���������q�Ũ����i�����ŦX����Rule 144A�W�w�㦳�p�ҩʽ褧�Ũ�A�ѩ����Rule144 A�Ũ�ȭ����c���H�ʶR�A��T���S�n�D���@��Ũ�e�P�A�󦸯ť�������ɥi��]�ѻP�̸��֡A�Υ�����X���N�@���C�A�ɭP���͸��j���R����t�A�i�Ӽv�T����b�ȡC</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��12:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�e�U�H�e�U��ꤧ����Y�k���������q�Ũ髬����A�e�U�H��Ѧ�����������Y�D�n����D��굥�Ť������I�Ũ�A���ꭷ�I�D�n�Ӧ۩�ҧ��Ũ�Ъ����Q�v�ΫH�έ��I�C�Ũ����P�Q�v�Y���ϦV���Y�A�����Q�v�W�ծɱN�ɭP�Ũ����U�^�A���ͧQ�v���I�F���~�A���󰪦��q�Ũ��i�����t�Ũ�o��D��L�k�v�I�������H�έ��I�C�G�����q�Ũ髬��������Ъ��o�ͤW�}�Q�v�ΫH�έ��I�ƥ�ɡA��b�겣���ȥ�N�]���Ӳ��ͪi�ʡC</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��13:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�s�������Ũ��������Ъ��]�t�F�v�B�g�٬۹����í�w���s��������a���Ũ�A�]���N���{�������F�v�B�g���ܰʭ��I�B�Q�v���I�B�ūH���I�P�~�תi�ʭ��I�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��14:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">������D�n��ꭷ�I���]�t�@��T�w���q���~���Q�v���I�B�y�ʭ��I�B�ײv���I�B�H�ΩιH�����I�~�A�ѩ����������곡�����s����a�Ũ�A�ӷs����a���ūH���Ŵ��M���w�}�o��a���C�A�ҥH�Ө����H�έ��I�]�۹�����A�ר��s����a�g�ٰ򥻭��P�F�v���p�ܰʮɡA���i��v�T���v�ů�O�P�Ũ�H�Ϋ~��C�����ꧡ�A�έ��I�B���t�����ܧ�����l����O�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��15:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">������t���i��Ѱ�������q�Υ�������I�C����A�Υѥ�����X�������A�i��ɭP��l�����B��l�C</td>"
	xxx = xxx & "</tr>" & vbcrlf
	xxx = xxx & "<tr>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"" nowrap valign=""top"">��16:</td>" & vbcrlf
	xxx = xxx & "<td class=""wfb4l"">�W�z��ƥu�ѰѦҥγ~�A�Ź��T�۷�ɤO���ѥ��T�T���A���p�����|�β����A�����q�����Y���~�P����󸳨ƩΥ�������H�A�����t����k�߳d���C</td>"
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
	'����ܬ��
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
	xxx = xxx & "<title>���~���" & title & "</title>" & chr(13) & chr(10)
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

'���o�C��϶�
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

'�s����converJS ����js����link
Function ConverJS( strContent)
	Response.Write "document.writeln('" & Replace(strContent, "'", "\'") & "\n');" & chr(13) & chr(10)
end Function

Function MakeButton(rno,bankFundID)
	dim bbb
'	if make_btn = "y" then
		bbb = "<input type=""button"" value=""����"" name=B" & rno & " onclick=""goBuyFund('" & bankfundid & "')"">"
'	else
'		bbb = ""
'	end if
	MakeButton = bbb
End Function

Function Wf09_Msg(idx)
	Dim xx : xx = ""
	
	select case idx
		case 1	'wf09a
			xx = xx & "��1:�Z�ĭp�⬰����O���S�A�B�Ҧ��Ҽ{�t�����p�C����t���v���N�������S�v�A�B�L�h�t���v���N���Ӱt���v�C�Ҧ�����Z�ġA�����L�h�Z�ġA���N���Ӥ��Z�Ī�{�A�礣�O�Ұ�����̧C��ꦬ�q�C<BR>" & vbcrlf
			xx = xx & "��2:����b�ȥi��]�����]���ӤW�U�i�ʡA����b�ȶȨѰѦҡA��ڥH������q���i���b�Ȭ��ǡF���~����������������A�H������Ѧ��L�����b�ȰѦһ��C<BR>" & vbcrlf
			xx = xx & "��3:����������t�ᤧ�O�Ρ]�t���P�O�Ρ^�w���S���������}�����ѤΧ��H�������A���H�i�ܹҥ~�����T�[����(http://www.fundclear.com.tw)�U���A�γw�V�`�N�z�H�����d�\�C<BR>" & vbcrlf
			xx = xx & "��4:�Х�""(�w�M�P�ֳ�)""��""(���ӳ��ͮ�)""�������ƶȴ��ѭ즳���̰ѦҡA�H�Ѩ䰵�R�^�B�ഫ���~��������M�w�F�u�w�M�P�ֳơv���b�x�W�w�U�[ (����) ������F�u���ӳ��ͮġv������k�W�W�w���ɮĤ��V���޷|��z�ӳ��ͮħ@�~���ҥ~����C<BR>" & vbcrlf
			xx = xx & "��5:�ҥ~����g��F�|���ĺʷ��޲z�e���|�֭�Υӳ��ͮĦb�ꤺ�Ҷ��ξP��A������ܵ��L���I�C����g�z���q�H�����g�z�Z�Ĥ��O�Ұ�����̧C��ꦬ�q�F����g�z���q���ɵ��}�޲z�H���`�N�q�ȥ~�A���t�d������������A�礣�O�ҳ̧C�����q�A���H���ʫe���Ծ\������}�����ѡC<BR>" & vbcrlf
			xx = xx & "��6:�W�z��ƥu�ѰѦҥγ~�A�򴼺��۷�ɤO���ѥ��T�T���A���p�����|�β����A�����q�����Y���~�P����󸳨ƩΥ�������H�A�����t����k�߳d���C<BR>" & vbcrlf
			xx = xx & "��7:�b�ꤺ�g�Ҵ����ֳƤ����~����A�D��X���y������B�ϰ쫬(�t��@��a)�M�S����(�]�A�෽�B�Q���ݡB����)�T������C<BR>" & vbcrlf
			xx = xx & "��8:�̤@�~���S�v�A��X���y���e10�ɰ���B�ϰ쫬�e5�ɰ���A�M�S�����e5�ɰ���A�N�o20�����̳��S�v�μзǮt���G�bX Y�ϤW�C<BR>" & vbcrlf
		case 2	'wf09b
			xx = xx & "��1:����Z�ĭp��Ҧ��Ҽ{�t���A����t���v���N�������S�v�A�B�L�h�t���v���N���Ӱt���v�C�Ҧ�����Z�ġA�����L�h�Z�ġA���N���Ӥ��Z�Ī�{�A�礣�O�Ұ�����̧C��ꦬ�q�C<BR>" & vbcrlf
			xx = xx & "��2:����b�ȶȨѰѦҡA��ڥH������q���i���b�Ȭ��ǡC<BR>" & vbcrlf
			xx = xx & "��3:�Ҥ�����g��F�|���ĺʷ��޲z�e���|�֭�b�ꤺ�Ҷ��ξP��A������ܵ��L���I�C����g�z���q�H�����g�z�Z�Ĥ��O�Ұ�����̧C��ꦬ�q�F����g�z���q���ɵ��}�޲z�H���`�N�q�ȥ~�A���t�d������������A�礣�O�ҳ̧C�����q�A���H���ʫe���Ծ\������}�����ѡC<BR>" & vbcrlf
			xx = xx & "��4:�W�z��ƥu�ѰѦҥγ~�A�򴼺�(FundDJ)�۷�ɤO���ѥ��T�T���A���p�����|�β����A�����q�����Y���~�P����󸳨ƩΥ�������H�A�����t����k�߳d���C<BR>" & vbcrlf
			xx = xx & "��5:�ҿװꤺ�Ѳ�������N����H�o��A���Ъ����ꤺ�Ѳ������(���]�t��~�Ҷ����ꤺ)�C<BR>" & vbcrlf
			xx = xx & "��6:�b�Ҧ�������D��SHARPE�ȫe20�W�A�ñN�o20�����̳��S�v�μзǮt���G�bX Y�ϤW�A�bX Y�ϤW�e�X�ꤺ�Ѳ���������������S�v�P�����зǮt�A�H�⥭���u�����I�����I�A�ĤG�P�ĤT�H���A�X�O�u�����H�F�Ĥ@�P�ĤG�H���A�X�n�������H�C�Y��������b�W�z�W�h�����A�N���������J��L���˰���C<BR>" & vbcrlf
		case 3	'wf09c
			xx = xx & "��1:����Z�ĭp��Ҧ��Ҽ{�t���A����t���v���N�������S�v�A�B�L�h�t���v���N���Ӱt���v�C�Ҧ�����Z�ġA�����L�h�Z�ġA���N���Ӥ��Z�Ī�{�A�礣�O�Ұ�����̧C��ꦬ�q�C<BR>" & vbcrlf
			xx = xx & "��2:����b�ȶȨѰѦҡA��ڥH������q���i���b�Ȭ��ǡC<BR>" & vbcrlf
			xx = xx & "��3:�Ҥ�����g��F�|���ĺʷ��޲z�e���|�֭�b�ꤺ�Ҷ��ξP��A������ܵ��L���I�C����g�z���q�H�����g�z�Z�Ĥ��O�Ұ�����̧C��ꦬ�q�F����g�z���q���ɵ��}�޲z�H���`�N�q�ȥ~�A���t�d������������A�礣�O�ҳ̧C�����q�A���H���ʫe���Ծ\������}�����ѡC<BR>" & vbcrlf
			xx = xx & "��4:�W�z��ƥu�ѰѦҥγ~�A�򴼺�(FundDJ)�۷�ɤO���ѥ��T�T���A���p�����|�β����A�����q�����Y���~�P����󸳨ƩΥ�������H�A�����t����k�߳d���C<BR>" & vbcrlf
			xx = xx & "��5:�ҿװꤺ�Ũ髬����]�t���R�^�����������P�L�R�^�����������C<BR>" & vbcrlf
			xx = xx & "��6:�b�Ҧ�������D��SHARPE�ȫe20�W�A�ñN�o20�����̳��S�v�μзǮt���G�bX Y�ϤW�A�bX Y�ϤW�e�X�ꤺ�Ѳ���������������S�v�P�����зǮt�A�H�⥭���u�����I�����I�A�ĤG�P�ĤT�H���A�X�O�u�����H�F�Ĥ@�P�ĤG�H���A�X�n�������H�C�Y��������b�W�z�W�h�����A�N���������J��L���˰���C<BR>" & vbcrlf
		case 4	'wf09d
			xx = xx & "��1:�Z�ĭp�⬰����O���S�A�B�Ҧ��Ҽ{�t�����p�C����t���v���N�������S�v�A�B�L�h�t���v���N���Ӱt���v�C�Ҧ�����Z�ġA�����L�h�Z�ġA���N���Ӥ��Z�Ī�{�A�礣�O�Ұ�����̧C��ꦬ�q�C<BR>" & vbcrlf
			xx = xx & "��2:����b�ȥi��]�����]���ӤW�U�i�ʡA����b�ȶȨѰѦҡA��ڥH������q���i���b�Ȭ��ǡF���~����������������A�H������Ѧ��L�����b�ȰѦһ��C<BR>" & vbcrlf
			xx = xx & "��3:����������t�ᤧ�O�Ρ]�t���P�O�Ρ^�w���S���������}�����ѤΧ��H�������A���H�i�ܹҥ~�����T�[����(http://www.fundclear.com.tw)�U���A�γw�V�`�N�z�H�����d�\�C<BR>" & vbcrlf
			xx = xx & "��4:�Х�""(�w�M�P�ֳ�)""��""(���ӳ��ͮ�)""�������ƶȴ��ѭ즳���̰ѦҡA�H�Ѩ䰵�R�^�B�ഫ���~��������M�w�F�u�w�M�P�ֳơv���b�x�W�w�U�[ (����) ������F�u���ӳ��ͮġv������k�W�W�w���ɮĤ��V���޷|��z�ӳ��ͮħ@�~���ҥ~����C<BR>" & vbcrlf
			xx = xx & "��5:�ҥ~����g��F�|���ĺʷ��޲z�e���|�֭�Υӳ��ͮĦb�ꤺ�Ҷ��ξP��A������ܵ��L���I�C����g�z���q�H�����g�z�Z�Ĥ��O�Ұ�����̧C��ꦬ�q�F����g�z���q���ɵ��}�޲z�H���`�N�q�ȥ~�A���t�d������������A�礣�O�ҳ̧C�����q�A���H���ʫe���Ծ\������}�����ѡC<BR>" & vbcrlf
			xx = xx & "��6:�W�z��ƥu�ѰѦҥγ~�A�򴼺��۷�ɤO���ѥ��T�T���A���p�����|�β����A�����q�����Y���~�P����󸳨ƩΥ�������H�A�����t����k�߳d���C<BR>" & vbcrlf
			xx = xx & "��7:���~�Ũ鬰�b�ꤺ�g�Ҵ����ֳƤ����~�Ũ髬����A�]�A�Ũ髬�B�����q�šB�u���Ũ�Υi�ഫ�Ũ鵥����C<BR>" & vbcrlf
			xx = xx & "��8:�b�Ҧ�������D��SHARPE�ȫe20�W�A�ñN�o20�����̳��S�v�μзǮt���G�bX Y�ϤW�A�bX Y�ϤW�e�X���~�Ũ髬������������S�v�P�����зǮt�A�H�⥭���u�����I�����I�A�ĤG�P�ĤT�H���A�X�O�u�����H�F�Ĥ@�P�ĤG�H���A�X�n�������H�C�Y��������b�W�z�W�h�����A�N���������J��L���˰���C<BR>" & vbcrlf
	end select
	
	Wf09_Msg = xx
end Function


'�Ǧ^������I���q����.
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

'���o�ӻȦ�P�⪺����O�_�i�H���� : approved field
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


'-- ���~��@����� �G�h������U�Կ�� ���� Script --
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