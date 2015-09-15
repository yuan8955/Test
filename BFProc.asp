<!--#include file="lib/format.Asp" -->
<!--#include file="Lib/Debug.Asp" -->
<!--#include Virtual="/Lib/FilterCSFun.asp" -->
<!--#include Virtual="/FundFun/FundCommonFun.asp" -->
<!--#include file="GlobalSetting.asp" -->
<%	
const defNewCode = 1	'判斷是否使用新版 Stored Procdure --  0:false; 1:true

'*****************************************************************************************
' Purpose : 把 HTML 結果轉換成 javascript 輸出
' Input : sCont - html content
' Return : 傳回 document.write 形式的結果
'*****************************************************************************************
Const NullValue = "N/A"
Const sDecimal = 2
Dim g_Decimal
g_Decimal = sDecimal 'Default Decimal format 2
Const sPDecimal = 2
Dim g_PDecimal
g_PDecimal = sPDecimal 'Default Pencent Decimal format 2
Const sSetDefFormat= "SetDef"

function formatMD(sDate)
	dim sT
	sT = cdate(sDate)
	sMonth = Month(sT)
	sDay = Day(sT)
	FormatMD = sMonth & "/2" & sDay
end function

function ChkNumCss(sData)
	sD = "t3n1"

	if isNumeric(sData) then
		sData = cdbl(sData)
		if error then
			err.Clear
			ChkNumCss = sD
			exit function
		else
			err.Clear
			sData = cdbl(sData)
		end if
		if isnumeric(sData) then
			if cdbl(sData) > 0 then
				sD = "t3n1"
			elseif cdbl(sData) < 0 then
				sD = "t3r1"
			end if
		end if
	end if
	ChkNumCss = sD	
end function 

Function ChkData(sData,sParam)
	dim isDef
	isDef = false
	if sParam = sSetDefFormat then
		isDef = true
	else
		isDef = false
	end if

	ChkData = sData

	select case varType(sData)
		case vbNull
			ChkData = NullValue
		case vbString
			ChkData = trim(sData)
		case vbDate
			if isDef then
				if formatdatetime(sData,4) = "00:00" then
					ChkData = FormatYMD(sData)
				else
					ChkData = FormatYMDT(sData)
				end if
			else
				select case ucase(sParam)
					case "A"	'96/11/12
						ChkData = FormatYMD(sData)
					case "B"	'96/11/12 13:24:00
						ChkData = FormatYMDT(sData)
					case "C"	'96/11/12 13:24:00
						ChkData = FormatMD(sData)	
					case "D"	'96/11
						ChkData = FormatYM(sData)
					case "E"	'2003/11/12
						ChkData = FormatYYMD(sData)		
					case "F"	'2003/11
						ChkData = FormatYME(sData)									
					case else
						ChkData = FormatYMD(sData)
				end select
			end if
		case vbInteger
			ChkData = sData
		case vbDecimal
			
			if formatnumber(sData,0,0,0,0) & "" = sData & "" then
				ChkData = sData				
			else
				if isDef then
					ChkData = sFormatNum(sData,g_Decimal)
					if cdbl(ChkData) = 0 then
						ChkData = cdbl(sData)
					end if
				else
					if sParam & "" = "" then
						ChkData = cdbl(sData)						
					else						
						ChkData = sFormatNum(sData,cint(sParam))
						if cdbl(ChkData) = 0 then
							ChkData = cdbl(sData)
						end if
					end if
				end if
			end if
		case vbDouble
			if formatnumber(sData,0,0,0,0) & "" = sData & "" then
				ChkData = sData				
			else
				if isDef then
					ChkData = sFormatNum(sData,g_Decimal)
					if cdbl(ChkData) = 0 then
						ChkData = cdbl(sData)
					end if
				else
					if sParam & "" = "" then
						ChkData = cdbl(sData)						
					else						
						ChkData = sFormatNum(sData,cint(sParam))
						if cdbl(ChkData) = 0 then
							ChkData = cdbl(sData)
						end if
					end if
				end if
			end if
		case else
			ChkData = sData
	end select
end Function

function sFormatNum(sData,sDecimal)
	sValue = "" & sData
	sLast =""
	sLast2 = ""
	sValue = ChkNumZero(sData,sDecimal)
	Dim sDp
	if left(sValue,1) = "." then
		sValue = "0" & sValue
	end if
	if left(sValue,2) = "-." then
		sValue = "-0" & replace(sValue,"-" ,"")
	end if
	sDp = instr(sValue,".")
	if sDp > 0 then
		if cdbl(left(sValue & "" ,sDp-1)) = cdbl(sValue) then
			sValue = left(sValue,sDp-1)
		end if
		sSD = mid(sValue,sDp+1)
		for i = 1 to len(sSD)
			if cdbl(right(sSD,i)) = 0 then
			else
				sValue = left(sValue,sDp) + left(sSD,len(sSD)-i + 1)
				exit for
			end if
		next
	end if
	
	sFormatNum = sValue
end function

''========================================================================
''	Check Format Num 為所要位數時會不會為0
''	如為0則傳回最接近的小數位數值
''========================================================================
function ChkNumZero(sData,sDecimal)
	Dim sValue
	sValue = sData & ""
	if sDecimal <> "" then
		if isnumeric(sValue) then
			sValue = formatnumber(cdbl(sValue),sDecimal,0,0,0)
		end if
	end if
	if cdbl(sValue) = 0 then
		sValue = cdbl(formatnumber(cdbl(sData),10,0,0,0)) & ""
		sValue = mid(sValue & "",instr(sValue,".")+1 )
		for i = 1 to len(sValue)
			if cdbl(left(sValue,i)) <> "0" then
				sValue = formatnumber(sData,i-1,0,0,0)
				if cdbl(sValue) = 0 then
					sValue = formatnumber(sData,i,0,0,0)
				end if
				exit for
			end if
		next
	end if
	ChkNumZero = sValue
end function	

Function ASPToJS(sCont)
    Dim arrCont, i, sContJS, sTmp
    Dim ScriptStart, ScriptEnd
    
    ScriptStart = "<script"
    ScriptEnd = "</script>"
    ScriptEndCRLF = vbCrLf + ScriptEnd
    
    sCont = Replace(sCont, ScriptEnd, ScriptEndCRLF, 1, -1, vbTextCompare)
            
    arrCont = Split(sCont, vbCrLf)
    sContJS = ""
    
    For i = LBound(arrCont) To UBound(arrCont)
        arrCont(i) = Trim(arrCont(i))
        
        If arrCont(i) <> "" Then
            sLine = Replace(arrCont(i), "'", "\'")
            nPos1 = InStr(1, sLine, ScriptStart, vbTextCompare)
            nPos2 = InStr(1, sLine, ScriptEnd, vbTextCompare)
            
            If nPos1 <> 0 Then 'Replace <script
                sTmp = "document.write('" & Left(sLine, nPos1 - 1) & "<scr' + 'ipt" & Mid(sLine, nPos1 + Len(ScriptStart)) & "');"
            ElseIf nPos2 <> 0 Then  'Replace </script>
                sTmp = "document.write('" & Left(sLine, nPos2 - 1) & "</scr' + 'ipt>" & Mid(sLine, nPos2 + Len(ScriptEnd)) & "');"
            Else
                sTmp = "document.write('" & sLine & "');"
            End If
            
            sContJS = sContJS & sTmp & vbCrLf
        End If
    Next
        
    ASPToJS = sContJS
End Function

Function OpenSQL_Fund(strSQL)
	OpenSQL_Fund = Empty
    Dim oRs
            
    If OpenDB_SQL(oRs, strSQL,"FUNDDB") Then
        OpenSQL_Fund = oRs.GetRows
    End If
    
    CloseDB_SQL oRs
End Function

Function OpenSQL_wFund(strSQL)
	OpenSQL_wFund = Empty
    Dim oRs
            
    If OpenDB_SQL(oRs, strSQL,"WFUNDDB") Then
        OpenSQL_wFund = oRs.GetRows
    End If
    
    CloseDB_SQL oRs
End Function

Function OpenSQL_CFO(strSQL)
	OpenSQL_CFO = Empty
    Dim oRs
            
    If OpenDB_SQL(oRs, strSQL,"CFODB") Then
        OpenSQL_CFO = oRs.GetRows
    End If
    
    CloseDB_SQL oRs
End Function

Function OpenDB_SQL(oRs, strSQL,DBSite)
    On Error Resume Next
    select case DBSite
		case "JUSTDB"
			strConn = GetDBconnStr()
		case "CFODB"
			strConn = GetDBconnStrCFO()
		case "FUNDDB"
			strConn = GetDBconnStrFUND()
		case "WFUNDDB"
			strConn = GetDBconnStrWFund()
		case else
			strConn = GetDBconnStr()
	end select 
    Dim oConn
    Set oConn = CreateObject("ADODB.Connection")
    oConn.CursorLocation = 3
    Call oConn.Open(strConn)
    oConn.CommandTimeout = 90
    Set oRs = CreateObject("ADODB.Recordset")
    
    Call oRs.Open(Replace(strSQL, """", "'"), oConn, adOpenStatic)
    Set oRs.ActiveConnection = Nothing '' disconnected
    oConn.Close
    Set oConn = Nothing
    OpenDB_SQL = Not oRs.EOF
    err.Clear
End Function

Function CloseDB_SQL(oRs)
	On Error Resume Next
	
    If Not oRs Is Nothing Then
        oRs.Close
        Set oRs = Nothing
    End If
End Function

'***********************************************************************
' Purpose: 執行 SQL 命命, 並且回傳一個 2 維陣列
' Param:
'		strSQL: SQL command
' Return:
'	成功會傳回 2 維陣列, 失敗時回傳 Empty
' Example:
'	a = OpenSQL("exec spj_mda0050")
'    
'   If IsEmpty(a) Then ' 失敗
'       ...
'	End If

' 	成功, 處理資料....
'	...
'***********************************************************************
Function OpenSQL(strSQL)
    OpenSQL = Empty
    Dim oRs
            
    If OpenDB_SQL(oRs, strSQL,"JUSTDB") Then
        OpenSQL = oRs.GetRows
    End If
    
    CloseDB_SQL oRs
End Function

Function GetDataTable(iD)
	Dim sD
	sD = GetTableStart()
	sD = sD & iD
	sD = sD & GetTableEnd()
	GetDataTable = sD
end function

Function GetDataTableJS(iD)
	Dim sD
	sD = GetTableStartJS()
	sD = sD & ASPToJS(iD)
	sD = sD & GetTableEndJS()
	GetDataTableJS = sD
end function

function GetTableStart()
	Dim sTable
	sTable = "<SCRIPT LANGUAGE=javascript>" & vbcrlf
	sTable = sTable & "<!--" & vbcrlf
	sTable = sTable & "MakeTableStart();" & vbcrlf
	sTable = sTable & "//-->" & vbcrlf
	sTable = sTable & "</SCRIPT>" & vbcrlf
	GetTableStart = sTable
end function

function GetTableEnd()
	Dim sTable
	sTable = "<SCRIPT LANGUAGE=javascript>" & vbcrlf
	sTable = sTable & "<!--" & vbcrlf
	sTable = sTable & "MakeTableEnd();" & vbcrlf
	sTable = sTable & "//-->" & vbcrlf
	sTable = sTable & "</SCRIPT>" & vbcrlf
	GetTableEnd = sTable
end function

function GetTableStartJS()
	GetTableStartJS = "MakeTableStart();" & vbcrlf
end function

function GetTableEndJS()
	GetTableEndJS = "MakeTableEnd();" & vbcrlf
end function


'***********************************************************************
' Purpose: 取得 expired time
' Param:
'		sType : expired time type
' Return:
'	成功會傳回 expired time, 失敗時回傳 Now + 3分鐘
'***********************************************************************
Function GetExpTime(sType)
    GetExpTime = Empty
    
    If sType = "" Then
        Exit Function
    End If
    
    sql = "exec spj_mda00401 '" & sType & "'"
    aData = OpenSQL_fund(sql)
    
    If IsEmpty(aData) Then
        sExpTime = dateadd("s",180,now)
    else
  		sExpTime = CDate(aData(0, 0))
    End If
 
    GetExpTime = sExpTime
End Function

'***********************************************************************
' Function : SetExpTime(exptime)
' Purpose: 設定 expired time
' Param:
'		exptime : expired time
' Return:
'	無
'***********************************************************************
Function SetExpTime(exptime)
	SetExpTime = ""
	xxx = "<!--" & FormatExpT(exptime) & "-->"
	Response.AddHeader "DJ_Expired", xxx
End Function


function GetSelectOption(pageName,formName,fid,cid)
  
  xxx = ""
  xxx = xxx &"<script language=""javascript"" src=""/w/js/WFundlistJS.djjs""></script>" & chr(13) & chr(10)
  xxx = xxx &"<Table cellSpacing=""0"" cellPadding=""0"" width=""100%"" border=""0"" >" & chr(13) & chr(10)
  xxx = xxx &"<tr><td>" & chr(13) & chr(10)
  xxx = xxx & GenComboList(cid,fid,formName)
  xxx = xxx & "</td>"  		  
  xxx = xxx & "</tr></table>"  			

  GetSelectOption = xxx	
end function

'-- 海外基金的 基金下拉選單 相關 Script --
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

'得到基金的名稱,淨值,幣別等資料
'傳入基金的ID , 多個基金ID則以","相隔
'ex : GetFundInfo("AIZ16,AIZ04")
'傳回一個資料2維陣列, 分別為 : Date ,基金ID,基金Name ,幣別名稱 ,淨值 ,漲跌 ,三個月報酬率 ,六個月報酬率,一年報酬率 ,ASPFundID ,幣別ID 
Function GetFundInfo(FId) 
	GetFundInfo = empty
	on error resume next
	strSql = "exec spj_mda70301 '" & FId & "'"
	if OpenDB_SQL(rs, strSql,"FUNDDB") = True then
		GetFundInfo = rs.GetRows     ' 傳回 2 維陣列
		CloseDB_SQL rs        ' 在這裡就可以把 recordset object 清除了...	
	end if
	if err.number then
		GetFundInfo = empty
		exit  Function
	end if
	err.Clear
end function

function isSelected(a,b)
	isSelected = ""
	
	if UCase(a) = UCase (b) then
		isSelected = "selected"
	end if
	
end function

'得到海外基金的基本資料
function GetBasicInfo(fid)
	strSql = "exec spj_mda72151 '" & fid & "'"
	GetBasicInfo =  OpenSQL_Fund(strSql)
end function

'得到日期選單,並有設定applet bcd功能
'傳入 基金代號(fid), form name(formName),表格合併欄數(colns), applet 名稱(appletName),初始開始日期(BDate),初始結束日期(EDate),BCD 檔名(BCDName)
'ex : GetAppletSelect(fid,formName,colns,"CURVE",BDate,EDate,"BCDNavList")
'日期未設,default 是一年前今天 ~ 昨日
'BCD 請傳 A=Fid&B=BDate&C=EDate
function GetAppletSelect(fid,formName,colns,appletName,BDate,EDate,BCDName)
	
	xxx = ""
	xxx = xxx & "<script language=""JavaScript"" src=""/y/js/month.js""></script>" & chr(13) & chr(10)
	xxx = xxx & "<tr><td align=center colspan="& colns &" class=""t3n0"">&nbsp;</td></tr>" & chr(13) & chr(10)
	xxx = xxx & "<tr><td class=t100 colspan="& colns &">"       
	xxx = xxx & "從　<SELECT name=""Y2"" onChange=""javascript:SetMonthDate(document."& formName &".Y2,document."& formName &".M2,document."& formName &".D2);"">" & chr(13) & chr(10)
	xxx = xxx & "	<OPTION value=""91""></OPTION>" & chr(13) & chr(10)
	xxx = xxx & "</SELECT>年 " & chr(13) & chr(10)
	xxx = xxx & "<SELECT name=""M2"" onChange=""javascript:SetMonthDate(document."& formName &".Y2,document."& formName &".M2,document."& formName &".D2);"">" & chr(13) & chr(10)
	xxx = xxx & "   <OPTION value=""-1""></OPTION>" & chr(13) & chr(10)
	xxx = xxx & "</SELECT>月 " & chr(13) & chr(10)
	xxx = xxx & "<SELECT name=""D2"">" & chr(13) & chr(10)
	xxx = xxx & "   <OPTION value=""-1""></OPTION>" & chr(13) & chr(10)
	xxx = xxx & "</SELECT>日 至" & chr(13) & chr(10)
	xxx = xxx & "<SELECT name=""Y1"" onChange=""javascript:SetMonthDate(document."& formName &".Y1,document."& formName &".M1,document."& formName &".D1);"">" & chr(13) & chr(10)
	xxx = xxx & "	<OPTION value=""91""></OPTION>" & chr(13) & chr(10)
	xxx = xxx & "</SELECT>年 " & chr(13) & chr(10)
	xxx = xxx & "<SELECT name=""M1"" onChange=""javascript:SetMonthDate(document."& formName &".Y1,document."& formName &".M1,document."& formName &".D1);"">" & chr(13) & chr(10)
	xxx = xxx & "   <OPTION value=""-1""></OPTION>" & chr(13) & chr(10)
	xxx = xxx & "</SELECT>月 " & chr(13) & chr(10)
	xxx = xxx & "<SELECT name=""D1"">" & chr(13) & chr(10)
	xxx = xxx & "   <OPTION value=""-1""></OPTION>" & chr(13) & chr(10)
	xxx = xxx & "</SELECT>日  " & chr(13) & chr(10)
	xxx = xxx & "<input type=""button"" name=""b1"" value=""查詢"" onClick=""CheckSubmit();"">　　" & chr(13) & chr(10)
	xxx = xxx & "</td></tr>" & chr(13) & chr(10)
	xxx = xxx & "</form>"
	xxx = xxx & "<SCRIPT LANGUAGE=javascript><!--" & chr(13) & chr(10)
	if IsDate(BDate) and IsDate(EDate) then
		xxx = xxx & "	var getYMD1 = '"& EDate &"';" & chr(13) & chr(10)
		xxx = xxx & "	var getYMD2 = '"& BDate &"';" & chr(13) & chr(10)
	else
		xxx = xxx & "	var getYMD1 = '"& (date()-1) &"';" & chr(13) & chr(10)
		xxx = xxx & "	var getYMD2 = '"& (date()-365) &"';" & chr(13) & chr(10)
	end if	
	xxx = xxx & "   PageInit(document."& formName &");" & chr(13) & chr(10)
	
	xxx = xxx & "   function PageInit(obj){		//在 BODY onLoad 時候的初始化" & chr(13) & chr(10)
	xxx = xxx & "		ShowYear(obj.Y1);						" & chr(13) & chr(10)
	xxx = xxx & "		ShowYear(obj.Y2);						" & chr(13) & chr(10)
	xxx = xxx & "		SetOptionValue(obj.M1,1,12);		" & chr(13) & chr(10)
	xxx = xxx & "		SetOptionValue(obj.M2,1,12);		" & chr(13) & chr(10)				
	xxx = xxx & "		//設定初始值						" & chr(13) & chr(10)
	xxx = xxx & "		var YMDary1 = getYMD1.split('/');	" & chr(13) & chr(10)
	xxx = xxx & "		SetFocus(obj.Y1,YMDary1[0]);   " & chr(13) & chr(10)
	xxx = xxx & "		SetFocus(obj.M1,YMDary1[1]);		" & chr(13) & chr(10)
	xxx = xxx & "		var YMDary2 = getYMD2.split('/');	" & chr(13) & chr(10)
	xxx = xxx & "		SetFocus(obj.Y2,YMDary2[0]);   " & chr(13) & chr(10)
	xxx = xxx & "		SetFocus(obj.M2,YMDary2[1]);		" & chr(13) & chr(10)
	xxx = xxx & "		SetMonthDate(obj.Y1,obj.M1,obj.D1);		" & chr(13) & chr(10)
	xxx = xxx & "		SetMonthDate(obj.Y2,obj.M2,obj.D2);		" & chr(13) & chr(10)
	xxx = xxx & "		SetFocus(obj.D2,YMDary2[2]);		" & chr(13) & chr(10)
	xxx = xxx & "		SetFocus(obj.D1,YMDary1[2]);		" & chr(13) & chr(10)	
	xxx = xxx & "	}										" & chr(13) & chr(10)
	
	xxx = xxx & "  function CheckSubmit (){												" & chr(13) & chr(10)
	xxx = xxx & "		var Frm = document."& formName &";										" & chr(13) & chr(10)
	xxx = xxx & "       var y1 = parseInt(Frm.Y1.options[Frm.Y1.selectedIndex].value);	" & chr(13) & chr(10)
	xxx = xxx & "      	var m1 = parseInt(Frm.M1.options[Frm.M1.selectedIndex].value);	" & chr(13) & chr(10)
	xxx = xxx & "      	var d1 = parseInt(Frm.D1.options[Frm.D1.selectedIndex].value);	" & chr(13) & chr(10)
	xxx = xxx & "       var y2 = parseInt(Frm.Y2.options[Frm.Y2.selectedIndex].value);	" & chr(13) & chr(10)
	xxx = xxx & "      	var m2 = parseInt(Frm.M2.options[Frm.M2.selectedIndex].value);	" & chr(13) & chr(10)
	xxx = xxx & "      	var d2 = parseInt(Frm.D2.options[Frm.D2.selectedIndex].value);	" & chr(13) & chr(10)
	xxx = xxx & "      	var sEDate = y1+'/'+m1+'/' + d1;							" & chr(13) & chr(10)		
	xxx = xxx & "      	var sBDate = y2+'/'+m2+'/' + d2;							" & chr(13) & chr(10)		
	xxx = xxx & "      	if(checkBEdate(sBDate,sEDate)){								" & chr(13) & chr(10)	
	xxx = xxx & "			applet_onload('/y/bcd/"& BCDName &".djbcd?a="& fid &"&B='+ sBDate + '&C='+ sEDate ) ;"& chr(13) & chr(10)
	xxx = xxx & "		}																	" & chr(13) & chr(10)
	xxx = xxx & "	}																	" & chr(13) & chr(10)
	
	xxx = xxx & "  function checkBEdate(sBDate,sEDate){												" & chr(13) & chr(10)
	xxx = xxx & "	var aymd1 = sBDate.split('/');													" & chr(13) & chr(10)
	xxx = xxx & "	var aymd2 = sEDate.split('/');													" & chr(13) & chr(10)
	xxx = xxx & "   if (aymd1.length < 3 || aymd2.length < 3 ) {									" & chr(13) & chr(10)
	xxx = xxx & "   	 alert ('日期選擇錯誤!!');													" & chr(13) & chr(10)
	xxx = xxx & "		  return false;																" & chr(13) & chr(10)
	xxx = xxx & "	}																				" & chr(13) & chr(10)
	xxx = xxx & "   var nbdate = parseInt(aymd1[0])*10000+parseInt(aymd1[1])*100+parseInt(aymd1[2]);" & chr(13) & chr(10)
	xxx = xxx & "   var nedate = parseInt(aymd2[0])*10000+parseInt(aymd2[1])*100+parseInt(aymd2[2]);" & chr(13) & chr(10)
	xxx = xxx & "   if (nbdate > nedate) {															" & chr(13) & chr(10)
	xxx = xxx & "   	   alert('您所輸入的起始日期大於結束日期');									" & chr(13) & chr(10)
	xxx = xxx & "		  return false;																" & chr(13) & chr(10)
	xxx = xxx & "	}																				" & chr(13) & chr(10)
	xxx = xxx & "	    return true;																" & chr(13) & chr(10)
	xxx = xxx & "   }																				" & chr(13) & chr(10)
	xxx = xxx & "//--></SCRIPT>" & chr(13) & chr(10)    

	GetAppletSelect = xxx
end function


'做出國內基金的select
'傳入 目前頁面(pageName), form name(formName),表格合併欄數(colns), 基金代號(fid)
'ex : GetSelectOptionTW("wr02","wr02_frm",5,"AIZ16")
'傳回選單項目
function GetSelectOptionTW(pageName,formName,colns,fid)
  
  xxx = ""    
  xxx = xxx & "		" & "<tr><td class=t10 colspan="& colns &">"
  xxx = xxx & "		" & "<SELECT name=selFID onchange=selopn(this.options[this.selectedIndex].value)>" & chr(13) & chr(10)
  for k=0 to 9
    xxx = xxx & "		" & "<OPTION>ＸＸＸＸＸＸＸＸＸＸＸ</OPTION>" & chr(13) & chr(10)
  next
  xxx = xxx & "		" & "</SELECT>" & chr(13) & chr(10)      
  
  xxx = xxx & "		" & "<select onchange=""selopn(this.options[this.selectedIndex].value )"" name=""IDS"" size=""1"">" & chr(13) & chr(10)
  xxx = xxx & "		" & "<option " & "value=""/w/wr/wr01_" & fid & ".djhtm"" "& isSelected(pageName,"wr01") &">基金基本資料</option>" & chr(13) & chr(10)
  xxx = xxx & "		" & "<option " & "value=""/w/wr/wr02_" & fid & ".djhtm"" "& isSelected(pageName,"wr02") &">基金淨值表</option>" & chr(13) & chr(10)
  xxx = xxx & "		" & "<option " & "value=""/w/wr/wr03_" & fid & ".djhtm"" "& isSelected(pageName,"wr03") &">基金績效表</option>" & chr(13) & chr(10)
  xxx = xxx & "		" & "<option " & "value=""/w/wr/wr04_" & fid & ".djhtm"" "& isSelected(pageName,"wr04") &">基金持股狀況</option>" & chr(13) & chr(10)
  xxx = xxx & "		" & "<option " & "value=""/w/wr/wr05_" & fid & ".djhtm"" "& isSelected(pageName,"wr05") &">基金相關新聞</option>" & chr(13) & chr(10)
  xxx = xxx & "		" & "<option " & "value=""/w/wr/wr06_" & fid & ".djhtm"" "& isSelected(pageName,"wr06") &">基金績效-風險報酬</option>" & chr(13) & chr(10)
  xxx = xxx & "		" & "<option " & "value=""/w/wr/wr07_" & fid & ".djhtm"" "& isSelected(pageName,"wr07") &">基金績效-同業排名</option>" & chr(13) & chr(10)
  xxx = xxx & "		" & "<option " & "value=""/w/wr/wr08_" & fid & ".djhtm"" "& isSelected(pageName,"wr08") &">基金績效-多空市場</option>" & chr(13) & chr(10)
  'xxx = xxx & "		" & "<option " & "value=""/w/wr/wr09_" & fid & ".djhtm"" "& isSelected(pageName,"wr09") &">即時新聞</option>" & chr(13) & chr(10)
  xxx = xxx & "		" & "</select>" & chr(13) & chr(10)
  xxx = xxx & "		" & "</td></tr>"  			
  xxx = xxx & "<script language=""javascript""><!--" & chr(13) & chr(10)
  xxx = xxx & "InitComboList(document."& formName &".selFID, '/w/wr/"& pageName &"_', '.djhtm', '" & fid & "', tfund_fund,'');" & chr(13) & chr(10)
  
  xxx = xxx & "setTimeout(""initSelect()"", 300);" & chr(13) & chr(10)
  
  xxx = xxx & "function initSelect() { " & chr(13) & chr(10)
  xxx = xxx & "		initSelect2(document."& formName &".IDS, '/w/wr/"& pageName &"_"& fid &".djhtm'); 	 " & chr(13) & chr(10)
  xxx = xxx & "}" & chr(13) & chr(10)
  
  xxx = xxx & "function initSelect2(obj,sVal) { " & chr(13) & chr(10)
  xxx = xxx & "		for (i=0;i < obj.length ;i++)   	 " & chr(13) & chr(10)
  xxx = xxx & "			if (obj.options[i].value == sVal) obj.selectedIndex = i ;" & chr(13) & chr(10)  
  xxx = xxx & "}" & chr(13) & chr(10)
  
  xxx = xxx & "// --></script>" & chr(13) & chr(10)

  GetSelectOptionTW = xxx	
end function  


'========== 92/03/06 整合 wFundProc.asp & wtFundProc.asp 共用的相同 Function =================
Function FormatYM(d)
	Dim xxx, mm, dd
	
	FormatYM = d
	if not IsDate(d) then
		exit Function
	end if
	
	mm = Month(CDate(d))
	If CInt(mm) < 10 Then
		mm = "0" & mm
	End If		
	if CInt(Year(CDate(d)) - 1911) <= 0 then
		xxx = "N/A"
	else
		xxx = (Year(CDate(d)) - 1911) & "/" & CStr(mm)
	end if		

	FormatYM = CStr(xxx)
End Function

'for 西元的YYYY/MM
Function FormatYME(d)
	Dim xxx, mm, dd
	
	FormatYME = d
	if not IsDate(d) then
		exit Function
	end if
	
	mm = Month(CDate(d))
	If CInt(mm) < 10 Then
		mm = "0" & mm
	End If		
	if CInt(Year(CDate(d))) <= 0 then
		xxx = "N/A"
	else
		xxx = Year(CDate(d)) & "/" & CStr(mm)
	end if		
	
	FormatYME = CStr(xxx)
End Function

Function FormatYMD(d)
	Dim xxx, mm, dd

	FormatYMD = d
	if not IsDate(d) then
		exit Function
	end if
	
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
	
	FormatYMD = CStr(xxx)
End Function

Function FormatYYMD(d)
	Dim xxx, mm, dd
	
	FormatYYMD = d
	if not IsDate(d) then
		exit Function
	end if
	
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
	
	FormatYYMD = CStr(xxx)
End Function

Function FormatYMDT(d)
	dim xxx, hh, mm, ss
	
	FormatYMDT = d
	if not IsDate(d) then
		exit Function
	end if
	
	xxx = FormatYMD(d)
	if xxx = NullValue then
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

	FormatExpT = d
	if not IsDate(d) then
		exit Function
	end if

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

    FormatExpT = CStr(xxx)
End Function

''========================================================================
'' sURL:傳入的網址
'' 傳回Check過的網址
''========================================================================
function ChkURL(sURL)
	ChkURL = lcase(sURL)
	if instr(ChkURL,"http://") > 0 then
	else
		ChkURL = "http://" & sURL
	end if
end function


'=== 2008/10/09 : 國內外基金文件下載 Function ===
'--- 國內外單一基金文件下載相關 Function Start ---
'基金資料下載的Host
function GetFundInfoHost()
	GetFundInfoHost = "http://fundreport.funddj.com/"
	if bUseTestDB then 
		GetFundInfoHost = "http://fundreport.funddj.com/"
	end if	
end function

'國內基金資料下載的連結
function GetTWFundInfoURL(sFID,sType)
	GetTWFundInfoURL = GetFundInfoHost() & "GetTWFundInfo1.asp?A=" & sFID & "&b=" & sType & "&c="& CalcKey((sFID&sType))	
end function

'海外基金資料下載的連結
function GetFundInfoURL(sFID,sType)
	GetFundInfoURL = GetFundInfoHost() & "GetFundInfo1.asp?A=" & sFID & "&b=" & sType & "&c="& CalcKey((sFID&sType))	
end function


'計算字元Key 值, 直接相加
Function CalcKey(sStr)
	ikey = 0
	for i = 1 to len(sStr)
		ikey = ikey + Asc(mid(sStr, i, 1)) 
	next
	CalcKey = ikey
end function


'== 取得國內基金 : 簡式公開說明書 ==
Function GetFundEasyReport(sFID)
	Dim sql,conn,rs
	Dim show_yReport : show_yReport = ""
	
	sql = "select * from yReport where FundID='" & sFID & "' "
	if OpenFundDJ(conn, rs, sql) then		
		while not rs.EOF 
			show_yReport = "      <li><a href=""" & GetTWFundInfoURL(sFID,"4") & """ target=""_blank"">簡式公開說明書</a></li>" & vbcrlf
			rs.movenext		
		wend	
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing		
	end if
	
	GetFundEasyReport = show_yReport
End Function

'--- 國內外單一基金文件下載相關 Function End ---


'== SP 參數資料確認 Start ==
Function checkVar(sVarName,sFlag)
	Dim tmpVar : tmpVar = ""
	Dim tmpVar2 
	
	if sFlag = true then
		if sVarName = "" then
			tmpVar = ",null,null"
		else
			tmpVar2 = split(sVarName,"~")			
			if isArray(tmpVar2) then
				if tmpVar2(0) = 0 and tmpVar2(1) = 0 then
					tmpVar = ",null,null"
				else				
					tmpVar = "," & tmpVar2(0) & "," & tmpVar2(1)
				end if
			else
				tmpVar = ",null,null"
			end if
		end if
	else
		if sVarName = "" or sVarName = "0" or isNull(sVarName) then
			tmpVar = ",null"
		else
			tmpVar = ",'" & sVarName & "'"
		end if
	end if	
	
	checkVar = tmpVar
end Function


'如果參數值為 "" 、"0" 、null 則忽略該參數 ,因為sp 有預設值
function CheckSPParam(sParamName , sParamVal)
	Dim retStr : retStr = ""
	CheckSPParam = retStr
	if sParamVal = "" or sParamVal = "0" or isNull(sParamVal) then
		'傳回空字串,使用 sp 的 default
	else
		retStr = ", " & sParamName & " ='" & sParamVal & "' "
	end if
	
	CheckSPParam = retStr
end function

'如果參數值為 "" 、"0" 、null 則忽略該參數 ,因為sp 有預設值
function CheckSPParam2(sParamName1 , sParamName2, sParamVal)
	Dim retStr : retStr = ""
	CheckSPParam2 = retStr
	if sParamVal = "" or sParamVal = "0~0" or isNull(sParamVal) Then
		'傳回空字串
	else
		tmpVar2 = split(sParamVal,"~")		
		if isArray(tmpVar2) then
			if tmpVar2(0) = 0 and tmpVar2(1) = 0 then
					'傳回空字串
			else				
				retStr = ", " & sParamName1 & " ='" & tmpVar2(0) & "'," & sParamName2 & "='" &  tmpVar2(1) & "'"
			end if
		else
				'傳回空字串
		end if
	End if	
	CheckSPParam2 = retStr
end function
'== SP 參數資料確認 End   ==

'== 列表頁面 筆數切頁顯示 Function ==
Sub GenPageList(sFrmName,sPathUrl,sPageCount,nowPage,sColNum)
	Dim sCT
	
	Response.Write "<TR><TD class=""wfb2c"" colspan=""" & sColNum & """>" & vbcrlf

	If nowPage <> 1 Then
		Response.Write "<A HREF=" & sPathUrl & "&Page=1>第一頁</A>　" & vbcrlf
		Response.Write "<A HREF=" & sPathUrl & "&Page=" & page-1 & ">上一頁</A>　" & vbcrlf
	End If
	If nowPage <> sPageCount Then
		Response.Write "<A HREF=" & sPathUrl & "&Page=" & page+1 & ">下一頁</A>　" & vbcrlf
		Response.Write "<A HREF=" & sPathUrl & "&Page=" & sPageCount & ">最後一頁</A>　" & vbcrlf
	End If
	
	Response.Write "直接到第" & vbcrlf
	Response.Write "<SELECT name=""sel"" class=""s"" onchange=""javascript:chgPage('" & sFrmName & "','" & sPathUrl & "');"">" & vbcrlf
	for sCT = 1 to sPageCount
		if sCT = nowPage then
			Response.Write "<OPTION value=""" & sCT & """ selected>" & sCT & "</OPTION>" & vbcrlf
		else
			Response.Write "<OPTION value=""" & sCT & """>" & sCT & "</OPTION>" & vbcrlf
		end if
	next
	Response.Write "</SELECT>" & vbcrlf
	Response.Write "頁 共 <FONT color=#ff6600>" & sTotalCount & "</FONT> 筆基金資料" & vbcrlf
	Response.Write "</TD></TR>" & vbcrlf

	Response.Write "<SCRIPT LANGUAGE=javascript><!--	" & vbcrlf
	Response.Write "function chgPage(sFrm,sPath)	" & vbcrlf
	Response.Write "{	" & vbcrlf	
	Response.Write "	var sURL = sPath + '&Page=' ;	" & vbcrlf
	Response.Write "	var sObj = eval('document.' + sFrm + '.sel');	" & vbcrlf
	Response.Write "	var idx = sObj.options[sObj.selectedIndex].value;	" & vbcrlf
	Response.Write "	sURL = sURL + idx;	" & vbcrlf
	Response.Write "	document.location = sURL;	" & vbcrlf
	Response.Write "}	" & vbcrlf
	Response.Write "//--></SCRIPT>" & vbcrlf
end Sub

'== 新聞 / 報告 : 日期查詢 Function ==
Sub GenDataQueryForm(sColNum,sURL,sQueryDate)
	Response.Write "<form name=sch onSubmit=""return go1()"">" & vbcrlf
	Response.Write "<tr id=""oScrollFoot""><td class=wfb2c colspan=" & sColNum & ">" & vbcrlf
	Response.Write "以西元日期(yyyy/mm/dd)查詢<input type=text name=B size=8 value=" & datechg(sQueryDate) & ">" & chr(13)  & chr(10)
	Response.Write "<input type=button value=GO name=b1 onclick=""return go1()"">" & vbcrlf
	Response.Write "<script language=""Javascript"" src=""/w/js/jschkd.djjs""></script>" & vbcrlf
	Response.Write "<script language=""JavaScript""><!--" & vbcrlf
	Response.Write "	function go1() {" & vbcrlf
	Response.Write "   var B = document.sch.B.value;" & vbcrlf
	Response.Write "	if (B == '') {" & vbcrlf
	Response.Write "		B='NA';" & vbcrlf
	Response.Write "	}" & vbcrlf
	Response.Write "	else if ((B = chkYDate(B,1)) == false){" & vbcrlf
	Response.Write "		return false;" & vbcrlf
	Response.Write "	}" & vbcrlf
	Response.Write "	self.location='" & sURL & "' + B;"& vbcrlf
	Response.Write "	return false;} " & vbcrlf
	Response.Write "// --></script>" & vbcrlf
	Response.Write "<br>(基金新聞至查詢日為止)" & vbcrlf
	Response.Write "</td></tr>" & vbcrlf	    
	Response.Write "</form>" & vbcrlf
end Sub


' --- 產生 商品績效頁面上的 國內外基金 單一三層式選單 Function ---
Function GenAllFundComboList(FormName,AreaID,CID,FID)
	Dim xx : xx = ""
	xx = xx & "<SCRIPT LANGUAGE=javascript><!--" & vbCrLf
	'xx = xx & " alert(cuteduck); " & vbcrlf
	xx = xx & " var oFormObj = eval('document.' + '" & FormName & "'); " & vbcrlf
	xx = xx & "	iID = '" & FID & "';" & vbCrLf
	xx = xx & "	GenALLFundCorpCombo('" & AreaID & "','" & CID & "','" & FID & "','" & FormName & "');" & vbCrLf
	
'	xx = xx & "	for (i=0;i<oFormObj.oFund_area.options.length;i++)" & vbCrLf
'	xx = xx & "	{" & vbCrLf
'	xx = xx & "		var sTID = oFormObj.oFund_area.options[i].value.toUpperCase();" & vbCrLf
'	xx = xx & "		if (sTID == '" & AreaID & "') " & vbCrLf
'	xx = xx & "		{" & vbCrLf
'	xx = xx & "			oFormObj.oFund_area.selectedIndex = i;" & vbCrLf
'	xx = xx & "			break;" & vbCrLf
'	xx = xx & "		}" & vbCrLf
'	xx = xx & "	}" & vbCrLf
	
	xx = xx & "	for (i=0;i<oFormObj.oFund_corp.options.length;i++)" & vbCrLf
	xx = xx & "	{" & vbCrLf
	xx = xx & "		var tmpID1 = oFormObj.oFund_corp.options[i].value.toUpperCase();" & vbCrLf
	xx = xx & "		if (tmpID1 == '" & CID & "') " & vbCrLf
	xx = xx & "		{" & vbCrLf
	xx = xx & "			oFormObj.oFund_corp.selectedIndex = i;" & vbCrLf
	xx = xx & "			break;" & vbCrLf
	xx = xx & "		}" & vbCrLf
	xx = xx & "	}" & vbCrLf
	
	xx = xx & "	for (i=0;i<oFormObj.oFund3.options.length;i++)" & vbCrLf
	xx = xx & "	{" & vbCrLf
	xx = xx & "		var tmpID2 = oFormObj.oFund3.options[i].value.toUpperCase();" & vbCrLf
	xx = xx & "		if (iID != '')" & vbCrLf
	xx = xx & "		{" & vbCrLf
	xx = xx & "			if (tmpID2 == iID )" & vbCrLf
	xx = xx & "			{" & vbCrLf
	xx = xx & "				oFormObj.oFund3.selectedIndex = i;" & vbCrLf
	xx = xx & "				break;" & vbCrLf
	xx = xx & "			}" & vbCrLf
	xx = xx & "		}" & vbCrLf
	xx = xx & "		else" & vbCrLf
	xx = xx & "			oFormObj.oFund3.selectedIndex = 0;" & vbCrLf
	xx = xx & "	}" & vbCrLf
	xx = xx & "//--></SCRIPT>" & vbCrLf
	
	GenAllFundComboList = xx
end Function

function getFundDJ_IDName(id)
	Dim sql,conn,rs
	Dim rc

	rc = ""  
	sql = "select yb000020 from ya000000 where yb000010='" & id & "'"
	if OpenFundDJ(conn, rs, sql) then
		rc = replace(trim(rs(0))," ","")
		rs.close
		conn.close
		set rs = nothing
		set conn = nothing
	end if

	if rc = "" then 
		sURL = "/z/mda.djxml?x=72502&a=1"
		set oXML1 = GetXMLfromURL(sURL)
		queryXML="/Result/Data/Row[@V1='"& id &"']"	
		Set xml_subNodes = oXML1.selectSingleNode(queryXML)
        if not (xml_subNodes Is Nothing) then
		    'Response.Write xml_subNodes.length
		    rc = xml_subNodes.Attributes.getNamedItem("V2").nodeValue
        end if
	end if
	  
	rc = replace(rc, " ", "_")

	getFundDJ_IDName = rc
end function

'********************************************************************************
' Purpose: 取得客戶所需 xml 資料
' Param: 
'		sURL: 客戶所需的資料的 URL (從客戶的 xml 檔取得此 URL)
' Return: 傳回 xml object or Nothing(If error)
'********************************************************************************
Function GetXMLfromURL(sURL)
	Dim sXML, oXML
	Set GetXMLfromURL = Nothing
	
	If sURL = "" Then
		Exit Function
	End If
	
	sXML = mdjHTTP(sURL) & ""
	
	If sXML = "" Then
		Exit Function
	End If
	
	Set oXML = LoadXMLStr(sXML)
	
	Set GetXMLfromURL = oXML
	
	Set oXML = Nothing
End Function
'********************************************************************************
' Purpose: 透過 xml 字串取得 xml
' Param:
'		sStr: 傳入的 xml 字串
' Return: 傳回 xml object or Nothing(If error)
'********************************************************************************
Function LoadXMLStr(sStr)
	Dim oXML
	
	Set LoadXMLStr = Nothing
	Set oXML = CreateObject("MSXML2.FreeThreadedDOMDocument")
	oXML.async = False
	oXML.loadXML(sStr)
	
	If Err.Number <> 0 Then
		Set oXML = Nothing
		Exit Function
	End If
	Set LoadXMLStr = oXML
End Function


Function mdjHTTP(sURL)
	Err.Clear
    On Error Resume Next
	
	mdjHTTP = ""
	set httpObj = server.CreateObject("DJHTTP.Http")
	
	
	sHost = "http://127.0.0.1"
	
	sURL = sHost & sURL

	httpObj.Url = sURL
	httpObj.Request
	
	sResult = Trim(CStr(httpObj.ResponseString))
	if InStr(sResult, "伺服器目前太忙碌了") then
		sResult = ""
	end if
		
	mdjHTTP = sResult
	set httpObj = nothing
End Function



'國內外基金月報資料下載的連結
function GetFundMonthReport(sFID,sType)
	GetFundMonthReport = GetFundInfoHost() & "GetFundMonthReport1.asp?A=" & sFID & "&b=" & sType & "&c="& CalcKey((sFID&sType))	
end function

'***********************************************************************
' Purpose:  SQL Injection Protect
' Param:    p_str:帶入的參數
' Return:   String
'***********************************************************************
Function SqlTok(p_str)
  if isnull(p_str) or ucase(p_str)= "NULL"then
    SqlTok = "null"
  else
  
    SqlTok = "'" & Replace(p_str,"'","''") & "'"
 end if
End Function

function gotTopic(Str,StrLen)
    If Str="" Then
            gotTopic=""
            Exit Function
    End If
    Dim l,t,c, i
    Str=Replace(Replace(Replace(Replace(Str,regStr," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
    l=Len(Str)
    t=0    
    For i=1 To l
            c=Abs(Asc(Mid(str,i,1)))    '用ascii碼來檢查是否為漢字
            If c>255 Then
                    t=t+2                   
            Else
                    t=t+1
            End If
            If t>=Strlen Then
                    gotTopic=Left(Str,i) & "..."
                    Exit For
            Else
                    gotTopic=Str
            End if
    next
    
    gotTopic=replace(Replace(Replace(Replace(Replace(gotTopic," ",regStr),chr(34),"&quot;"),">","&gt;"),"<","&lt;"),"+","%2B")
end function

function showAppletName(sFundName)
	Dim sName : sName = ""
	sName =gotTopic(trim(sFundName),30)

	showAppletName = sName
end function
%>