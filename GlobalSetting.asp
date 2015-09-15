<%
'*************************************************************************************
'***************************   共用全域  變數 / 常數 設定  ***************************
'*************************************************************************************
'== 客戶AspID
Const FUNDASPID = "5920trustfund"

'== Table 最大寬度設定
Const g_TableWidth = 715	

'== 基金申購鈕顯示的位置 (左邊或右邊)
Const g_OrderBtnPosition = 1	'設定 1:LEFT / 2:RIGHT / 預設值為1

'== 基金申購鈕顯示的數量 (Default: 一個申購鈕; 若有其他申購或是自選..等類似用途,請修改 count value)
Const g_OrderBtnCount = 1	

'== 基金申購鈕欄位標題名稱(僅有提供一個,若出現多個,請自行設定其他變數名稱)
Const g_OrderFiledTitle = "下單"	

'== 是否顯示 下拉選單中基金名稱後高風險警語文字 '是:true / 否:false
Const g_ShowComboList_WarningNote = false
	
'*************************************************************************************
'*************************** 國內基金 全域變數 / 常數 設定 ***************************
'*************************************************************************************
'== 是否顯示 國內基金銷售申購鈕 變數 :: false->不顯示  / true->顯示  ==
Const g_Customer_ShowFundOrderBtn_Flag = false
	
'== 是否顯示 客戶銷售之國內基金代碼 變數 :: false->不顯示  / true->顯示  ==
Const g_Customer_ShowFundCode_Flag = false

'== 是否顯示 僅客戶銷售之國內基金檔數 變數 :: false->顯示銷售基金  / true->顯示全市場  ==
Const g_Customer_FundListShowall_Flag = false

'== 國內基金申購紐網址 ==
Const g_Customer_FundOrderUrl1 = ""

Dim deftCID : deftCID = "BFZFHA"	'國內基金公司ID
Dim deftFID : deftFID = "ACML08"	'國內基金ID
Dim deftTID : deftTID = "ET001005"	'國內基金類型ID
Dim deftHID : deftHID = "HM000897"	'國內基金經理人ID


'*************************************************************************************
'*************************** 海外基金 全域變數 / 常數 設定 ***************************
'*************************************************************************************
'== 是否顯示 海外基金銷售申購鈕 變數 :: false->不顯示  / true->顯示  ==
Const g_Customer_ShowWFundOrderBtn_Flag = false
	
'== 是否顯示 客戶銷售之海外基金代碼 變數 :: false->不顯示  / true->顯示  ==
Const g_Customer_ShowWFundCode_Flag = false

'== 是否顯示 僅客戶銷售之海外基金檔數 變數 :: false->顯示銷售基金  / true->顯示全市場  ==
Const g_Customer_WFundListShowall_Flag = false

'== 海外基金申購紐網址 ==
Const g_Customer_FundOrderUrl2 = ""

Dim defCID : defCID = "104"	'海外台灣總代理基金公司ID
Dim defFID : defFID = "FTZ01"	'海外基金ID
'Dim defCID1 : defCID1 = "BFC003"	'海外發行基金公司ID


%>