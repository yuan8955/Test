<%
'*************************************************************************************
'***************************   �@�Υ���  �ܼ� / �`�� �]�w  ***************************
'*************************************************************************************
'== �Ȥ�AspID
Const FUNDASPID = "5920trustfund"

'== Table �̤j�e�׳]�w
Const g_TableWidth = 715	

'== ������ʶs��ܪ���m (����Υk��)
Const g_OrderBtnPosition = 1	'�]�w 1:LEFT / 2:RIGHT / �w�]�Ȭ�1

'== ������ʶs��ܪ��ƶq (Default: �@�ӥ��ʶs; �Y����L���ʩάO�ۿ�..�������γ~,�Эק� count value)
Const g_OrderBtnCount = 1	

'== ������ʶs�����D�W��(�Ȧ����Ѥ@��,�Y�X�{�h��,�Цۦ�]�w��L�ܼƦW��)
Const g_OrderFiledTitle = "�U��"	

'== �O�_��� �U�Կ�椤����W�٫ᰪ���Iĵ�y��r '�O:true / �_:false
Const g_ShowComboList_WarningNote = false
	
'*************************************************************************************
'*************************** �ꤺ��� �����ܼ� / �`�� �]�w ***************************
'*************************************************************************************
'== �O�_��� �ꤺ����P����ʶs �ܼ� :: false->�����  / true->���  ==
Const g_Customer_ShowFundOrderBtn_Flag = false
	
'== �O�_��� �Ȥ�P�⤧�ꤺ����N�X �ܼ� :: false->�����  / true->���  ==
Const g_Customer_ShowFundCode_Flag = false

'== �O�_��� �ȫȤ�P�⤧�ꤺ����ɼ� �ܼ� :: false->��ܾP����  / true->��ܥ�����  ==
Const g_Customer_FundListShowall_Flag = false

'== �ꤺ������ʯú��} ==
Const g_Customer_FundOrderUrl1 = ""

Dim deftCID : deftCID = "BFZFHA"	'�ꤺ������qID
Dim deftFID : deftFID = "ACML08"	'�ꤺ���ID
Dim deftTID : deftTID = "ET001005"	'�ꤺ�������ID
Dim deftHID : deftHID = "HM000897"	'�ꤺ����g�z�HID


'*************************************************************************************
'*************************** ���~��� �����ܼ� / �`�� �]�w ***************************
'*************************************************************************************
'== �O�_��� ���~����P����ʶs �ܼ� :: false->�����  / true->���  ==
Const g_Customer_ShowWFundOrderBtn_Flag = false
	
'== �O�_��� �Ȥ�P�⤧���~����N�X �ܼ� :: false->�����  / true->���  ==
Const g_Customer_ShowWFundCode_Flag = false

'== �O�_��� �ȫȤ�P�⤧���~����ɼ� �ܼ� :: false->��ܾP����  / true->��ܥ�����  ==
Const g_Customer_WFundListShowall_Flag = false

'== ���~������ʯú��} ==
Const g_Customer_FundOrderUrl2 = ""

Dim defCID : defCID = "104"	'���~�x�W�`�N�z������qID
Dim defFID : defFID = "FTZ01"	'���~���ID
'Dim defCID1 : defCID1 = "BFC003"	'���~�o�������qID


%>