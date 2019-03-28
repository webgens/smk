<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ReceiptList.asp - 영수증 발급 리스트
'Date		: 2018.12.31
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include virtual = "/Common/OpenXpay/lgdacom/md5.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oRs1							'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM wQuery							'# WHERE 절
DIM sQuery							'# SORT 절

DIM x
DIM y

Dim rType
Dim LGD_MID
Dim LGD_TID
Dim AuthData
Dim OrderCode
Dim LGD_CASSEQNO
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
rType = sqlFilter(Request("rType"))
LGD_MID = sqlFilter(Request("LGD_MID"))
LGD_TID = sqlFilter(Request("LGD_TID"))
AuthData = sqlFilter(Request("AuthData"))
OrderCode = sqlFilter(Request("OrderCode"))
LGD_CASSEQNO = sqlFilter(Request("LGD_CASSEQNO"))

'Response.Write "rType = " & rType & "<br>"
'Response.Write "LGD_MID = " & LGD_MID & "<br>"
'Response.Write "LGD_TID = " & LGD_TID & "<br>"
'Response.Write "AuthData = " & AuthData & "<br>"
'Response.Write "LGD_CASSEQNO = " & LGD_CASSEQNO & "<br>"

SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>
<%IF PAY_PLATFORM = "test" THEN%>
<script type="text/javascript" src ="<%=MALL_RECEIPT_LINK_TEST%>"></script>
<!--<script type="text/javascript" src ="<%=MALL_ESCROW_LINK_TEST%>"></script>-->
<%ELSE%>
<script type="text/javascript" src ="<%=MALL_RECEIPT_LINK%>"></script>
<!--<script type="text/javascript" src ="<%=MALL_ESCROW_LINK%>"></script>-->
<%END IF%>

<% If rType = "C" Then %>
<script type="text/javascript">
	showReceiptByTID('<%=LGD_MID%>', '<%=LGD_TID%>', '<%=AuthData%>');
</script>
<% ElseIf rType = "B" Then %>
<script type="text/javascript">
	showCashReceipts('<%=LGD_MID%>', '<%=OrderCode%>', '001', 'BANK', '<%=PAY_PLATFORM%>');
</script>
<% ElseIf rType = "V" Then %>
<script type="text/javascript">
	showCashReceipts('<%=LGD_MID%>', '<%=OrderCode%>', '<%=LGD_CASSEQNO%>', 'CAS', '<%=PAY_PLATFORM%>');
</script>
<% End If %>
<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>