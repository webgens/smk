<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************/
'MispwapUrl.asp - 카드결제ISP / 계좌이체 결제 완료 후 리턴 되는 페이지(결제완료 주문 처리하는 곳은 note_url.asp)
'Date		: 2018.12.30
'Update	: 
'/****************************************************************************************/

'//페이지 응답헤더 설정------------------------------------------------------
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//-------------------------------------------------------------------------------
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include virtual = "/Common/ProgID1.asp" -->

<%
'/****************************************************************************************/
'변수 선언 START
'-----------------------------------------------------------------------------------------------------------'
DIM oConn								'# ADODB Connection 개체
DIM oRs									'# ADODB Recordset 개체
DIM oCmd								'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery								'# WHERE 절
DIM sQuery								'# SORT 절

DIM OrderCode
DIM SettleFlag
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode       = Trim(Request("LGD_OID"))		'# 주문번호


Set oConn		= ConnectionOpen()	'//커넥션 생성
Set oRs			= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'//주문아이디(LGD_OID)에 해당하는 아이디를 검색
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_Order_Select_By_OrderCode"
		.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,	adParamInput,	20,		OrderCode)
END WITH
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing


IF NOT oRs.EOF THEN
		SettleFlag				= oRs("SettleFlag")
ELSE
		oRs.Close
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing

		Call AlertMessage2("정상적으로 결제는 되었으나 결제 정보에 오류가 있습니다. 관리자에게 문의 바랍니다.", "location.replace('/');")
		Response.End
END IF
oRs.Close


Set oRs = Nothing
oConn.Close
Set oConn = Nothing


IF SettleFlag = "N" THEN
		Call AlertMessage2("정상적으로 결제는 되었으나 결제 정보에 오류가 있습니다. 관리자에게 문의 바랍니다.", "location.replace('/');")
		Response.End
ELSE
%>
		<script type="text/javascript">
			location.replace("/ASP/Order/OrderComplete.asp?OrderCode=<%=LGD_OID%>");
		</script>
<%
END IF
%>