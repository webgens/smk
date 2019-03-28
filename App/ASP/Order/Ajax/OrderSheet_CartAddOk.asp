<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheet_CartAddOk.asp - 장바구니에 있는 상품을 주문하기 위해 장바구니에 담긴 상품을 주문서테이블(OrderSheet)에 복사
'Date		: 2018.12.27
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM Flag

DIM ProductCode
DIM ProductName
DIM OrderCnt
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


Flag			= sqlFilter(Request("Flag"))
IF Flag = "" THEN Flag = "PARTICIAL"



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성




'# 품절된 상품 존재여부 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Cart_Select_For_Soldout_Count"

		.Parameters.Append .CreateParameter("@CartID",		adVarChar,	adParamInput,  20,	U_CARTID)
		.Parameters.Append .CreateParameter("@UserID",		adVarChar,	adParamInput,  20,	U_NUM)
		.Parameters.Append .CreateParameter("@Flag",		adVarChar,	adParamInput,  20,	Flag)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF CInt(oRs("SoldoutCount")) > 0 THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||선택하신 상품에 품절된 상품이 있습니다."
				Response.End
		END IF
END IF
oRs.Close




oConn.BeginTrans



'# 장바구니 상품을 주문서로 등록 처리
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Insert_From_Cart"

		.Parameters.Append .CreateParameter("@CartID",				adVarChar,	adParamInput,  20,	 U_CARTID)
		.Parameters.Append .CreateParameter("@UserID",				adVarChar,	adParamInput,  20,	 U_NUM)
		.Parameters.Append .CreateParameter("@Flag",				adVarChar,	adParamInput,  20,	 Flag)
		.Parameters.Append .CreateParameter("@Location",			adChar,		adParamInput,   1,	 "M")
		.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,  20,	 U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,  15,	 U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		oConn.RollbackTrans

		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||주문서 등록 처리중 오류가 발생하였습니다."
		Response.End
END IF



oConn.CommitTrans


Response.Write "OK|||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>