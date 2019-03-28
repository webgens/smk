<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'CartProductOptionChangeOk.asp - 장바구니 상품 사이즈 변경 처리
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

DIM CartIdx
DIM ChgSizeCD

DIM ProductCode
DIM ProductName
DIM OrderCnt
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


CartIdx			= sqlFilter(Request("CartIdx"))
ChgSizeCD		= sqlFilter(Request("ChgSizeCD"))


IF CartIdx = "" OR ChgSizeCD = "" THEN
		Response.Write "FAIL|||||변경할 입력정보가 부족합니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성




SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Cart_Select_By_Idx"

		.Parameters.Append .CreateParameter("@Idx", adInteger, adParamInput, , CartIdx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		ProductCode		= oRs("ProductCode")
		ProductName		= oRs("ProductName")
		OrderCnt		= CInt(oRs("OrderCnt"))
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||해당 상품은 장바구니에 없습니다."
		Response.End
END IF
oRs.Close


'# 기존에 장바구니에 담긴 상품의 재고를 합산
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Cart_Select_For_OrderCnt_By_CartID_ProductCode_SizeCD"

		.Parameters.Append .CreateParameter("@CartID",		adVarChar, adParamInput, 20, U_CARTID)
		.Parameters.Append .CreateParameter("@UserID",		adVarChar, adParamInput, 20, U_NUM)
		.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput,   , ProductCode)
		.Parameters.Append .CreateParameter("@SizeCD",		adVarChar, adParamInput, 50, ChgSizeCD)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		OrderCnt		= OrderCnt + CInt(oRs("OrderCnt"))
END IF
oRs.Close


'# 재고 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Select_For_Available_By_ProductCode_N_SizeCD"

		.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput,   , ProductCode)
		.Parameters.Append .CreateParameter("@SizeCD",		adVarChar, adParamInput, 50, ChgSizeCD)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF CInt(oRs("StockCnt")) < 1 THEN
				oConn.RollbackTrans

				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||" & ProductName & "[" & ChgSizeCD & "] 상품은 품절된 상품 입니다."
				Response.End

		ELSEIF CInt(oRs("StockCnt")) < CInt(OrderCnt) THEN
				oConn.RollbackTrans

				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||" & ProductName & "[" & ChgSizeCD & "] 상품은 재고가 부족합니다."
				Response.End
		END IF
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||주문할 수 없는 상품 입니다."
		Response.End
END IF
oRs.Close



'# 장바구니 사이즈 변경 처리
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Cart_Update_For_SizeCD"

		.Parameters.Append .CreateParameter("@Idx",					adInteger,	adParamInput,    ,	 CartIdx)
		.Parameters.Append .CreateParameter("@SizeCD",				adVarChar,	adParamInput,  50,	 ChgSizeCD)
		.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,  20,	 U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,  15,	 U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||사이즈 변경 처리중 오류가 발생하였습니다."
		Response.End
END IF



Response.Write "OK|||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>