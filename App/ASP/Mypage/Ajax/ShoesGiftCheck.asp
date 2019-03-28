<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ShoesGiftCheck.asp - 슈즈 상품권 입력 체크
'Date		: 2019.01.07
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
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM wQuery							'# WHERE 절
DIM sQuery							'# SORT 절

DIM x
DIM y

DIM CPNO
DIM PCost : PCost = 0

DIM SalePrice
DIM UseFlag
DIM AvailableDT



DIM Socket


DIM ProductNum
DIM ProductID
DIM ProductBarcode
DIM ProductCnt
DIM SupplyCode
DIM ProductCost
DIM PartnerAmt
DIM SupplyAmt
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
CPNO			 = sqlFilter(Request("cpno"))
IF CPNO = "" THEN
		Response.Write "FAIL|||||슈즈 상품권 번호를 입력하여 주십시오."
		Response.End
END IF


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성






'# 사용된 쿠폰 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Happy_Select_For_Check_By_CPNo"

		.Parameters.Append .CreateParameter("@cpno", adVarchar, adParaminput, 16, CPNO)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		PCost = oRs("PCost")
END IF
oRs.Close


IF PCost > 0 THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||이미 등록된 슈즈 상품권 번호입니다."
		Response.End
END IF

	


'# 미사용 쿠폰 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_SCash_Rest_Select_For_Check_By_CPNo"

		.Parameters.Append .CreateParameter("@cpno", adVarchar, adParaminput, 20, CPNO)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		SalePrice	 = oRs("SalePrice")
		UseFlag		 = oRs("UseFlag")
		AvailableDT	 = oRs("AvailableDT")
END IF
oRs.Close


IF UseFlag = "Y" THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||이미 등록된 슈즈 상품권 번호입니다."
		Response.End
END IF
IF CDate(AvailableDT) < Date THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||유효기간이 만료된 슈즈 상품권 번호입니다. [" & AvailableDT & "]"
		Response.End
END IF

Response.Cookies("CPNO")			 = Encrypt(CPNO)
Response.Cookies("ProductNum")		 = Encrypt("1")
Response.Cookies("ProductID")		 = Encrypt("")
Response.Cookies("ProductBarcode")	 = Encrypt("")
Response.Cookies("ProductCnt")		 = Encrypt(1)
Response.Cookies("SupplyCode")		 = Encrypt("SC00000001")
Response.Cookies("ProductCost")		 = Encrypt(SalePrice)
Response.Cookies("PartnerAmt")		 = Encrypt(0)
Response.Cookies("SupplyAmt")		 = Encrypt(0)


	

SET oRs = Nothing
oConn.Close
SET oConn = Nothing



Response.Write "OK|||||"
%>