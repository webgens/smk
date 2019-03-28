<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ShoesGiftAddOk.asp - 슈즈 상품권 입력 처리
'Date		: 2018.12.10
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
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM CPNO
DIM ProductNum
DIM ProductID
DIM ProductBarcode
DIM ProductCnt
DIM SupplyCode
DIM ProductCost
DIM PartnerAmt
DIM SupplyAmt
DIM AdmitNum

DIM PCost

DIM SalePrice
DIM UseFlag
DIM AvailableDT




DIM Socket

DIM SDate
DIM EDate
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
CPNO			 = TRIM(Decrypt(Request.Cookies("CPNO")))


IF CPNO = "" THEN
		Response.Write "FAIL|||||슈즈 상품권 정보가 없습니다. 다시 시도하여 주십시오."
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

		.Parameters.Append .CreateParameter("@CPNo", adVarchar, adParaminput, 16, CPNO)
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





'# 시작/종료일 int_coupon_happy_ev
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_int_coupon_happy_ev_Select_By_cpno"

		.Parameters.Append .CreateParameter("@cpno", adVarchar, adParaminput, 12, CPNO)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		SDate = oRs("date1")
		EDate = oRs("date2")
ELSE
		SDate = CStr(Date)
		EDate = CStr(DateAdd("d", 90, Date))
END IF
oRs.Close








ProductNum			 = TRIM(Decrypt(Request.Cookies("ProductNum")))
ProductID			 = TRIM(Decrypt(Request.Cookies("ProductID")))
ProductBarcode		 = TRIM(Decrypt(Request.Cookies("ProductBarcode")))
ProductCnt			 = TRIM(Decrypt(Request.Cookies("ProductCnt")))
SupplyCode			 = TRIM(Decrypt(Request.Cookies("SupplyCode")))
ProductCost			 = TRIM(Decrypt(Request.Cookies("ProductCost")))
PartnerAmt			 = TRIM(Decrypt(Request.Cookies("PartnerAmt")))
SupplyAmt			 = TRIM(Decrypt(Request.Cookies("SupplyAmt")))








'# TRANSACTION START
oConn.BeginTrans





'# 슈즈상품권 정보 사용정보 입력
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Happy_Insert"

		.Parameters.Append .CreateParameter("@CPNo",			 adVarchar,	 adParaminput, 16, CPNO)
		.Parameters.Append .CreateParameter("@ProductNum",		 adVarchar,	 adParaminput,  4, ProductNum)
		.Parameters.Append .CreateParameter("@ProductID",		 adVarchar,	 adParaminput, 13, ProductID)
		.Parameters.Append .CreateParameter("@ProductBarcode",	 adVarchar,	 adParaminput, 13, ProductBarcode)
		.Parameters.Append .CreateParameter("@ProductCnt",		 adVarchar,	 adParaminput,  4, ProductCnt)
		.Parameters.Append .CreateParameter("@SupplyCode",		 adVarchar,	 adParaminput, 15, SupplyCode)
		.Parameters.Append .CreateParameter("@ProductCost",		 adInteger,	 adParaminput,   , ProductCost)
		.Parameters.Append .CreateParameter("@PartnerAmt",		 adInteger,	 adParaminput,   , PartnerAmt)
		.Parameters.Append .CreateParameter("@SupplyAmt",		 adInteger,	 adParaminput,   , SupplyAmt)
		.Parameters.Append .CreateParameter("@AdmitNum",		 adVarchar,	 adParaminput, 20, AdmitNum)
		.Parameters.Append .CreateParameter("@SDate",			 adDate,	 adParaminput,   , SDate)
		.Parameters.Append .CreateParameter("@EDate",			 adDate,	 adParaminput,   , EDate)
		.Parameters.Append .CreateParameter("@UserID",			 adVarchar,	 adParaminput, 30, U_ID)
		.Parameters.Append .CreateParameter("@CreateID",		 adVarchar,	 adParaminput, 20, U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",		 adVarchar,	 adParaminput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing





'# 슈즈상품권 정보 입력
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_SCash_Insert"
	
		.Parameters.Append .CreateParameter("@MemberNum",		 adBigInt,	 adParaminput,    , U_NUM)
		.Parameters.Append .CreateParameter("@SCode",			 adChar,	 adParaminput,   3, "111")
		.Parameters.Append .CreateParameter("@AddSCash",		 adCurrency, adParaminput,    , ProductCost)
		.Parameters.Append .CreateParameter("@Memo",			 adVarchar,	 adParaminput, 300, "슈즈 상품권 전환")
		.Parameters.Append .CreateParameter("@OrderCode",		 adVarchar,	 adParaminput,  20, "")
		.Parameters.Append .CreateParameter("@OPIdx_Org",		 adInteger,	 adParaminput,    , 0)
		.Parameters.Append .CreateParameter("@CPNo",			 adVarchar,	 adParaminput,  20, CPNO)
		.Parameters.Append .CreateParameter("@AvailableDT",		 adDate,	 adParaminput,    , DateAdd("yyyy", 5, Date))
		.Parameters.Append .CreateParameter("@CreateID",		 adVarchar,	 adParaminput,  20, U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",		 adVarchar,	 adParaminput,  15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing





'# 회원 슈즈상품권 합계 정보 수정
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Update_For_SCash"
		
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,    , U_NUM)
		.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar,	 adParamInput,  20, U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",	 adVarChar,	 adParamInput,  15, U_IP)
	
		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing




'# 슈즈상품권 잔여 정보 수정
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_SCash_Rest_Update_For_Use"
		
		.Parameters.Append .CreateParameter("@CPNo",		 adVarChar,	 adParamInput,  20, CPNO)
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,    , U_NUM)
		.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar,	 adParamInput,  20, U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",	 adVarChar,	 adParamInput,  15, U_IP)
	
		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing








'# TRANSACTION END
oConn.CommitTrans



	




	
SET oRs = Nothing
oConn.Close
SET oConn = Nothing



Response.Write "OK|||||"
Response.End
%>