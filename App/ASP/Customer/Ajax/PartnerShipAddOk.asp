<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'PartnerShipAddOk.asp - 입점/대량 문의, 단체구매 글등록
'Date		: 2019.01.07
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절


DIM Category
DIM Company
DIM Contents
'DIM Tel
DIM Phone
DIM HP1
DIM HP23
DIM ReceiveAgree
DIM Email
DIM WrtName
DIM BrandName

'//입점/대량문의
DIM Homepage
DIM CompanyCate
DIM Distribution
'//단체구매
DIM ProductCode
DIM OrderQty
DIM NeedDate

DIM UF
DIM SaveFolder
DIM SaveFile : SaveFile = ""
DIM FileExt
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET UF				 = Server.CreateObject("TABSUpload4.Upload")
UF.CodePage			 = "65001"
UF.Start Server.MapPath(D_UPLOAD)

SaveFolder			 = Server.MapPath(D_PARTNERSHIP)


Category			 = sqlFilter(UF("Category"))
Company				 = sqlFilter(UF("Company"))
Contents			 = sqlFilter(UF("Contents"))
Phone				 = sqlFilter(UF("HP1")) & "-" & sqlFilter(UF("HP2")) & "-" & sqlFilter(UF("HP2"))
ReceiveAgree		 = sqlFilter(UF("ReceiveAgree"))
IF ReceiveAgree		 = "" THEN ReceiveAgree = "N"
Email				 = sqlFilter(UF("Email"))
WrtName				 = sqlFilter(UF("WrtName"))
BrandName			 = sqlFilter(UF("BrandName"))
Homepage			 = sqlFilter(UF("Homepage"))
CompanyCate			 = sqlFilter(UF("CompanyCate"))
Distribution		 = sqlFilter(UF("Distribution"))
ProductCode			 = sqlFilter(UF("ProductCode"))
OrderQty			 = sqlFilter(UF("OrderQty"))
NeedDate			 = sqlFilter(UF("NeedDate"))




SET oConn			= ConnectionOpen()							'# 커넥션 생성



oConn.BeginTrans





IF UF("FileName").FileSize > 0 THEN
		'//기본 이미지 저장
		FileExt = Mid(UF("FileName").FileName, Instr(UF("FileName").FileName, ".") + 1)
		SaveFile = UF("FileName").SaveAs(SaveFolder & "\" & U_DATE & U_TIME & right("000" & (timer * 1000) Mod 1000, 3) & "." & FileExt, False)
		SaveFile = UF("FileName").ShortSaveName
END IF



'# 1:1상담 기본정보 저장
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_PartnerShip_Insert"

		.Parameters.Append .CreateParameter("@Category"			, adVarChar		, adParamInput	, 20		, Category)
		.Parameters.Append .CreateParameter("@Company"			, adVarChar		, adParamInput	, 100		, Company)
		.Parameters.Append .CreateParameter("@Contents"			, adLongVarChar	, adParamInput	, 10000000	, Contents)
		.Parameters.Append .CreateParameter("@Phone"			, adVarChar		, adParamInput	, 20		, Phone)
		.Parameters.Append .CreateParameter("@receiveAgree"		, adChar		, adParamInput	, 1			, receiveAgree)
		.Parameters.Append .CreateParameter("@Email"			, adVarChar		, adParamInput	, 50		, Email)
		.Parameters.Append .CreateParameter("@FileName"			, adVarChar		, adParamInput	, 255		, SaveFile)
		.Parameters.Append .CreateParameter("@WrtName"			, adVarChar		, adParamInput	, 50		, WrtName)
		.Parameters.Append .CreateParameter("@BrandName"		, adVarChar		, adParamInput	, 50		, BrandName)
		.Parameters.Append .CreateParameter("@Homepage"			, adVarChar		, adParamInput	, 50		, Homepage)
		.Parameters.Append .CreateParameter("@CompanyCate"		, adVarChar		, adParamInput	, 50		, CompanyCate)
		.Parameters.Append .CreateParameter("@Distribution"		, adVarChar		, adParamInput	, 50		, Distribution)
		.Parameters.Append .CreateParameter("@ProductCode"		, adVarChar		, adParamInput	, 50		, ProductCode)
		.Parameters.Append .CreateParameter("@OrderQty"			, adVarChar		, adParamInput	, 20		, OrderQty)
		.Parameters.Append .CreateParameter("@NeedDate"			, adVarChar		, adParamInput	, 10		, NeedDate)
		.Parameters.Append .CreateParameter("@CreateIP"			, adVarChar		, adParamInput	, 15		, U_IP)


		.Execute, , adExecuteNoRecords

END WITH
SET oCmd = Nothing


IF Err.number<>0 THEN
		oConn.RoolbackTrans
	
		IF SaveFile <> "" THEN
				UF.Delete
		END IF

		SET UF = Nothing : oConn.Close : SET oConn = Nothing

		Response.Write "FAIL|||||입력 중 에러가 발생하였습니다."
		Response.End
END IF



oConn.CommitTrans


SET UF = Nothing
oConn.Close
SET oConn = Nothing


Response.Write "OK|||||"
%>
