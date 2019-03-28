<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ReviewWriteOk.asp - 상품후기 등록 처리
'Date		: 2018.12.20
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
IF U_NUM = "" AND N_NAME = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다. 다시 로그인하여 주십시오."
		Response.End
END IF

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

DIM OrderCode
DIM Idx
DIM OPIdx_Org
DIM ProductCode
DIM SizeGrade
DIM WearGrade
DIM DesignGrade
DIM QualityGrade
DIM Contents
DIM UploadFiles
DIM arrUploadFiles

DIM ReviewIdx
DIM MemberNum
DIM ReviewType
DIM ReviewPoint

DIM FSO
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode			= sqlFilter(Request("OrderCode"))
Idx					= sqlFilter(Request("Idx"))
OPIdx_Org			= sqlFilter(Request("OPIdx_Org"))
ProductCode			= sqlFilter(Request("ProductCode"))
SizeGrade			= sqlFilter(Request("SizeGrade"))
WearGrade			= sqlFilter(Request("WearGrade"))
DesignGrade			= sqlFilter(Request("DesignGrade"))
QualityGrade		= sqlFilter(Request("QualityGrade"))
Contents			= sqlFilter(Request("Contents"))
UploadFiles			= sqlFilter(Request("UploadFiles"))


IF OrderCode = "" OR Idx = "" OR ProductCode = "" OR Contents = "" THEN
		Response.Write "FAIL|||||후기등록할 입력정보가 부족합니다."
		Response.End
END IF


IF OPIdx_Org		= "" THEN OPIdx_Org			= "0"
IF SizeGrade		= "" THEN SizeGrade			= "0"
IF WearGrade		= "" THEN WearGrade			= "0"
IF DesignGrade		= "" THEN DesignGrade		= "0"
IF QualityGrade		= "" THEN QualityGrade		= "0"


IF U_NUM <> "" THEN
		MemberNum	= U_NUM
ELSE
		MemberNum	= "0"
END IF

IF UploadFiles <> "" THEN
		ReviewType	= "P"		'# 포토상품후기
		ReviewPoint	= MALL_REVIEW_POINT_P
ELSE
		ReviewType	= "B"		'# 일반상품후기
		ReviewPoint	= MALL_REVIEW_POINT_B
END IF


SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'-----------------------------------------------------------------------------------------------------------'
'# 상품후기 등록여부 체크 Start
'-----------------------------------------------------------------------------------------------------------'
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Review_Select_By_Order_Product_Idx"

		.Parameters.Append .CreateParameter("@Order_Product_Idx",	 adInteger, adParaminput, 	, Idx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||이미 상품후기를 등록하셨습니다."
		Response.End
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'# 상품후기 등록여부 체크 End
'-----------------------------------------------------------------------------------------------------------'


'ON ERROR RESUME NEXT



oConn.BeginTrans



'-----------------------------------------------------------------------------------------------------------'	
'# 상품후기 등록 Start
'-----------------------------------------------------------------------------------------------------------'	
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Review_Insert"

		.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,     ,	 MemberNum)
		.Parameters.Append .CreateParameter("@Order_Product_Idx",	adInteger,	adParamInput,     ,	 Idx)
		.Parameters.Append .CreateParameter("@ProductCode",			adInteger,	adParamInput,     ,	 ProductCode)
		.Parameters.Append .CreateParameter("@ReviewType",			adChar,		adParamInput,    1,	 ReviewType)
		.Parameters.Append .CreateParameter("@SizeGrade",			adInteger,	adParamInput,     ,	 SizeGrade)
		.Parameters.Append .CreateParameter("@WearGrade",			adInteger,	adParamInput,     ,	 WearGrade)
		.Parameters.Append .CreateParameter("@DesignGrade",			adInteger,	adParamInput,     ,	 DesignGrade)
		.Parameters.Append .CreateParameter("@QualityGrade",		adInteger,	adParamInput,     ,	 QualityGrade)
		.Parameters.Append .CreateParameter("@Contents",			adVarWChar,	adParamInput, 8000,	 Contents)
		.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,   20,	 U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,   15,	 U_IP)
		.Parameters.Append .CreateParameter("@Idx",					adInteger,	adParamOutput	 )

		.Execute, , adExecuteNoRecords

		ReviewIdx = .Parameters("@Idx").Value
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		oConn.RollbackTrans

		oRs.Close
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||상품후기 등록 중 오류가 발생하였습니다.[1]"
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'	
'# 상품후기 등록 End
'-----------------------------------------------------------------------------------------------------------'	



'-----------------------------------------------------------------------------------------------------------'	
'# 상품후기 이미지 등록 Start
'-----------------------------------------------------------------------------------------------------------'	
IF UploadFiles <> "" THEN
		SET FSO = Server.CreateObject("Scripting.FileSystemObject")

		arrUploadFiles		= Split(UploadFiles, "|||||")

		FOR i = 0 TO UBOUND(arrUploadFiles)
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Product_Review_Image_Insert"

						.Parameters.Append .CreateParameter("@ReviewIdx",			adInteger,	adParamInput,     ,	 ReviewIdx)
						.Parameters.Append .CreateParameter("@FileName",			adVarChar,	adParamInput,  255,	 arrUploadFiles(i))
						.Parameters.Append .CreateParameter("@Contents",			adVarWChar,	adParamInput, 8000,	 "")

						.Execute, , adExecuteNoRecords
				END WITH
				SET oCmd = Nothing

				IF Err.Number <> 0 THEN
						oConn.RollbackTrans

						oRs.Close
						SET oRs = Nothing
						oConn.Close
						SET oConn = Nothing

						Response.Write "FAIL|||||상품후기 등록 중 오류가 발생하였습니다.[2]"
						Response.End
				END IF

				IF FSO.FileExists(Server.MapPath(D_REVIEW & "Temp/" & arrUploadFiles(i))) THEN
						FSO.MoveFile Server.MapPath(D_REVIEW & "Temp/" & arrUploadFiles(i)), Server.MapPath(D_REVIEW & arrUploadFiles(i))
				END IF
		NEXT

		SET FSO = Nothing
END IF
'-----------------------------------------------------------------------------------------------------------'	
'# 상품후기 이미지 등록 End
'-----------------------------------------------------------------------------------------------------------'	


IF U_MFLAG = "Y" THEN
		'-----------------------------------------------------------------------------------------------------------'	
		'# 상품후기 작성 포인트 Start
		'-----------------------------------------------------------------------------------------------------------'	
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_Point_Insert"

				.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,     ,	 MemberNum)
				.Parameters.Append .CreateParameter("@PCode",				adChar,		adParamInput,    3,	 "201")
				.Parameters.Append .CreateParameter("@AddPoint",			adCurrency,	adParamInput,     ,	 ReviewPoint)
				.Parameters.Append .CreateParameter("@Memo",				adVarChar,	adParamInput,  300,	 "상품후기작성")
				.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,   20,	 OrderCode)
				.Parameters.Append .CreateParameter("@OPIdx_Org",			adInteger,	adParamInput,     ,	 OPIdx_Org)
				.Parameters.Append .CreateParameter("@AvailableDT",			adVarChar,	adParamInput,   10,	 DateAdd("yyyy", 1, Date))		'# 실제 만기일자는 프로시져에서 다시 계산한다.
				.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,   20,	 U_NUM)
				.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,   15,	 U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing

		IF Err.Number <> 0 THEN
				oConn.RollbackTrans

				SET oRs = Nothing
				oConn.Close
				SET oConn = Nothing

				Response.Write "FAIL|||||상품후기 등록 중 오류가 발생하였습니다.[3]"
				Response.End
		END IF
		'-----------------------------------------------------------------------------------------------------------'	
		'# 상품후기 작성 포인트 End
		'-----------------------------------------------------------------------------------------------------------'	
END IF


oConn.CommitTrans



Response.Write "OK|||||"

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>