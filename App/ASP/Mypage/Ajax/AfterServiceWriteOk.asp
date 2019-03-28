<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MtmQnaWriteOk.asp - 1:1상담 글등록
'Date		: 2018.12.28
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

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderCode
DIM OrderProductIDX
DIM ProductCode
DIM ShopCD
DIM RequestCode
DIM StateCode : StateCode = 0
DIM Contents
DIM UploadFiles
DIM UploadFilesCount

DIM MemberNum

DIM AsIdx
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

IF U_NUM = "" AND N_NAME = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다. 다시 로그인하여 주십시오."
		Response.End
END IF


OrderCode			= sqlFilter(Request("OrderCode"))
OrderProductIDX		= sqlFilter(Request("OrderProductIDX"))
ProductCode			= sqlFilter(Request("ProductCode"))
ShopCD				= sqlFilter(Request("ShopCD"))
RequestCode			= sqlFilter(Request("RequestCode"))
Contents			= sqlFilter(Request("Contents"))
UploadFiles			= sqlFilter(Request("UploadFiles"))
UploadFilesCount	= sqlFilter(Request("UploadFilesCount"))


IF OrderCode = "" OR OrderProductIDX = "" OR ProductCode = "" OR ShopCD = "" THEN
		Response.Write "FAIL|||||A/S정보가 부족합니다."
		Response.End
END IF



IF U_NUM <> "" THEN
		MemberNum	= U_NUM
ELSE
		MemberNum	= ""
END IF



SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'-----------------------------------------------------------------------------------------------------------'
'# A/S 진행여부 체크 Start
'-----------------------------------------------------------------------------------------------------------'
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_AfterService_Select_For_Count"

		.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,   20,	 OrderCode)
		.Parameters.Append .CreateParameter("@Order_Product_Idx",	adInteger,	adParamInput,     ,	 OrderProductIDX)
		.Parameters.Append .CreateParameter("@ProductCode",			adInteger,	adParamInput,     ,	 ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
	IF oRs(0) > 0 THEN
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||이미 진행 중인 A/S건이 있습니다."
		Response.End
	END IF
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'# A/S 등록여부 체크 End
'-----------------------------------------------------------------------------------------------------------'


'# 이미지 자르기 함수
SUB setImageInfo(ByVal needWidth, ByVal needHeight, ByVal width, ByVal height, ByRef thumbFlag, ByRef thumbWidth, ByRef thumbHeight, ByRef cropX, ByRef cropY, ByRef cropWidth, ByRef cropHeight)
	DIM tempHeight
	DIM topSpace
	DIM tempWidth
	DIM leftSpace

	thumbFlag	 = "N"
	thumbWidth	 = 0
	thumbHeight	 = 0

	IF CDbl(height / width) >= CDbl(needHeight / needWidth) THEN
		IF CDbl(width) >= CDbl(needWidth) THEN
			thumbFlag	 = "Y"
			tempHeight	 = ROUND(needWidth * height /  width)
			thumbWidth	 = needWidth
			thumbHeight	 = tempHeight
			topSpace		 = ABS(ROUND((tempHeight - needHeight) / 3) -1)
			cropX			 = 0
			cropY			 = topSpace
			cropWidth	 = needWidth
			cropHeight	 = needHeight
		ELSE
			thumbFlag	 = "N"
			tempHeight	 = ROUND(needWidth * height /  width)
			topSpace		 = ABS(ROUND((tempHeight - needHeight) / 3) -1)
			cropX			 = 0
			cropY			 = topSpace
			cropWidth	 = width
			cropHeight	 = ROUND(width * needHeight / needWidth)
		END IF
	ELSE
		IF CDbl(height) >= CDbl(needHeight) THEN
			thumbFlag	 = "Y"
			tempWidth	 = ROUND(needHeight * width /  height)
			thumbWidth	 = tempWidth
			thumbHeight	 = needHeight
			leftSpace		 = ABS(ROUND((tempWidth - needWidth) / 2) -1)
			cropX			 = leftSpace
			cropY			 = 0
			cropWidth	 = needWidth
			cropHeight	 = needHeight
		ELSE
			thumbFlag	 = "N"
			tempWidth	 = width
			leftSpace		 = ABS(ROUND((tempWidth - (needWidth * height / needHeight)) / 2) -1)
			cropX			 = leftSpace
			cropY			 = 0
			cropWidth	 = ROUND(height * needWidth / needHeight)
			cropHeight	 = height
		END IF
	END IF
END SUB

'ON ERROR RESUME NEXT



oConn.BeginTrans



'-----------------------------------------------------------------------------------------------------------'	
'# A/S 등록 Start
'-----------------------------------------------------------------------------------------------------------'	
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_AfterService_Insert"

		.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,   20,	 OrderCode)
		.Parameters.Append .CreateParameter("@Order_Product_Idx",	adInteger,	adParamInput,     ,	 OrderProductIDX)
		.Parameters.Append .CreateParameter("@ProductCode",			adInteger,	adParamInput,     ,	 ProductCode)
		.Parameters.Append .CreateParameter("@ShopCD",				adVarChar,	adParamInput,   10,	 ShopCD)
		.Parameters.Append .CreateParameter("@RequestCode",			adChar,		adParamInput,    1,	 RequestCode)
		.Parameters.Append .CreateParameter("@StateCode",			adInteger,	adParamInput,     ,	 StateCode)
		.Parameters.Append .CreateParameter("@Contents",			adVarWChar,	adParamInput, 8000,	 Contents)
		.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,   30,	 MemberNum)
		.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,   15,	 U_IP)
		.Parameters.Append .CreateParameter("@Idx",					adInteger,	adParamOutput	 )

		.Execute, , adExecuteNoRecords

		AsIdx = .Parameters("@Idx").Value
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		oConn.RollbackTrans

		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||A/S 등록 중 오류가 발생하였습니다.[1]"
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'	
'# A/S 등록 End
'-----------------------------------------------------------------------------------------------------------'	



'-----------------------------------------------------------------------------------------------------------'	
'# A/S 이미지 등록 Start
'-----------------------------------------------------------------------------------------------------------'	
IF UploadFiles <> "" THEN
	DIM FSO
	DIM arrUploadFiles

	SET FSO = Server.CreateObject("Scripting.FileSystemObject")

	arrUploadFiles		= Split(UploadFiles, "|||||")

	FOR i = 0 TO UBOUND(arrUploadFiles)
			SET oCmd = Server.CreateObject("ADODB.Command")
			WITH oCmd
					.ActiveConnection	 = oConn
					.CommandType		 = adCmdStoredProc
					.CommandText		 = "USP_Front_EShop_Order_AfterService_Image_Insert"

					.Parameters.Append .CreateParameter("@AIdx",				adInteger,	adParamInput,     ,	 AsIdx)
					.Parameters.Append .CreateParameter("@FileName",			adVarChar,	adParamInput,  255,	 arrUploadFiles(i))

					.Execute, , adExecuteNoRecords
			END WITH
			SET oCmd = Nothing


			IF FSO.FileExists(Server.MapPath(D_ORDERAS & "Temp/" & arrUploadFiles(i))) THEN
					FSO.MoveFile Server.MapPath(D_ORDERAS & "Temp/" & arrUploadFiles(i)), Server.MapPath(D_ORDERAS & arrUploadFiles(i))
			END IF


			IF Err.Number <> 0 THEN
					oConn.RollbackTrans

					SET oRs = Nothing
					oConn.Close
					SET oConn = Nothing

					Response.Write "FAIL|||||A/S 등록 중 오류가 발생하였습니다."
					Response.End
			END IF

	NEXT

	SET FSO = Nothing


	IF Err.Number <> 0 THEN
			oConn.RollbackTrans

			SET oRs = Nothing
			oConn.Close
			SET oConn = Nothing

			Response.Write "FAIL|||||A/S 등록 중 오류가 발생하였습니다.[2]"
			Response.End
	END IF

END IF
'-----------------------------------------------------------------------------------------------------------'	
'# A/S 이미지 등록 End
'-----------------------------------------------------------------------------------------------------------'	



'-----------------------------------------------------------------------------------------------------------'	
'# A/S Log 등록 Start
'-----------------------------------------------------------------------------------------------------------'	
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_AfterService_Log_Insert"

		.Parameters.Append .CreateParameter("@AIdx",				adInteger,	adParamInput,     ,	 AsIdx)
		.Parameters.Append .CreateParameter("@StateCode",			adInteger,	adParamInput,     ,	 StateCode)
		.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,   30,	 MemberNum)
		.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,   15,	 U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing


IF Err.Number <> 0 THEN
		oConn.RollbackTrans

		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||A/S 등록 중 오류가 발생하였습니다.[3]"
		Response.End
END IF

'-----------------------------------------------------------------------------------------------------------'	
'# A/S Log 등록 End
'-----------------------------------------------------------------------------------------------------------'	







oConn.CommitTrans



Response.Write "OK|||||"

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>
