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
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

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

Dim Category1
Dim Category2
Dim Category3
Dim Category4
Dim Title
Dim Contents
Dim SMSReturnFlag
Dim Mobile
Dim EMailReturnFlag
Dim EMail

Dim MtmIdx

Dim UploadFiles
Dim UploadFilesCount
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Category1			= sqlFilter(Request("Category1"))
Category2			= sqlFilter(Request("Category2"))
Category3			= sqlFilter(Request("Category3"))
Category4			= sqlFilter(Request("Category4"))
Title				= sqlFilter(Request("Title"))
Contents			= sqlFilter(Request("Contents"))
SMSReturnFlag		= sqlFilter(Request("SMSReturnFlag"))
IF SMSReturnFlag	= "1" THEN
Mobile				= ChgTel(TRIM(sqlFilter(Request("Mobile1"))) & TRIM(sqlFilter(Request("Mobile23"))))
ELSE
Mobile				= ""
END IF
EMailReturnFlag		= sqlFilter(Request("EMailReturnFlag"))
IF EMailReturnFlag	= "1" THEN
EMail				= sqlFilter(Request("EMail"))
ELSE
EMail				= ""
END IF
UploadFiles			= sqlFilter(Request("UploadFiles"))
UploadFilesCount	= sqlFilter(Request("UploadFilesCount"))


SET oConn			= ConnectionOpen()							'# 커넥션 생성

oConn.BeginTrans

'# 1:1상담 기본정보 저장
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_QNA_INSERT"

		.Parameters.Append .CreateParameter("@MemberNum"		, adInteger	, adParamInput	, 		, U_NUM)
		.Parameters.Append .CreateParameter("@Category1"		, adVarChar	, adParamInput	, 50	, Category1)
		.Parameters.Append .CreateParameter("@Category2"		, adVarChar	, adParamInput	, 50	, Category2)
		.Parameters.Append .CreateParameter("@Category3"		, adVarChar	, adParamInput	, 50	, Category3)
		.Parameters.Append .CreateParameter("@Category4"		, adVarChar	, adParamInput	, 50	, Category4)
		.Parameters.Append .CreateParameter("@Title"			, adVarChar	, adParamInput	, 255	, Title)
		.Parameters.Append .CreateParameter("@Contents"			, adVarChar	, adParamInput	, 8000	, Contents)
		.Parameters.Append .CreateParameter("@SMSReturnFlag"	, adInteger	, adParamInput	,		, SMSReturnFlag)
		.Parameters.Append .CreateParameter("@Mobile"			, adVarChar	, adParamInput	, 15	, Mobile)
		.Parameters.Append .CreateParameter("@EMailReturnFlag"	, adInteger	, adParamInput	,		, EMailReturnFlag)
		.Parameters.Append .CreateParameter("@EMail"			, adVarChar	, adParamInput	, 50	, EMail)
		.Parameters.Append .CreateParameter("@CreateID"			, adVarChar	, adParamInput	, 30	, U_ID)
		.Parameters.Append .CreateParameter("@CreateIP"			, adVarChar	, adParamInput	, 15	, U_IP)
		.Parameters.Append .CreateParameter("@Idx"				, adInteger , adParamOutput	)

		.Execute, , adExecuteNoRecords

		MtmIdx = .Parameters("@Idx").Value
END WITH
SET oCmd = Nothing


'response.write UploadFiles &"<br>"& UploadFilesCount &"<br>"& ubound(SPLIT(UploadFiles,"|||||")) &"<br>"
'# 1:1상담 첨부파일 저장
IF UploadFilesCount>=1 THEN
	Dim FSO
	SET FSO = Server.CreateObject("Scripting.FileSystemObject")
	IF UploadFilesCount=1 THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_QNA_ATTACHFILE_INSERT"

				.Parameters.Append .CreateParameter("@QNA_IDX"			, adInteger	, adParamInput	, 		, MtmIdx)
				.Parameters.Append .CreateParameter("@UploadFile"		, adVarChar	, adParamInput	, 100	, UploadFiles)

				.Execute, , adExecuteNoRecords

		END WITH
		SET oCmd = Nothing
	ELSEIF UploadFilesCount>1 THEN
		Dim UploadFilesArr : UploadFilesArr = SPLIT(UploadFiles,"|||||")
	
		FOR	i = 0 TO UBOUND(UploadFilesArr)
			SET oCmd = Server.CreateObject("ADODB.Command")
			WITH oCmd
					.ActiveConnection	 = oConn
					.CommandType		 = adCmdStoredProc
					.CommandText		 = "USP_Front_EShop_QNA_ATTACHFILE_INSERT"

					.Parameters.Append .CreateParameter("@QNA_IDX"			, adInteger	, adParamInput	, 		, MtmIdx)
					.Parameters.Append .CreateParameter("@UploadFile"		, adVarChar	, adParamInput	, 100	, UploadFilesArr(i))

					.Execute, , adExecuteNoRecords

			END WITH
			SET oCmd = Nothing
			IF FSO.FileExists(Server.MapPath(D_MTMQNA & "Temp/" & UploadFilesArr(i))) THEN
					FSO.MoveFile Server.MapPath(D_MTMQNA & "Temp/" & UploadFilesArr(i)), Server.MapPath(D_MTMQNA & UploadFilesArr(i))
			END IF
		NEXT
	END IF
	SET FSO = Nothing
END IF

IF Err.number<>0 THEN
	oConn.RoolbackTrans
	oConn.Close : SET oConn = Nothing
	Response.Write "FAIL|||||입력도 중 에러가 발생하였습니다."
	Response.End
END IF

oConn.CommitTrans

oConn.Close
SET oConn = Nothing


Response.Write "OK|||||등록처리가 완료되었습니다."
%>