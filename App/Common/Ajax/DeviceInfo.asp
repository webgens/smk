<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'DeviceInfo.asp - 앱 최초 실행시 디바이스 정보 insert / update
'Date		: 2018.12.06
'Update		: 
'Writer		: Hong
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드 ---------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "00"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
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

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절
DIM sqlQuery					'# SQL 문
	
DIM i
DIM j
DIM x
DIM y
	
DIM InstallationId				'# 디바이스 고유번호
DIM DeviceToken					'# 푸쉬를 보내기 위한 토큰값
DIM DeviceType					'# android / ios
DIM AppVersion					'# 현재 앱버젼
DIM AppModelName				'# 모델

DIM ObjSC
DIM MemberNum
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


InstallationId		= sqlFilter(Request("installationId"))
DeviceToken			= sqlFilter(Request("deviceToken"))
DeviceType			= sqlFilter(Request("deviceType"))
AppVersion			= sqlFilter(Request("appVersion"))
AppModelName		= sqlFilter(Request("appModelName"))



' INSTALLATIONID 쿠키저장
Response.Cookies("U_DEVICEID")			 = Encrypt(installationId)
Response.Cookies("U_DEVICE")			 = Encrypt(DeviceType)
Response.Cookies("U_PUSHKEY")			 = Encrypt(DeviceToken)
Response.Cookies("U_MODELNAME")			 = Encrypt(AppModelName)
Response.Cookies("U_APPVERSION")		 = Encrypt(AppVersion)


IF U_NUM = "" THEN
		MemberNum = 0
ELSE
		MemberNum = U_NUM
END IF

SET oConn			 = ConnectionOpen()							'# 커넥션 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_App_Device_Check"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput,   20,	 MemberNum)					'# 회원고유번호
		.Parameters.Append .CreateParameter("@DeviceID",	 adVarChar, adParamInput,  200,	 InstallationId)			'# 디바이스 고유번호
		.Parameters.Append .CreateParameter("@DeviceType",	 adVarChar, adParamInput,   10,	 DeviceType)				'# DeviceType (iphone, android)
		.Parameters.Append .CreateParameter("@ModelName",	 adVarChar, adParamInput,  255,	 AppModelName)				'# ModelName
		.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar, adParamInput,   15,	 U_IP)
		.Execute, , adExecuteNoRecords
END WITH	
SET oCmd = Nothing


oConn.Close
SET oConn = Nothing


Response.Write "OK|||||"
%>