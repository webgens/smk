<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'NewAgreementOk.asp - 신규약관 동의 처리
'Date		: 2018.12.19
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드----------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "01"
PageCode2 = "04"
PageCode3 = "02"
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

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절


DIM Agr1
DIM Agr2
DIM Agr3
DIM Agr4
DIM Agr5
DIM Agr6
DIM Agr7

DIM Agreement
DIM MemberNum
DIM ProgID
DIM SaveID
DIM OrderFlag

DIM DB_Name
DIM DB_GroupCode


DIM CouponName
DIM CouponCount : CouponCount = 0

DIM StartDT
DIM EndDT

DIM Message
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


Agr1			 = sqlFilter(Request("Agr1"))
Agr2			 = sqlFilter(Request("Agr2"))
Agr3			 = sqlFilter(Request("Agr3"))
Agr4			 = sqlFilter(Request("Agr4"))
Agr5			 = sqlFilter(Request("Agr5"))
Agr6			 = sqlFilter(Request("Agr6"))
Agr7			 = sqlFilter(Request("Agr7"))

IF Agr1			 = "" THEN Agr1 = "N"
IF Agr2			 = "" THEN Agr2 = "N"
IF Agr3			 = "" THEN Agr3 = "N"
IF Agr4			 = "" THEN Agr4 = "N"
IF Agr5			 = "" THEN Agr5 = "N"
IF Agr6			 = "" THEN Agr6 = "N"
IF Agr7			 = "" THEN Agr7 = "N"


IF Agr1 = "N" THEN
		Response.Write "FAIL|||||쇼핑몰 이용약관에 동의 하셔야 됩니다."
		Response.End
END IF
IF Agr2 = "N" THEN
		Response.Write "FAIL|||||개인정보 이용 및 수집에 대해 동의 하셔야 됩니다."
		Response.End
END IF


Agreement		 = Agr1 & "|" & Agr2 & "|" & Agr3 & "|" & Agr4 & "|" & Agr5 & "|" & Agr6 & "|" & Agr7
MemberNum		 = TRIM(Decrypt(Request.Cookies("TEMP_UNUM")))
ProgID			 = TRIM(Decrypt(Request.Cookies("TEMP_PROGID")))
SaveID			 = TRIM(Decrypt(Request.Cookies("TEMP_UIDSAVE")))
OrderFlag		 = TRIM(Decrypt(Request.Cookies("TEMP_ORDERFLAG")))
IF ProgID		 = "" THEN ProgID = "/"
IF MemberNum = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다.<br />다시 로그인하여 주십시오."
		Response.End
END IF



SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성





'# 회원 정보 - 아이디 찾기
SET oCmd = SErver.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput, , MemberNum)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		DB_Name				 = oRs("Name")
		DB_GroupCode		 = oRs("GroupCode")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||회원 정보가 없습니다. 다시 로그인 하여 주십시오."
		Response.End
END IF
oRs.Close




'# 신규 약관 동의 여부
SET oCmd = SErver.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_NewAgreement_Insert"
	
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput,   , MemberNum)
		.Parameters.Append .CreateParameter("@Agreement",	 adVarChar, adParamInput, 20, Agreement)
		.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	adParamInput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing



'# 신규 약관 동의에 따른 회원정보 수정
SET oCmd = SErver.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_NewAgreement_Update"
	
		.Parameters.Append .CreateParameter("@MemberNum"	, adInteger	, adParamInput,   , MemberNum)
		.Parameters.Append .CreateParameter("@FTFlag"		, adChar	, adParamInput, 1 , Agr5)
		.Parameters.Append .CreateParameter("@SmsFlag"		, adChar	, adParamInput, 1 , Agr6)
		.Parameters.Append .CreateParameter("@EmailFlag"	, adChar	, adParamInput, 1 , Agr7)
		.Parameters.Append .CreateParameter("@UpdateIP"		, adVarChar	, adParamInput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing












'# 배포된 쿠폰 발급 받기 - EShop_Coupon_Member 에 배포되었지만 아직 안받은 쿠폰
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Member_Select_For_Not_Receive"
	
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		CouponName	 = oRs("CouponName")
		CouponCount	 = CouponCount + oRs.RecordCount

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Member_Update_For_Receive"
	
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
				.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar,	 adParamInput, 20, MemberNum)
				.Parameters.Append .CreateParameter("@UpdateIP",	 adVarChar,	 adParamInput, 15, U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
END IF
oRs.Close



'# 배포될 쿠폰 발급 받기 - EShop_Coupon_Member 에 데이터가 없고 EShop_Coupon 에 내가 배포 대상인 쿠폰
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Select_For_Not_Receive"
	
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
		.Parameters.Append .CreateParameter("@GroupCode",	 adVarChar,	 adParamInput, 10, DB_GroupCode)
	
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF CouponName = "" THEN CouponName	 = oRs("CouponName")
		CouponCount	 = CouponCount + oRs.RecordCount


		Do Until oRs.EOF
				IF oRs("UseDateType") = "U" THEN
						StartDT	 = U_DATE & R_HOUR & "0000"
						EndDT	 = "99999999999999"
				ELSEIF oRs("UseDateType") = "P" THEN
						StartDT	 = U_DATE & R_HOUR & "0000"
						EndDT	 = oRs("UseEDate")
				ELSE
						StartDT	 = U_DATE & "000000"
						EndDT	 = REPLACE(DATEADD("d", oRs("UseDay"), Date), "-", "") & "240000"
				END IF

						
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Coupon_Member_Insert"
	
						.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
						.Parameters.Append .CreateParameter("@CouponIdx",	 adBigInt,	 adParamInput,   , oRs("Idx"))
						.Parameters.Append .CreateParameter("@StartDT",		 adVarChar,	 adParamInput, 14, StartDT)
						.Parameters.Append .CreateParameter("@EndDT",		 adVarChar,	 adParamInput, 14, EndDT)
						.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput, 20, MemberNum)
						.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput, 15, U_IP)

						.Execute, , adExecuteNoRecords
				END WITH
				SET oCmd = Nothing

				oRs.MoveNext
		Loop

END IF
oRs.Close


IF CouponCount > 0 THEN
	
		Message = ""
		Message = Message & "			<div class=""area-pop"">"&vbLf
		Message = Message & "				<div class=""alert"">"&vbLf
		Message = Message & "					<div class=""tit-pop"">"&vbLf
		Message = Message & "						<p class=""tit"" id=""confirm_title"">SHOEMARKER</p>"&vbLf
		Message = Message & "						<button id=""confirm_close"" onclick=""closePop('messagePop')"" class=""btn-hide-pop"">닫기</button>"&vbLf
		Message = Message & "					</div>"
		Message = Message & "					<div class=""container-pop"">"&vbLf
		Message = Message & "						<div class=""contents"">"&vbLf
		Message = Message & "							<div class=""ly-cont"">"&vbLf
		Message = Message & "								<p id=""confirm_content"" class=""t-level4"" style=""text-align:left"">"&vbLf
	
		Message = Message &									"신규약관 동의 완료 되었습니다.<br>"
		Message = Message &									DB_Name & "님께<br>"
															IF CouponCount = 1 THEN
		Message = Message &											"[" & CouponName & "] 쿠폰이 발급 되었습니다."
															ELSE
		Message = Message &											"[" & CouponName & "] 외 " & CouponCount - 1 & "장이 발급 되었습니다."
															END IF

		Message = Message & "								</p>"&vbLf
		Message = Message & "							</div>"&vbLf
		Message = Message & "						</div>"&vbLf
		Message = Message & "						<div class=""btns"">"&vbLf
		IF OrderFlag <> "Y" THEN
		Message = Message & "							<button type=""button"" id=""message_btn1"" onclick=""location.href='/ASP/Mypage/CouponList.asp'"" class=""button ty-black"">쿠폰 확인</button>"
		END IF
		Message = Message & "							<button type=""button"" id=""message_btn2"" onclick=""location.href='" & ProgID & "'"" class=""button ty-red"">확인</button>"
		Message = Message & "						</div>"&vbLf
		Message = Message & "					</div>"&vbLf
		Message = Message & "				</div>"&vbLf
		Message = Message & "			</div>"&vbLf

END IF


















'# 로그인 정보 입력
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Login_Insert"
	
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
		.Parameters.Append .CreateParameter("@Location",	 adChar,	 adParamInput,  1, "P")
		.Parameters.Append .CreateParameter("@LoginIP",		 adVarChar,	 adParamInput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing



	
SET oRs = Nothing : oConn.Close : SET oConn = Nothing

	
Response.Cookies("UIP")				 = Request.Cookies("TEMP_UIP")
Response.Cookies("UMFLAG")			 = Request.Cookies("TEMP_MFLAG")
Response.Cookies("UNUM")			 = Request.Cookies("TEMP_UNUM")
Response.Cookies("UID")				 = Request.Cookies("TEMP_UID")
Response.Cookies("UNAME")			 = Request.Cookies("TEMP_UNAME")
Response.Cookies("UEFLAG")			 = Request.Cookies("TEMP_EFLAG")
Response.Cookies("UETYPE")			 = Request.Cookies("TEMP_ETYPE")
Response.Cookies("UGROUP")			 = Request.Cookies("TEMP_UGROUP")
Response.Cookies("UGNAME")			 = Request.Cookies("TEMP_UGNAME")


'# 아이디 저장
IF SaveID = "Y" THEN
		Response.Cookies("SMEM_ID")			 = Request.Cookies("TEMP_UID")
		Response.Cookies("SMEM_ID").Expires	 = Date() + 1000
END IF

	
Response.Cookies("TEMP_DOR")		 = ""
Response.Cookies("TEMP_NEW")		 = ""
Response.Cookies("TEMP_MFLAG")		 = ""
Response.Cookies("TEMP_PROGID")		 = ""
Response.Cookies("TEMP_UIDSAVE")	 = ""
Response.Cookies("TEMP_UIP")		 = ""
Response.Cookies("TEMP_UNUM")		 = ""
Response.Cookies("TEMP_UID")		 = ""
Response.Cookies("TEMP_UNAME")		 = ""
Response.Cookies("TEMP_EFLAG")		 = ""
Response.Cookies("TEMP_ETYPE")		 = ""
Response.Cookies("TEMP_UGROUP")		 = ""
Response.Cookies("TEMP_UGNAME")		 = ""
Response.Cookies("TEMP_ORDERFLAG")	 = ""




Response.Write "OK|||||" & ProgID & "|||||" & Message
Response.End
%>