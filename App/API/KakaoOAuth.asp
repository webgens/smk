<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'GoogleOAuth.asp - 카카오 콜백 페이지
'Date		: 2019.01.08
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<!-- #include virtual="/API/json_for_asp/aspJSON1.17.asp" -->

<!-- #include virtual="/INC/Header.asp" -->

<script language="javascript" runat="server" charset="utf-8">
	function replaceAll(strTemp, strValue1, strValue2){ 
		while (1) {
			if (strTemp.indexOf(strValue1) != -1)
				strTemp = strTemp.replace(strValue1, strValue2);
			else
				break;
		}
		return strTemp;
	}
	function unicodeToKor(a) {
		var str = a;
		return unescape(replaceAll(str, "\\", "%"));
	}
</script>
<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절
DIM sqlQuery





DIM MemberNum
DIM UserID
DIM DelFlag
DIM SNSChangeFlag
DIM MemberFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



Dim ResponseText
Dim HTTP_Object

Set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
With HTTP_Object
		'API 통신 Timeout 을 30초로 지정
		.SetTimeouts 30000, 30000, 30000, 30000
		.Open "POST", "https://kauth.kakao.com/oauth/token", False
		.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		.Send "grant_type=authorization_code&client_id="& KAKAO_LOGIN_CLIENTID &"&code="&request("code")
		.WaitForResponse

		If .Status = 200 Then
				ResponseText = .ResponseText
		Else
				ResponseText = ""
		End If
End With


Dim access_token
Dim access_token_secret
Dim refresh_token
Dim token_type
Dim expires_in
Dim state
Dim Errorcode
Dim errorTxt

Dim Read_Data
If ResponseText <> "" Then
		Set Read_Data = New aspJSON
		Read_Data.loadJSON(ResponseText)
		With Read_Data
				access_token = .data("access_token")
				refresh_token = .data("refresh_token")
				token_type = .data("token_type")
				expires_in = .data("expires_in")
'				Errorcode = .data("error")
'				errorTxt = .data("error_description")
		End With
End If

Dim resultCode
DIM ID
DIM Email
DIM KName
DIM NickName


If access_token <> "" Then

		With HTTP_Object
				'API 통신 Timeout 을 30초로 지정
				.SetTimeouts 30000, 30000, 30000, 30000
				.Open "GET", "https://kapi.kakao.com/v1/user/me", False
				.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
				.SetRequestHeader "Authorization", "Bearer " & access_token
				.Send
				.WaitForResponse

				If .Status = 200 Then
						ResponseText = .ResponseText
				Else
						ResponseText = ""
				End If
		End With


		If ResponseText <> "" Then
				Set Read_Data = New aspJSON
				Read_Data.loadJSON(ResponseText)
				With Read_Data
						ID		 = .data("id")
						IF .data("kaccount_email") <> "" THEN
								Email	 = .data("kaccount_email")
						ELSE
								Email	 = ""
						END IF
						KName	 = .data("properties").item("nickname")
				End With
		End If



		SET oConn		 = ConnectionOpen()							'# 커넥션 생성
		SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


		'# SNS ID 체크
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_SNS_Select_By_SNSID"
	
				.Parameters.Append .CreateParameter("@SNSKind",	 adChar,	 adParamInput,  1, "K")
				.Parameters.Append .CreateParameter("SNSID",	 adVarChar,	 adParamInput, 50, ID)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				MemberNum = oRs("MemberNum")
		ELSE
				'// SNS계정연결에 사용
				IF U_NUM <> "" THEN
					MemberNum = U_NUM
				ELSE
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing

						Response.Cookies("SNS_UID")		= Encrypt(ID)
						Response.Cookies("SNS_Email")	= Encrypt(Email)
						Response.Cookies("SNS_KName")	= Encrypt(KName)
						Response.Cookies("SNS_Kind")		= Encrypt("K")


						Response.Redirect "/ASP/Member/SnsGate.asp"
						Response.End
				END IF
		END IF
		oRs.Close



		'# 회원정보 체크
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_EShop_Member_Select_By_MemberNum"
	
				.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput,  , MemberNum)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				UserID			 = oRs("UserID")
				DelFlag			 = oRs("DelFlag")
				SNSChangeFlag	 = oRs("SNSChangeFlag")
				MemberFlag		 = oRs("MemberFlag")

				Response.Cookies("SNS_UID")		= Encrypt(ID)
				Response.Cookies("SNS_Email")	= Encrypt(Email)
				Response.Cookies("SNS_KName")	= Encrypt(KName)
				Response.Cookies("SNS_Kind")		= Encrypt("K")
				Response.Cookies("SNS_UserID")	= Encrypt(UserID)
				Response.Cookies("SNS_UNUM")		= Encrypt(MemberNum)
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing


				Response.Redirect "/ASP/Member/SnsGate.asp"
				Response.End
		END IF
		oRs.Close


		IF DelFlag = "Y" THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=회원님은 탈퇴하신 정보가 있습니다.<br />재가입하여 주시기 바랍니다.&Script=APP_PopupHistoryBack();"
				Response.End
		END IF





		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing
%>
		<script type="text/javascript">
			var val = "<%=U_NUM%>///<%=ID%>///<%=Email%>///<%=Kname%>///K";
			APP_HistoryBack_SNS_Login(val);
		</script>
<%
ELSE
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=카카오 로그인 오류 입니다.<br />다시 시도해 주십시오.&Script=APP_PopupHistoryBack();"
		Response.End
END IF
%>