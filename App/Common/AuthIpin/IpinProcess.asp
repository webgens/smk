<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'IpinProcess.asp - IPIN 본인인증 인증페이지
'Date		: 2018.11.06
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM sResponseData
DIM sReservedParam1
DIM sReservedParam2
DIM sReservedParam3
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



'사용자 정보 및 CP 요청번호를 암호화한 데이타입니다. (ipin_main.asp 페이지에서 암호화된 데이타와는 다릅니다.)
sResponseData	 = Request("enc_data")
    
'ipin_main.asp 페이지에서 설정한 데이타가 있다면, 아래와 같이 확인가능합니다.
sReservedParam1	 = Request("param_r1")
sReservedParam2	 = Request("param_r2")
sReservedParam3	 = Request("param_r2")


IF sResponseData = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=정보가 존재하지 않습니다. 다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
		Response.End
ELSE
%>
<html>
<head>
	<title>NICE신용평가정보 - IPIN 본인인증</title>
</head>
<body>
	<!-- 본인인증 서비스 팝업을 호출하기 위해서는 다음과 같은 form이 필요합니다. -->
	<form name="form_chk" method="post" action="IpinResult.asp">
		<input type="hidden" name="enc_data" value="<%= sResponseData %>">		<!-- 위에서 업체정보를 암호화 한 데이타입니다. -->
	    
	    <!-- 업체에서 응답받기 원하는 데이타를 설정하기 위해 사용할 수 있으며, 인증결과 응답시 해당 값을 그대로 송신합니다.
	    	 해당 파라미터는 추가하실 수 없습니다. -->
		<input type="hidden" name="param_r1" value="<%=sReservedParam1 %>">
		<input type="hidden" name="param_r2" value="<%=sReservedParam2 %>">
		<input type="hidden" name="param_r3" value="<%=sReservedParam3 %>">
	    
	</form>
   <script type="text/javascript">
   	document.form_chk.submit();
   </script>
</body>
</html>
<%
END IF
%>