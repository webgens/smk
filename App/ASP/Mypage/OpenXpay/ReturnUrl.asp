<%@ Language=VBScript codepage="65001" %>
<%
'*****************************************************************************************'
'ReturnUrl.asp - 배송비 결제 리턴 페이지
'Date		: 2019.01.02
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
<%
'# Set payReqMap = Session.Contents("PAYREQ_MAP")


DIM LGD_OID
LGD_OID = Request.Cookies("PAYREQ_MAP")("LGD_OID")


'payreq_crossplatform.asp 에서 세션에 저장했던 파라미터 값이 유효한지 체크
'세션 유지 시간(로그인 유지시간)을 적당히 유지 하거나 세션을 사용하지 않는 경우 DB처리 하시기 바랍니다.
'# IF IsNull(payReqMap) THEN
IF IsNull(LGD_OID) OR LGD_OID = "" THEN
		'Response.Write "세션이 만료 되었거나 유효하지 않은 요청 입니다."
		Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 세션이 만료 되었거나 유효하지 않은 요청 입니다. 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
		Response.End
END IF
%>

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no">
	<script type="text/javascript">
		function setLGDResult() {
			document.getElementById('LGD_PAYINFO').submit();
		}
<%IF Trim(Request("LGD_RESPCODE")) = "0000" THEN%>
		window.onload = function () {
			setLGDResult();
		}
<%END IF%>
	</script>
</head>
<%IF U_ID = "distance1" OR U_ID = "cybersky" THEN %>
<body>
<%ELSE%>
<body oncontextmenu="return false" onselectstart="return false" ondragstart="return false">
<%END IF %>
	<div style="text-align: center; position: absolute; top: 50%; left: 50%; margin-top: -18px; margin-left: -80px;">
		결제가 진행중입니다.<br />
		잠시만 기다려 주십시오...
	</div>
<%
LGD_RESPCODE	 = Trim(Request("LGD_RESPCODE"))
LGD_RESPMSG		 = Trim(Request("LGD_RESPMSG"))
LGD_PAYKEY		 = ""

IF LGD_RESPCODE = "0000" THEN
		LGD_PAYKEY						= Trim(Request("LGD_PAYKEY"))
		'payReqMap.item("LGD_RESPCODE")	= LGD_RESPCODE
		'payReqMap.item("LGD_RESPMSG")	= LGD_RESPMSG
		'payReqMap.item("LGD_PAYKEY")	= LGD_PAYKEY
	
		Response.Cookies("PAYREQ_MAP")("LGD_RESPCODE")	= LGD_RESPCODE
		Response.Cookies("PAYREQ_MAP")("LGD_RESPMSG")	= LGD_RESPMSG
		Response.Cookies("PAYREQ_MAP")("LGD_PAYKEY")	= LGD_PAYKEY
%>
	<form method="post" name="LGD_PAYINFO" id="LGD_PAYINFO" action="/ASP/Mypage/OpenXpay/PayRes.asp">
<%
		FOR EACH eachitem In Request.Cookies("PAYREQ_MAP")
			Response.Write "		<input type=""hidden"" name="""& eachitem &""" id="""& eachitem &""" value=""" & Request.Cookies("PAYREQ_MAP")(eachitem) & """><br>"
		NEXT
		'FOR EACH eachitem In payReqMap
		'	Response.Write "		<input type=""hidden"" name="""& eachitem &""" id="""& eachitem &""" value=""" & payReqMap.item(eachitem) & """><br>"
		'NEXT
%>
	</form>
<%
ELSEIF LGD_RESPCODE = "S053" THEN
%>
		<script type="text/javascript">
			location.replace('/ASP/Mypage/OrderList.asp');
		</script>
<%
ELSE
		'Response.Write "LGD_RESPCODE:" & LGD_RESPCODE & " ,LGD_RESPMSG:" & LGD_RESPMSG '인증 실패에 대한 처리 로직 추가
		Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다." & "LGD_RESPCODE:" & LGD_RESPCODE & " ,LGD_RESPMSG:" & LGD_RESPMSG, "location.replace('/ASP/Mypage/OrderList.asp');")
		Response.End
END IF
%>
</body>
</html>
