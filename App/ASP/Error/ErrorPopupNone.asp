<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************/
'ErrorPopup.asp - 에러알림팝업
'Date		: 2018.12.28
'Update		: 
'/****************************************************************************************/

'//페이지 응답헤더 설정------------------------------------------------------
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//-------------------------------------------------------------------------------

'//페이지 코드-----------------------------------------------------------------
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "ER"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'//-------------------------------------------------------------------------------
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include virtual = "/Common/ProgID1.asp" -->

<%
'/****************************************************************************************/
'변수 선언 START
'-----------------------------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

Dim x
DIM i
DIM j
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM Title
DIM Msg
DIM Script
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


Title					 = sqlFilter(Request("Title"))
Msg						 = sqlFilter(Request("Msg"))
Script					 = Request("Script")
%>


<!--#INCLUDE VIRTUAL = "/INC/Header.asp"-->
</head>
<body>

	<main id="container" class="container">
	</main>



    <section class="wrap-pop" id="msgPopup"></section>
    <!-- // Layer 메시지 -->

	
	<div class="dim"></div>

	<section id="alertPop" class="wrap-pop">
		<div class="area-pop">
			<div class="alert">
				<div class="tit-pop">
					<p class="tit" id="alert_title">SHOEMARKER</p>
					<button id="alert_close" onclick="closePop('alertPop')" class="btn-hide-pop">닫기</button>
				</div>
				<div class="container-pop">
					<div class="contents">
						<div class="ly-cont">
							<p id="alert_content" class="t-level4"></p>
						</div>
					</div>
					<div class="btns">
						<button type="button" id="alert_confirm" class="button ty-red">확인</button>
					</div>
				</div>
			</div>
		</div>
	</section>

	<section id="confirmPop" class="wrap-pop">
		<div class="area-pop">
			<div class="alert">
				<div class="tit-pop">
					<p class="tit" id="confirm_title">SHOEMARKER</p>
					<button id="confirm_close" onclick="closePop('confirmPop')" class="btn-hide-pop">닫기</button>
				</div>
				<div class="container-pop">
					<div class="contents">
						<div class="ly-cont">
							<p id="confirm_content" class="t-level4"></p>
						</div>
					</div>
					<div class="btns">
						<button type="button" id="confirm_cancel" class="button ty-black">취소</button>
						<button type="button" id="confirm_confirm" class="button ty-red">확인</button>
					</div>
				</div>
			</div>
		</div>
	</section>

	<section id="messagePop" class="wrap-pop"></section>

	<%IF Title = "" AND Msg = "" THEN%>
	<script type="text/javascript">
		openPop("loading");
		<%=Script%>
	</script>
	<%ELSE%>
	<script type="text/javascript">
			common_msgPopOpen("<%=Title%>", "<%=Msg%>", "<%=Script%>", "msgPopup", "N");
	</script>
	<%END IF%>

</body>

</html>

