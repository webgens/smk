<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'FindID.asp - 아이디찾기 폼 페이지
'Date		: 2018.11.27
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
PageCode2 = "01"
PageCode3 = "02"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
IF U_ID <> "" THEN
		Response.Redirect("/")
		Response.End
END IF


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
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

	
SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
    <script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
    <script type="text/javascript" src="/JS/dev/join.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript">
		function chg_FindIDPwForm(num) {
			$(".part-2 > a").removeClass("current");
			$(".part-2 > a").eq(num).addClass("current");
			if (num == 0) {
				get_FindIDPwdForm("I");
			}
			else {
				get_FindIDPwdForm("P");
			}
		}
		function get_FindIDPwdForm(fType) {
			var url = "";
			if (fType == "I") {
				url = "/ASP/Member/Ajax/FindIDForm.asp";
			}
			else {
				url = "/ASP/Member/Ajax/FindPWForm.asp";
			}

			$.ajax({
				type		 : "post",
				url			 : url,
				async		 : false,
				dataType	 : "text",
				success		 : function (data) {
								$("#FindForm").html(data);
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}
	</script>

<%TopSubMenuTitle = "아이디/비밀번호찾기"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="content">

            <nav class="login-method">
                <ul>
                    <li class="part-2"><a href="javascript:void(0)" onclick="chg_FindIDPwForm(0);" data-ftype="I" class="current">아이디 찾기</a></li>
                    <li class="part-2"><a href="javascript:void(0)" onclick="chg_FindIDPwForm(1);" data-ftype="P">비밀번호 찾기</a></li>
                </ul>
            </nav>

            <div id="FindForm" class="find-form">

            </div>

        </div>
    </main>

	<script type="text/javascript">
		$(function () {
			chg_FindIDPwForm(0);
		});
	</script>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>