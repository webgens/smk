<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'index.asp - 마이페이지
'Date		: 2018.12.17
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
PageCode1 = "05"
PageCode2 = "05"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

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


DIM mType						'# 회원정보 수정타입
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

mType			 = sqlFilter(request("mType"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
    <script type="text/javascript" src="/JS/dev/mypage.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
<!-- #include virtual="/INC/TopMain.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>


				
                        <div id="OrderMenu" class="ly-title accordion">
                            <div class="selector">
	                            <button type="button" class="btn-list clickEvt" data-target="OrderMenu">나의 멤버십</button>
							</div>
							<div class="option my-recode">
								<ul>
									<li><a href="/ASP/Mypage/MyMemberShip.asp">나의 멤버십</a></li>
									<li><a href="/ASP/Mypage/AddressList.asp">배송지관리</a></li>
									<li><a href="/ASP/Mypage/MyInfoModify.asp">나의 정보 수정</a></li>
									<li><a href="/ASP/Mypage/SnsList.asp">SNS 계정설정</a></li>
								</ul>
							</div>
                        </div>



                <div class="mypage-membership">
                    <section id="contentList_1" class="accord-mypage">
                        <div class="ly-title">
                            <button type="button" class="btn-list clickEvt" data-target="contentList_1">나의 멤버십</button>
                        </div>
                        <div class="ly-content">
							<%IF U_MFLAG = "Y" THEN%>
                            <!-- 나의 멤버십 등급 -->
                            <div class="h-line">
                                <h2 class="h-level4">나의 멤버십 등급</h2>
                            </div>
                            <div class="membership">
                                <p class="grade"><%=U_NAME%> 님의 등급은 <span class="bold">VVIP</span>입니다</p>
                                <p class="accure">최근 1년간 총 누적금액 <span class="bold">382,500</span>원</p>
                                <p class="remain">다음 등급까지 <span class="bold">59,000</span>원 남았습니다.</p>
                            </div>
                            <!-- 받은 등급 혜택 -->
                            <div class="h-line">
                                <h2 class="h-level4">받은 등급 혜택</h2>
                            </div>
                            <div class="grade-benefit">
                                <div class="cnt">
                                    <p class="date">2019.01.01</p>
                                    <p class="tit">등급 상향</p>
                                    <p>VIP > VVIP 등급 상향 혜택 적용</p>
                                    <p><span>2.5</span>%</p>
                                </div>
                                <div class="cnt">
                                    <p class="date">2019.01.01</p>
                                    <p class="tit">장바구니 쿠폰 지급</p>
                                    <p>VVIP 전용 장바구니 쿠폰</p>
                                    <p><span>7.0</span>%</p>
                                </div>
                                <div class="cnt">
                                    <p class="date">2019.01.01</p>
                                    <p class="tit">배송비 할인 쿠폰</p>
                                    <p>구매 1회 배송비 할인 쿠폰</p>
                                    <p><span>1</span>장</p>
                                </div>
                            </div>
							<%ELSE%>
                            <!-- 나의 멤버십 등급 -->
                            <div class="h-line">
                                <h2 class="h-level4">나의 멤버십 등급</h2>
                            </div>
                            <div class="membership">
                                <p class="grade"><%=U_NAME%> 님은 간편로그인 회원입니다.</p>
                                <p class="accure">최근 1년간 총 누적금액 <span class="bold">382,500</span>원</p>
                                <p class="remain">정회원 전환시 각종 쿠폰/혜택을 더 받아보실 수 있습니다.</p>
                            </div>
							<%END IF%>
                            <!-- 등급 상향 기준 안내 -->
                            <div class="h-line">
                                <h2 class="h-level4">등급별 상향 기준안내</h2>
                                <span class="h-date">등급상향 기준은 매달 1일 입니다.</span>
                            </div>
                            <div class="grade-standard">
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit">BRONZE</p>
                                        <p class="ratio">기본 적립율 <span class="bold">1.3%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : 6만원 미만 구매고객</p>
                                        <p>지급 : 회원가입 축하 10% 쿠폰</p>
                                    </div>
                                </div>
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit">SILVER</p>
                                        <p class="ratio">기본 적립율 <span class="bold">1.5%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : 6만원 이상 구매고객</p>
                                        <p>지급 : 장바구니 7% 쿠폰 + 무료 교환쿠폰 1장</p>
                                        <p>매월 : 6만원 이상 구매시 10% 할인쿠폰</p>
                                    </div>
                                </div>
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit">GOLD</p>
                                        <p class="ratio">기본 적립율 <span class="bold">2.0%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : 15만원 이상 구매고객</p>
                                        <p>지급 : 장바구니 10% 쿠폰 + 무료 교환쿠폰 2장</p>
                                        <p>매월 : 6만원 이상 구매시 12% 할인쿠폰</p>
                                    </div>
                                </div>
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit">VIP</p>
                                        <p class="ratio">기본 적립율 <span class="bold">2.3%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : 30만원 이상 구매고객</p>
                                        <p>지급 : 장바구니 12% 쿠폰 + 무료 교환쿠폰 3장</p>
                                        <p>매월 : 8만원 이상 구매시 15% 할인쿠폰</p>
                                    </div>
                                </div>
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit">VVIP</p>
                                        <span class="my-grade">내등급</span>
                                        <p class="ratio">기본 적립율 <span class="bold">2.5%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : 100만원 이상 구매고객</p>
                                        <p>지급 : 장바구니 12% 쿠폰 + 무료 교환쿠폰 3장</p>
                                        <p>매월 : 8만원 이상 구매시 20% 할인쿠폰</p>
                                    </div>
                                </div>
                            </div>
                            <!-- 멤버십 혜택 안내 -->
                            <div class="h-line">
                                <h2 class="h-level4">슈마커 멤버십 혜택 안내</h2>
                                <span class="h-date">모든 회원에게 지급되는 혜택</span>
                            </div>
                            <div class="mbs-benefit">
                                <div class="cnt coupon">
                                    <div class="txt-area">
                                        <p>쿠폰</p>
                                        <%IF U_MFLAG = "Y" THEN%>
										<button type="button" class="button-ty3 ty-bd-black">
											<span>나의 보유 쿠폰</span>
										</button>
										<%END IF%>
                                    </div>
                                    <div class="explain">
                                        <p>첫 구매 감사 쿠폰 <span class="bold">5,000</span>원 지급 (2만원 이상 구매시)</p>
                                        <p>생일 축하 <span class="bold">10,000</span>원 쿠폰 지급 (5만원 이상 구매시)</p>
                                    </div>
                                </div>
                                <div class="cnt Scash">
                                    <div class="txt-area">
                                        <p>S 캐시</p>
                                        <%IF U_MFLAG = "Y" THEN%>
                                        <button type="button" class="button-ty3 ty-bd-black">
											<span>나의 보유 S캐시</span>
										</button>
										<%END IF%>
                                    </div>
                                    <div class="explain">
                                        <div class="cnt1">
                                            <p class="bold">상품 후기 작성</p>
                                            <p>일반후기 작성 : <span class="bold">200</span>원</p>
                                            <p>포토후기 작성 : <span class="bold">500</span>원</p>
                                            <p>(매월1일 후기왕 선정 최대 3만원 쿠폰 증정)</p>
                                        </div>
                                        <div class="cnt2">
                                            <p class="bold">출석체크 참여</p>
                                            <p>10일 : <span class="bold">500</span>원</p>
                                            <p>15일 : <span class="bold">1,000</span>원</p>
                                            <p>20일 : <span class="bold">2,000</span>원</p>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </section>
					<%IF U_MFLAG = "Y" THEN%>
                    <section id="contentList_2" class="accord-mypage">
                        <div class="ly-title">
                            <button type="button" class="btn-list clickEvt" data-target="contentList_2" onclick="common_getPage('getMyAddrList','MyAddrList');">배송지 관리</button>
                        </div>
                        <div class="ly-content" id="getMyAddrList">
                        </div>
                    </section>
                    <section id="contentList_3" class="accord-mypage">
                        <div class="ly-title">
                            <button type="button" class="btn-list clickEvt" data-target="contentList_3" id="contentList_31" onclick="common_getPage('getMyInfoModify','MyInfoModifyPwChk');">나의 정보 수정</button>
                        </div>
                        <div class="ly-content" id="getMyInfoModify">
                        </div>
                    </section>
                    <section id="contentList_4" class="accord-mypage">
                        <div class="ly-title">
                            <button type="button" class="btn-list clickEvt" data-target="contentList_4" onclick="common_getPage('getMySnsList','MySnsList');">SNS 계정설정</button>
                        </div>
                        <div class="ly-content" id="getMySnsList">
                        </div>
                    </section>
					<%END IF%>
                </div>
            </div>
        </div>
    </main>


	<!-- SNS계정 연결 공통 시작 -->
	<!-- SNS계정 로그인 Form -->
 	<form name="SimpleLoginForm" id="SimpleLoginForm" method="post">
		<input type="hidden" name="UID">
		<input type="hidden" name="Email">
		<input type="hidden" name="KName">
		<input type="hidden" name="SNSKind">
	</form>

	<script type="text/javascript">
		// 심플로그인 사용자 삭제
		function SnsLoginDel(uid)
		{
			$.ajax({
				url			 : '/ASP/Mypage/Ajax/MySnsLoginDel.asp',
				data		 : "UID="+uid,
				async	 : false,
				type		 : 'post',
				dataType	 : 'html',
				success	 : function (data, textStatus, jqXHR) {	
					var splitData = data.split("|||||");
					var result = splitData[0];
					var msg = splitData[1];
					if (result == "OK")
					{
						common_msgPopOpen("", "해제 되었습니다.", "common_getPage('getMySnsList','MySnsList');");
						return;
					}
					else if(result == "FAIL"){
						common_msgPopOpen("", msg);
						return;
					}
					else
					{
						common_msgPopOpen("", "오류로 인하여 해제되지 않았습니다.<br>다시 확인 하여 주세요.");
						return;
					}
				},
				error		 : function (data, textStatus, jqXHR) {
								//alert(data.responseText);
								common_msgPopOpen("", "로그인 처리 중 오류.");
				}
			});
		}
	</script>

	<!-- 페이스북 로그인 -->
	<script type="text/javascript">
		var fbLoginFlag = "";
		function fbLogin() {
			fbLoginProcess();
		}

		function fbLoginProcess() {
			FB.login(function (response) {
				if (response.status === 'connected') {
					var fbID	 = "";
					var fbEName	 = "";
					var fbKName	 = "";
					var fbEmail	 = "";

					FB.api('/me', { locale: 'ko_KR' }, function (response) {
						fbKName = response.name;
					});

					FB.api('/me', { fields: 'id, name, email' }, function (response) {
						fbID = response.id;
						fbEName = response.name;
						fbEmail = response.email;

						document.SimpleLoginForm.UID.value = fbID;
						document.SimpleLoginForm.Email.value = response.email;
						document.SimpleLoginForm.KName.value = fbKName;
						document.SimpleLoginForm.SNSKind.value = "F";

						$.ajax({
							url: '/API/Ajax/MemberChk.asp',
							data: $("form[name='SimpleLoginForm']").serialize(),
							async: false,
							type: 'post',
							dataType: 'html',
							success: function (data, textStatus, jqXHR) {
								var splitData = data.split("|||||");
								var result = splitData[0];
								var msg = splitData[1];
								var goUrl = splitData[2];

								if (result == "OK") {
									snsConnection();
								}
								else if (result == "FAIL") {
									alert(msg)
									return;
								}
								else if (result == "FAIL_JOIN") {
									location.href = goUrl;
									return;
								}
								else {
									alert(msg)
									return;
								}
							},
							error: function (data, textStatus, jqXHR) {
								alert(data.responseText);
								alert("로그인 처리 중 오류.");
							}
						});
						/*
						alert(fbID);
						alert(fbKName);
						alert(fbEName);
						alert(response.email);
						*/
					});
				} else if (response.status === undefined) {
				} else if (response.status === unknown) {
				} else {
					alert("페이스북 로그인이 정상적으로 이루어지지 않았습니다.");
					return;
				}
			}, {scope: 'public_profile,email'});
		}

		function statusChangeCallback(response) {
			//console.log(JSON.stringify(response));
			if (response.status === 'connected') {
				fbLoginFlag = "Y";
			} else if (response.status === 'not_authorized') {
			} else {
			}
		}

		window.fbAsyncInit = function() {
			FB.init({
				appId      : '<%=FACEBOOK_LOGIN_CLIENTID%>',
				xfbml      : true,
				version    : 'v3.2'
			});
			FB.AppEvents.logPageView();
		};

		(function(d, s, id){
			var js, fjs = d.getElementsByTagName(s)[0];
			if (d.getElementById(id)) {return;}
			js = d.createElement(s); js.id = id;
			js.src = "//connect.facebook.net/ko_KR/sdk.js";
			fjs.parentNode.insertBefore(js, fjs);
		}(document, 'script', 'facebook-jssdk'));
	</script>
	<!-- 페이스북 로그인 -->
	<!-- 카카오 로그인 -->
	<script src="//developers.kakao.com/sdk/js/kakao.min.js"></script>
	<script type='text/javascript'>
	  //<![CDATA[
	  // 사용할 앱의 JavaScript 키를 설정해 주세요.
		Kakao.init('<%=KAKAO_LOGIN_CLIENTID%>');
			
		function loginWithKakao() {
			// 로그인 창을 띄웁니다.
			Kakao.Auth.login({
				success: function (authObj) {
					Kakao.API.request({
						url: '/v2/user/me',
						success: function (res) {
							var KakaoEmail = res.kakao_account.email;
							if (KakaoEmail == undefined) {
								KakaoEmail = "";
							}
							document.SimpleLoginForm.UID.value = res.id;
							document.SimpleLoginForm.Email.value = KakaoEmail;
							document.SimpleLoginForm.KName.value = res.properties.nickname;
							document.SimpleLoginForm.SNSKind.value = "K";

							$.ajax({
								url: '/API/Ajax/MemberChk.asp',
								data: $("form[name='SimpleLoginForm']").serialize(),
								async: false,
								type: 'post',
								dataType: 'html',
								success: function (data, textStatus, jqXHR) {
									var splitData = data.split("|||||");
									var result = splitData[0];
									var msg = splitData[1];
									var goUrl = splitData[2];

									if (result == "OK") {
										snsConnection();
									}
									else if (result == "FAIL") {
										alert(msg)
										return;
									}
									else if (result == "FAIL_JOIN") {
										location.href = goUrl;
										return;
									}
									else {
										alert(msg)
										return;
									}
								},
								error: function (data, textStatus, jqXHR) {
									alert(data.responseText);
									alert("로그인 처리 중 오류.");
								}
							});
							/*
							alert(JSON.stringify(res));
							alert(res.id);
							alert(res.properties.nickname);
							alert(res.kaccount_email);
							alert(Kakao.Auth.getAccessToken());
							*/
						},
						fail: function (err) {
							alert(JSON.stringify(err));
						}
					});
				},
				fail: function (err) {
					alert(JSON.stringify(err));
				}
			});
		}
	  //]]>

	</script>
	<!-- 카카오 로그인 -->
	<!-- SNS계정 연결 공통 끝 -->
	
<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%IF mType = "MemModify" THEN%>
<script>
	$("#contentList_3 .ly-title button").click();
</script>
<%END IF%>

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
