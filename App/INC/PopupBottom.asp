	<!-- Script -->
    <script src="/JS/app.js?ver=<%=U_DATE%><%=U_TIME%>"></script>



	
	<div class="dim"></div>

	<section id="alertPop" class="wrap-pop">
		<div class="area-dim" style="z-index:101"></div>
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
		<div class="area-dim" style="z-index:101"></div>
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


	<script type="text/javascript">
		$(function () {
			$(".loading_dim").hide();
			$("#loading").hide();
		});
	</script>


	<!-- 트래킹 스크립트 영역 -->

	<!-- WIDERPLANET  SCRIPT START 2019.1.8 -->
	<div id="wp_tg_cts" style="display:none;"></div>
	<script type="text/javascript">
		var wptg_tagscript_vars = wptg_tagscript_vars || [];
		wptg_tagscript_vars.push(
		(function () {
			return {
				wp_hcuid: "<%=U_Num%>",  	/*고객넘버 등 Unique ID (ex. 로그인  ID, 고객넘버 등 )를 암호화하여 대입.
					 *주의 : 로그인 하지 않은 사용자는 어떠한 값도 대입하지 않습니다.*/
				ti: "24585",	/*광고주 코드 */
				ty: "Home",	/*트래킹태그 타입 */
				device: "mobile"	/*디바이스 종류  (web 또는  mobile)*/
			};
		}));
	</script>
	<script type="text/javascript" async src="//cdn-aitg.widerplanet.com/js/wp_astg_4.0.js"></script>
	<!-- // WIDERPLANET  SCRIPT END 2019.1.8 -->

	<!-- 공통 적용 스크립트 , 모든 페이지에 노출되도록 설치. 단 전환페이지 설정값보다 항상 하단에 위치해야함 --> 
	<script type="text/javascript" src="//wcs.naver.net/wcslog.js"> </script> 
	<script type="text/javascript">
		if (!wcs_add) var wcs_add = {};
		wcs_add["wa"] = "s_15f3a5309d6f";
		if (!_nasa) var _nasa = {};
		wcs.inflow();
		wcs_do(_nasa);
	</script>

	<!-- Google Tag Manager -->
	<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
	new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
	j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
	'https://www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
	})(window,document,'script','dataLayer','GTM-TKWJ4DX');</script>
	<!-- End Google Tag Manager -->

	<!-- 트래킹 스크립트 영역 -->


</body>

</html>