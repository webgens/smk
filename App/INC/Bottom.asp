	<!-- Script -->
    <script src="/JS/app.js?ver=<%=U_DATE%><%=U_TIME%>"></script>

    <!-- PopUp -->
	<section class="wrap-pop" id="DimDepth1"></section>
	<section class="wrap-pop" id="DimDepth2"></section>
    <!-- // PopUp -->
    <!-- Layer 메시지 -->
    <section class="wrap-pop" id="msgPopup"></section>
    <!-- // Layer 메시지 -->

	<!-- 우편번호 검색 레이어 -->
    <section class="wrap-pop" id="PopupPostSearch">
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">우편번호 검색</p>
                    <button type="button" onclick="closePop('PopupPostSearch')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents" id="PopupPostContents">
                    </div>
                    <div class="btns">
                        <button type="button" onclick="closePop('PopupPostSearch')" class="button ty-red">닫기</button>
                    </div>
                </div>
            </div>
        </div>
    </section>
	<!-- 우편번호 검색 레이어 -->

	
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

    <section class="wrap-pop" id="TopSearch">
        <div class="area-dim"></div>

        <div class="area-pop" id="TopSearchView">
            <div class="full search" style="z-index:206;">
                <div class="input-search">
                    <span class="enter">
						<input type="text" name="SearchText" id="SearchText" placeholder="검색어를 입력해 주세요.">
					</span>
                    <button type="button" class="btn-search" onclick="TopSearchGo();">검색</button>
					<button type="button" class="btn-hide" onclick="TopSearchClose();">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="wrap-search">
                            <div id="tabs" class="tabaa" data-use="">
                                <ul class="tab-selector">
                                    <li style="width:50%;" id="ts1" class="active"><a href="javascript:TopSearchWordView('P');" data-target="tabs-col1">인기검색어</a></li>
                                    <li style="width:50%;" id="ts2"><a href="javascript:TopSearchWordView('R');" data-target="tabs-col2">추천검색어</a></li>
                                </ul>
                                <div id="tabs-col1" class="tab-panel" style="display:block;">
                                    <div class="related-search">
                                        <ul id="WordView" style="text-align:center; font-size:16px;">
                                        </ul>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>


            </div>

		</div>
	</section>


	<form name="botLoginForm" id="botLoginForm" method="post" action="/ASP/Member/Login.asp">
		<input type="hidden" name="ProgID" value="<%=ProgID%>" />
	</form>

	<form name="FooterNoticeViewForm" id="FooterNoticeViewForm" method="post">
		<input type="hidden" name="Idx" value="" />
	</form>

	<form name="TopSearchForm" id="TopSearchForm" method="get" action="/ASP/Product/SearchProductList.asp">
		<input type="hidden" name="SearchWord" id="SearchWord" />
	</form>

    <section class="wrap-pop" id="Category1PopView">
        <div class="area-pop">
            <div class="top-exposed" id="CategoryList"  style="z-index:105;">
                <!-- 더블 클래스 vertical로 팝업 호출/ 닫기 -->
                <div class="tit-pop">
                    <p class="tit">바로가기</p>
                    <button class="btn-hide-pop" onclick="GetCategory1Close();">닫기</button>
                </div>

                <div class="container-pop" id="Category1Cont">

                </div>
            </div>
        </div>
    </section>

	<section id="BrandErrPop" class="wrap-pop">
        <div class="area-dim" style="z-index:106;"></div>

        <div class="area-pop">
            <div class="alert" style="z-index:108;">
                <div class="tit-pop">
                    <p class="tit">SHOEMARKER</p>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="ly-cont">
                            <p class="t-level4" id="msg"></p>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button"  onclick="BrandErrPopclose();" class="button ty-red">확인</button>
                    </div>
                </div>
            </div>
        </div>
	</section>

    <section class="wrap-pop" id="ProductLatestView">
        <div class="area-pop">
            <div class="top-exposed" id="ProductLatest" style="z-index:105;">
                <div class="tit-pop">
                    <p class="tit">최근 본 상품</p>
                    <button class="btn-hide-pop" type="button" onclick="close_ProductLatest();">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents bg-ty1">
                        <div class="inner-cont full">
                            <div class="wrap-picked-list">
                                <ul class="picked-list" id="ProductLatestCont">
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>


    <section class="wrap-pop" id="TiemSalePop">
        <div class="area-dim"></div>

        <div class="area-pop" id="TimeSaleContent">
        </div>
    </section>

	<!-- 사이즈 레이어 //-->
    <section class="wrap-pop" id="SizePop">
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit">사이즈</p>
                    <button type="button" class="btn-hide-pop" onclick="$('#SizePop').hide();">닫기</button>
                </div>
                <div class="container-pop">
                    <div class="contents">
                        <div class="pop-size-check">
                            <ul class="size" id="sizelist">
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>

	<script type="text/javascript">
		$(function () {
			$(".loading_dim").hide();
			$("#loading").hide();
		});

		footerNoticeRolling();

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
		wcs_add["wa"] = "s_142651b7553d";
		if (!_nasa) var _nasa = {};
		wcs.inflow();
		wcs_do(_nasa);
	</script>

	<!-- AceCounter Mobile WebSite Gathering Script V.7.5.20170208 -->
	<script type="text/javascript">
		var _AceGID = (function () { var Inf = ['app.shoemarker.co.kr', 'app.shoemarker.co.kr', 'AZ1A74686', 'AM', '0', 'NaPm,Ncisy', 'ALL', '0']; var _CI = (!_AceGID) ? [] : _AceGID.val; var _N = 0; if (_CI.join('.').indexOf(Inf[3]) < 0) { _CI.push(Inf); _N = _CI.length; } return { o: _N, val: _CI }; })();
		var _AceCounter = (function () { var G = _AceGID; var _sc = document.createElement('script'); var _sm = document.getElementsByTagName('script')[0]; if (G.o != 0) { var _A = G.val[G.o - 1]; var _G = (_A[0]).substr(0, _A[0].indexOf('.')); var _C = (_A[7] != '0') ? (_A[2]) : _A[3]; var _U = (_A[5]).replace(/\,/g, '_'); _sc.src = (location.protocol.indexOf('http') == 0 ? location.protocol : 'http:') + '//cr.acecounter.com/Mobile/AceCounter_' + _C + '.js?gc=' + _A[2] + '&py=' + _A[1] + '&up=' + _U + '&rd=' + (new Date().getTime()); _sm.parentNode.insertBefore(_sc, _sm); return _sc.src; } })();
	</script>
	<noscript><img src='http://gmb.acecounter.com/mwg/?mid=AZ1A74686&tp=noscript&ce=0&' border='0' width='0' height='0' alt=''></noscript>
	<!-- AceCounter Mobile Gathering Script End -->

	<!-- adinsight 공통스크립트 start -->
	<script type="text/javascript">
		var TRS_AIDX = 11295;
		var TRS_PROTOCOL = document.location.protocol;
		document.writeln();
		var TRS_URL = TRS_PROTOCOL + '//' + ((TRS_PROTOCOL == 'https:') ? 'analysis.adinsight.co.kr' : 'adlog.adinsight.co.kr') + '/emnet/trs_esc.js';
		document.writeln("<scr" + "ipt language='javascript' src='" + TRS_URL + "'></scr" + "ipt>");
	</script>
	<!-- adinsight 공통스크립트 end -->

	<!-- 트래킹 스크립트 영역 -->


</body>

</html>