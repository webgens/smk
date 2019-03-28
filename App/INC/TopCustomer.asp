<%
DIM TopSubMenuTitle
IF IsEmpty(TopSubMenuTitle) OR TopSubMenuTitle = "" THEN
		TopSubMenuTitle = "SHOEMARKER"
END IF
%>
</head>

<body>
    <!-- Skip Nav -->
    <a href="#container" id="skipnav" class="skipnav">본문 바로가기</a>

    <!-- Header -->
    <header id="header" class="header">
        <div class="headline">
            <h1 class="hidden">SHOEMARKER</h1>
            <button type="button" onclick="APP_HistoryBack()" class="btn-goback">
		        <span class="hidden">이전 화면으로 돌아가기</span>
	        </button>
            <span class="tit"><%=TopSubMenuTitle%></span>
            <button type="button" class="btn-basket" onclick="APP_GoUrl('/ASP/Order/CartList.asp')">
				<span class="hidden">장바구니</span>
	            <span class="some" id="GNB_CartCount">0</span>
			</button>
		    <button type="button" class="btn-srch" onclick="TopSearch();">
				<span class="hidden">통합검색</span>
			</button>
        </div>
    </header>

    <nav id="gnb" class="gnb">
        <div class="wrap-gnb">
            <div class="area-gnb">
                <ul class="gnb-list">
					<!--<li class="current"><a href="/">홈</a></li>-->
					<li><a href="javascript:APP_GoUrl('/ASP/Customer/Faq_List.asp')">FAQ</a></li>
					<li><a href="javascript:APP_GoUrl('/ASP/Customer/Store.asp')">전국 매장안내</a></li>
					<li><a href="javascript:APP_GoUrl('/ASP/Customer/PartnerShip.asp')">입점/제휴문의</a></li>
					<li><a href="javascript:APP_GoUrl('/ASP/Customer/GroupPurchase.asp')">단체구매</a></li>
					<li><a href="javascript:APP_GoUrl('/ASP/Customer/Notice_List.asp')">슈마커소식</a></li>
                </ul>
            </div>
        </div>
    </nav>
