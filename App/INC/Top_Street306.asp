	<style type="text/css">
		.gnb .gnb-list3 { width: 100%; height: auto; font-size: 0; white-space: nowrap; overflow-x: auto; border-bottom: 1px solid #e1e1e1; }
		.gnb .gnb-list3 li:nth-of-type(1) { display: inline-block; height: 39px; font-size: 14px; text-align: center; outline: none; width:22%; }
		.gnb .gnb-list3 li:nth-of-type(2) { display: inline-block; height: 39px; font-size: 14px; text-align: center; outline: none; width:22%; }
		.gnb .gnb-list3 li:nth-of-type(3) { display: inline-block; height: 39px; font-size: 14px; text-align: center; outline: none; width:25%; }
		.gnb .gnb-list3 li:nth-of-type(4) { display: inline-block; height: 39px; font-size: 14px; text-align: center; outline: none; width:31%; }
		.gnb .gnb-list3 li>a { position: relative; display: inline-block; padding: 12px 5px 8px; margin: 0 5px; }
		.gnb .gnb-list3 li.current>a:after { width: 100%; background-color: #ff201b; }
		.gnb .gnb-list3 li>a:after { content: ''; display: block; position: absolute; bottom: 0; left: 0; width: 80%; height: 3px; background-color: transparent; transition: width .1s ease-in-out; }
	</style>
</head>

<body>

    <!-- Skip Nav -->
    <a href="#container" id="skipnav" class="skipnav">본문 바로가기</a>

    <!-- Header -->
    <header id="header" class="header special">
        <div class="headline">
            <h1 class="primary-logo street">
                <a href="/ASP/Street306/">Street 306</a>
            </h1>
            <button type="button" class="btn-basket" onclick="APP_GoUrl('/ASP/Order/CartList.asp')">
                <span class="hidden">장바구니</span>
	            <span class="some" id="GNB_CartCount">0</span>
            </button>
		    <button type="button" class="btn-srch" onclick="TopSearch();">
				<span class="hidden">통합검색</span>
			</button>
        </div>
    </header>

    <!-- Navigation -->
    <nav id="gnb" class="gnb">
        <ul class="gnb-list3 black-st">
            <li <%IF PageCode2 = "BP" THEN%>class="current"<%END IF%>><a href="/ASP/Street306/BEST.asp">BEST</a></li>
            <li <%IF PageCode2 = "NP" THEN%>class="current"<%END IF%>><a href="/ASP/Street306/NEW.asp">NEW</a></li>
            <li <%IF PageCode2 = "BR" THEN%>class="current"<%END IF%>><a href="/ASP/Street306/BRAND.asp">BRAND</a></li>
            <li <%IF PageCode2 = "LB" THEN%>class="current"<%END IF%>><a href="/ASP/Street306/LOOKBOOK.asp">LOOKBOOK</a></li>
        </ul>
    </nav>