</head>

<body>
    
	<!-- Skip Nav -->
    <a href="#container" id="skipnav" class="skipnav">본문 바로가기</a>

    <!-- Header -->
	<header id="header" class="header">
		<div class="headline">
			<h1 class="hidden">SHOEMARKER</h1>
            <button type="button" class="btn-goback" onclick="APP_HistoryBack()">
				<span class="hidden">이전 화면으로 돌아가기</span>
			</button>
            <span class="tit">상품검색</span>
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
        <section class="wrap-sort">
            <div class="sort" id="OrderType" onclick="OrderByLayerOpen();" style="width:50%;">
                <button type="button" id="orderbytype"><%=SortText%></button>
            </div>
            <div class="search" id="Search" onclick="SmartSearchLayerOpen();" style="width:50%;">
                <button type="button">스마트검색</button>
            </div>
        </section>
    </nav>
