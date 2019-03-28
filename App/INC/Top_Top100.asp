</head>

<body>

    <!-- Header -->
    <header id="header" class="header">
        <div class="headline">
            <h1 class="primary-logo">
                <a href="#">SHOEMARKER</a>
            </h1>
            <button type="button" class="btn-basket" onclick="APP_GoUrl('/ASP/Order/CartList.asp')">
				<span class="hidden">장바구니</span>
				<span class="some" id="GNB_CartCount">0</span>
			</button>
			<button type="button" class="btn-srch">
			    <span class="hidden">통합검색</span>
			</button>
        </div>
    </header>

    <!-- Navigation -->
    <nav id="gnb" class="gnb">
        <ul class="gnb-list">
            <li><a href="/">홈</a></li>
            <li><a href="/">BRANDS</a></li>
            <li class="current"><a href="/ASP/Product/Top100.asp">TOP100</a></li>
            <li><a href="/ASP/Product/Sale.asp">SALE</a></li>
            <li><a href="/ASP/Product/Today.asp">TODAY'S DEAL</a></li>
            <li><a href="#">이벤트</a></li>
            <li><a href="/ASP/Street306/">STREET306</a></li>
            <li><a href="/ASP/ShoemarkerOnly/">ONLY</a></li>
        </ul>
    </nav>
