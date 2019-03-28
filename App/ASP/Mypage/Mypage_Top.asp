            <section class="ly-mypage-lnb">
                <a href="/ASP/Mypage/" <%IF PageCode2 = "01" THEN%>class="current"<%END IF%>>메인</a>
                <a href="/ASP/Mypage/OrderList.asp" <%IF PageCode2 = "02" THEN%>class="current"<%END IF%>>쇼핑내역</a>
                <a href="/ASP/Mypage/MyShoemarker.asp" <%IF PageCode2 = "03" THEN%>class="current"<%END IF%>>MY슈마커</a>
                <%IF U_MFLAG = "Y" THEN%>
				<a href="#" <%IF PageCode2 = "04" THEN%>class="current"<%END IF%>>쇼핑혜택</a>
				<%END IF%>
                <a href="/ASP/Mypage/MyMemberShip.asp" <%IF PageCode2 = "05" THEN%>class="current"<%END IF%>>회원정보</a>
			</section>
