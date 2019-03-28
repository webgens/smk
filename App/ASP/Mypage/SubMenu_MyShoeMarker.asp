						<%IF U_MFLAG = "Y" THEN%>
						<ul>
							<li><a href="/ASP/Mypage/MyPickList.asp">MY&hearts;</a></li>
							<li><a href="/ASP/Mypage/MyReentry.asp">재입고 알림</a></li>
							<li><a href="/ASP/Mypage/MyReview.asp">상품후기</a></li>
							<li><a href="/ASP/Mypage/Qna.asp">상품문의</a></li>
						</ul>
						<%ELSE%>
						<ul>
							<li><a href="/ASP/Mypage/MyPickList.asp">MY&hearts;</a></li>
							<li><a href="/ASP/Mypage/MyReview.asp">상품후기</a></li>
							<li><a href="/ASP/Mypage/Qna.asp">상품문의</a></li>
							<li></li>
						</ul>
						<%END IF%>