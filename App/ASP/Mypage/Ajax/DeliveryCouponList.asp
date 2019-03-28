<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'DeliveryCouponList.asp - 교환/반품시 무료배송쿠폰 리스트 선택 폼 페이지
'Date		: 2019.01.02
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


Response.Write "OK|||||"
%>					
        <div class="area-pop" id="DeliveryCouponList">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">보유중인 쿠폰</p>
                    <button type="button" onclick="closePop('DimDepth2')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="coupon-lists">
<%
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Coupon_Member_Select_For_Delivery"

		.Parameters.Append .CreateParameter("@MemberNum",		adInteger,	adParamInput,   ,	U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		Do Until oRs.EOF
%>
                            <div class="coupon-list">
                                <div class="tit">
                                    <div class="inn">
                                        <div class="off">무료배송</div>
                                        <div class="name"><%=oRs("CouponName")%></div>
                                    </div>
                                </div>
                                <div class="time-limit">
                                    <em>사용기한</em>
									<%=Replace(GetDateYMD(Left(oRs("StartDT"),8)), "-", ". ")%> ~<br>
									<%IF oRs("EndDT") = "999999999999" THEN%>
									제한기간없음
									<%ELSE%>
									<%=Replace(GetDateYMD(Left(oRs("EndDT"),8)), "-", ". ")%>
									<%END IF%>
                                </div>
                                <div class="radiogroup sel">
                                    <div class="inner">
                                        <span class="radio">
											<input type="radio" name="DeliveryCouponIdx" id="DeliveryCouponIdx_<%=oRs("MemberCouponIdx")%>" value="<%=oRs("MemberCouponIdx")%>" />
										</span>
                                        <label for="DeliveryCouponIdx_<%=oRs("MemberCouponIdx")%>">선택</label>
                                    </div>
                                </div>
                            </div>
<%
				oRs.MoveNext
		Loop
ELSE
%>
                            <div class="coupon-list">
                                <div class="area-empty">
									<span class="icon-empty"></span>
									<p class="tit-empty">보유하신 무료배송 쿠폰이 없습니다.</p>
                                </div>
                            </div>
<%
END IF
oRs.Close
%>
                        </div>
                    </div>

                    <div class="btns">
                        <button type="button" onclick="orderChangeReturn()" class="button ty-red">적용</button>
                    </div>
                </div>
            </div>
        </div>

		<script type="text/javascript">
			$(function () {
				// radio 버튼 클릭 액션
				$('#DeliveryCouponList .radiogroup input').on('click', function () {
					var $this = $(this);

					if ($this.prop('checked') === true) {
						$this.closest('.radio').addClass('is-checked');	//.siblings().removeClass('checked');
					} else {
						$this.closest('.radio').removeClass('is-checked');
					}
				});
			})
		</script>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>