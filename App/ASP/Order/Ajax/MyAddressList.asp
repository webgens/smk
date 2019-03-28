<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'MyAddressList.asp - 배송지 리스트 폼 페이지
'Date		: 2018.12.28
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
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y
DIM z

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderSheetIdx
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

OrderSheetIdx			= sqlFilter(Request("OrderSheetIdx"))



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



Response.Write "OK|||||"
%>					
        <div class="area-pop" id="MyAddress">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">배송지 목록</p>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="wrap-shipping-list">
<%
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_MyAddress_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum",		adInteger,	adParamInput,   ,	U_NUM)
END WITH
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

i = 0			
IF NOT oRs.EOF THEN
%>
                            <ul>
<%
		Do Until oRs.EOF
%>
                                <li>
                                    <span class="shipping-list">
										<input type="hidden" name="AddressName_<%=oRs("Idx")%>" value="<%=oRs("AddressName")%>" />
										<input type="hidden" name="ReceiveName_<%=oRs("Idx")%>" value="<%=oRs("ReceiveName")%>" />
										<input type="hidden" name="ReceiveTel_<%=oRs("Idx")%>" value="<%=oRs("ReceiveTel")%>" />
										<input type="hidden" name="ReceiveHP_<%=oRs("Idx")%>" value="<%=oRs("ReceiveHP")%>" />
										<input type="hidden" name="ReceiveZipCode_<%=oRs("Idx")%>" value="<%=oRs("ReceiveZipCode")%>" />
										<input type="hidden" name="ReceiveAddr1_<%=oRs("Idx")%>" value="<%=oRs("ReceiveAddr1")%>" />
										<input type="hidden" name="ReceiveAddr2_<%=oRs("Idx")%>" value="<%=oRs("ReceiveAddr2")%>" />
										<input type="hidden" name="MainFlag_<%=oRs("Idx")%>" value="<%=oRs("MainFlag")%>" />
										<input type="radio" name="MyAddress" id="MyAddress_<%=oRs("Idx")%>" value="<%=oRs("Idx")%>" />
										<label for="MyAddress_<%=oRs("Idx")%>">
											<span class="line">
												<span class="name"><%=oRs("ReceiveName")%></span>
												<span class="addr">(<%=oRs("ReceiveZipCode")%>) <%=oRs("ReceiveAddr1")%> <%=oRs("ReceiveAddr2")%></span>
											</span>
										</label>
                                    </span>
                                </li>
<%
				oRs.MoveNext
				i = i + 1
		Loop
%>
                            </ul>
<%
ELSE
%>
							<div class="area-empty">
								<span class="icon-empty"></span>
								<p class="tit-empty">등록된 배송지가 없습니다</p>
							</div>
<%
END IF
oRs.Close
%>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="setMyAddress('list', '<%=OrderSheetIdx%>', '')" class="button ty-red">확인</button>
                    </div>
                </div>
            </div>
        </div>

		<form name="MyAddressInfo" method="post">
			<input type="hidden" name="OrderSheetIdx"	value="" />
			<input type="hidden" name="AddressName"		value="" />
			<input type="hidden" name="ReceiveName"		value="" />
			<input type="hidden" name="ReceiveTel1"		value="" />
			<input type="hidden" name="ReceiveTel23"	value="" />
			<input type="hidden" name="ReceiveHP1"		value="" />
			<input type="hidden" name="ReceiveHP23"		value="" />
			<input type="hidden" name="ReceiveZipCode"	value="" />
			<input type="hidden" name="ReceiveAddr1"	value="" />
			<input type="hidden" name="ReceiveAddr2"	value="" />
		</form>

		<script type="text/javascript">
			$(function () {
				// radio 버튼 클릭 액션
				/*
				$('.rd-type3 input').on('click', function () {
					var $this = $(this);

					if ($this.prop('checked') === true) {
						$this.closest('.list-type3 li').addClass('checked').siblings().removeClass('checked');
					}

				});
				*/
			})
		</script>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>