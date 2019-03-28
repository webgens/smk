<%@ Language=VBScript codepage="65001" %>
<%option Explicit%>
<%
'*****************************************************************************************'
'OrderMailSend.asp - 주문완료시 메일 보내기
'Date		: 2019.01.15
'Update	: 
'*****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------

%>

<!-- #include virtual="/ADO/ADODBCommon_NOHttps.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn
DIM oRs
DIM oCmd

DIM CompName
DIM MEmail
DIM MailTitle

DIM OrderCode
DIM ProductName
DIM OrderCnt
DIM OrderPrice
DIM SalePrice
DIM UseCouponPrice
DIM UsePointPrice
DIM UseScashPrice
DIM DeliveryPrice
DIM SettlePrice
DIM PayType
DIM OrderName
DIM OrderTel
DIM OrderHp
DIM OrderEmail
DIM DelvType
DIM ShopNM
DIM ReceiveName
DIM ReceiveTel
DIM ReceiveHp
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2
DIM ReceiveAddress
DIM ReceiptFlag
DIM Memo
DIM OrderDate
DIM OrderTime
DIM LGD_AMOUNT
DIM LGD_FINANCENAME
DIM LGD_CARDINSTALLMONTH
DIM LGD_ACCOUNTNUM
DIM LGD_TELNO

DIM SendFlag	: SendFlag	= "Y"
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


OrderCode			= Trim(request("LGD_OID"))
IF OrderCode = "" THEN
		OrderCode			= Trim(request("OrderCode"))
END IF



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'-----------------------------------------------------------------------------------------------------------'
'주문 정보 검색 START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Select_By_OrderCode"

		.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,		adParamInput, 20,		OrderCode)
		.Parameters.Append .CreateParameter("@UserID",			adVarChar,		adParamInput, 20,		U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		OrderName					 = oRs("OrderName")
		OrderHp						 = oRs("OrderHp")
		OrderEmail					 = oRs("OrderEmail")
ELSE
		SendFlag	= "N"
END IF
oRs.Close


IF SendFlag = "Y" THEN
		'-----------------------------------------------------------------------------------------------------------'
		'주문 정보 검색 START
		'-----------------------------------------------------------------------------------------------------------'
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_EShop_Order_Select_For_OrderInfo"

				.Parameters.Append .CreateParameter("@OrderCode",	adVarChar, adParaminput,	20,		OrderCode)
				.Parameters.Append .CreateParameter("@UserID",		adVarChar, adParamInput,	20,		U_NUM)
				.Parameters.Append .CreateParameter("@OrderName",	adVarChar, adParamInput,	50,		OrderName)
				.Parameters.Append .CreateParameter("@OrderHp",		adVarChar, adParamInput,	20,		OrderHp)
				.Parameters.Append .CreateParameter("@OrderEmail",	adVarChar, adParamInput,	50,		OrderEmail)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing

		IF NOT oRs.EOF THEN

				ProductName					 = oRs("ProductName")
				OrderCnt					 = oRs("OrderCnt")
				OrderPrice					 = oRs("OrderPrice")
				SalePrice					 = oRs("SalePrice")
				UseCouponPrice				 = oRs("UseCouponPrice")	
				UsePointPrice				 = oRs("UsePointPrice")	
				UseScashPrice				 = oRs("UseScashPrice")	
				DeliveryPrice				 = oRs("DeliveryPrice")	
				PayType						 = oRs("PayType")
				OrderName					 = oRs("OrderName")
				OrderTel					 = oRs("OrderTel")
				OrderHp						 = oRs("OrderHp")
				OrderEmail					 = oRs("OrderEmail")
				DelvType					 = oRs("DelvType")
				ShopNM						 = oRs("ShopNM")
				ReceiveName					 = oRs("ReceiveName")
				ReceiveTel					 = oRs("ReceiveTel")
				ReceiveHp					 = oRs("ReceiveHp")
				ReceiveZipCode				 = oRs("ReceiveZipCode")
				ReceiveAddr1				 = oRs("ReceiveAddr1")
				ReceiveAddr2				 = oRs("ReceiveAddr2")
				ReceiptFlag					 = oRs("ReceiptFlag")
				Memo						 = oRs("Memo")
				OrderDate					 = oRs("OrderDate")
				OrderTime					 = oRs("OrderTime")
				LGD_AMOUNT					 = oRs("LGD_AMOUNT")
				LGD_FINANCENAME				 = oRs("LGD_FINANCENAME")
				LGD_CARDINSTALLMONTH		 = oRs("LGD_CARDINSTALLMONTH")
				LGD_ACCOUNTNUM				 = oRs("LGD_ACCOUNTNUM")
				LGD_TELNO					 = oRs("LGD_TELNO")

				IF CInt(OrderCnt) > 1 THEN
						Productname = ProductName & " 외 " & FormatNumber(CInt(OrderCnt) - 1, 0) & "개"
				END IF

				IF DelvType = "S" THEN
						ReceiveAddress	= ShopNM
				ELSE
						ReceiveAddress	= ReceiveAddr1 & " " & ReceiveAddr2
				END IF

				IF IsNull(Memo) = false AND Memo <> "" THEN
						Memo	= "(" & Memo & ")"
				END IF

				IF PayType = "C" THEN
						IF LGD_CARDINSTALLMONTH = "00" THEN
							LGD_CARDINSTALLMONTH	= "(일시불)"
						ELSE
							LGD_CARDINSTALLMONTH	= "(" & FormatNumber(LGD_CARDINSTALLMONTH,0) & "개월 할부)"
						END IF
				END IF

				IF OrderEMail = "" THEN
						SendFlag	= "N"
				END IF
		ELSE
				SendFlag	= "N"
		End IF
		oRs.Close
		'-----------------------------------------------------------------------------------------------------------'
		'주문정보 검색 END
		'-----------------------------------------------------------------------------------------------------------'
END IF


IF SendFlag = "Y" Then

		'-----------------------------------------------------------------------------------------------------------'
		'쇼핑몰 정보 검색 Start
		'-----------------------------------------------------------------------------------------------------------'
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.Commandtype = adCmdStoredProc
				.CommandText = "USP_Admin_EShop_BizInfo_Select"
		End WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing

		IF NOT oRs.EOF THEN
			CompName	 = oRs("CompName")
			MEmail		 = oRs("Email")
		End IF
		oRs.Close
		'-----------------------------------------------------------------------------------------------------------'
		'쇼핑몰 정보 검색 End
		'-----------------------------------------------------------------------------------------------------------'
			


		'-----------------------------------------------------------------------------------------------------------'
		'주문정보 발송 START
		'-----------------------------------------------------------------------------------------------------------'
		MailContents = ""
        MailContents = MailContents & "<div class=""container"" style=""margin-bottom: 25px;"">" & VbCrLf 
        MailContents = MailContents & "    <div class=""wrap-order-confirm"" style=""position: relative;padding: 0 25px;text-align: left;"">" & VbCrLf 
        MailContents = MailContents & "        <div class=""area-tit"" style=""width: 100%;padding: 30px 0;text-align: center;"">" & VbCrLf 
        MailContents = MailContents & "            <h2 class=""tit"" style=""padding-top: 45px;margin-bottom: 20px;font-size: 28px;color: #282828;text-align: center;background: url('" & FRONT_URL & "/Images/ico/ico-order-num-tit.png')no-repeat center top;line-height: 36px;-webkit-background-size: 36px auto;background-size: 36px auto;"">주문이<br>완료되었습니다.</h2>" & VbCrLf 
        MailContents = MailContents & "            <span class=""order-num"" style=""width: 100%;margin-bottom: 14px;font-size: 18px;text-align: center;line-height: 1;display: inline-block;outline:none;"">주문번호 <em style=""font-weight: 700;color: #ff201b;font-style: normal;"">" & OrderCode & "</em></span>" & VbCrLf 
        MailContents = MailContents & "            <div class=""inform"" style=""position: relative;padding: 4px 10px;margin: 0 12px;border: 1px solid #191919;border-radius: 10px;font-size: 14px;line-height: 20px;-ms-word-break: keep-all;word-break: keep-all;text-align: center;"">" & VbCrLf 
        MailContents = MailContents & "                주문 및 배송확인은 ‘마이페이지 > 주문내역’ 에서 언제라도 하실 수 있습니다." & VbCrLf 
        MailContents = MailContents & "            </div>" & VbCrLf 
        MailContents = MailContents & "        </div>" & VbCrLf 
        MailContents = MailContents & "        <div class=""wrap-content"" style=""width: 100%;margin-bottom: 40px;"">" & VbCrLf 
        MailContents = MailContents & "            <div style=""width: 100%;height: 7px;background: url('" & FRONT_URL & "/Images/img/bg-pop.png')no-repeat;outline:none;""></div>" & VbCrLf 
        MailContents = MailContents & "            <div class=""area-content"" style=""position: relative;padding: 36px 20px;background-color: #fff;"">" & VbCrLf 
        MailContents = MailContents & "                <div class=""ly-order-confirm"">" & VbCrLf 
        MailContents = MailContents & "                    <h3 class=""tit"" style=""margin-bottom: 20px;font-size: 20px;color: #191919;font-weight: 800;text-align: center;"">주문내용 확인</h3>" & VbCrLf 
        MailContents = MailContents & "                    <div class=""order-list"">" & VbCrLf 
        MailContents = MailContents & "                        <table style=""width: 100%;table-layout: fixed;padding: 20px 0;margin-bottom: 25px;border-bottom: 1px solid #000;list-style: none;"">" & VbCrLf 
        MailContents = MailContents & "                            <tr>" & VbCrLf 
        MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">주문상품</th>" & VbCrLf 
        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & ProductName & "</td>" & VbCrLf 
        MailContents = MailContents & "                            </tr>" & VbCrLf 
        MailContents = MailContents & "                            <tr>" & VbCrLf 
        MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">배송지</th>" & VbCrLf 
        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & ReceiveAddress & " " & Memo & "</td>" & VbCrLf 
        MailContents = MailContents & "                            </tr>" & VbCrLf 
		IF PayType = "C" THEN
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">결제수단</th>" & VbCrLf 
				MailContents = MailContents & "                                <td style=""padding: 10px 0;"">신용카드</td>" & VbCrLf 
				MailContents = MailContents & "                            </tr>" & VbCrLf 
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">결제카드</th>" & VbCrLf 
		        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & LGD_FINANCENAME & "카드 " & LGD_CARDINSTALLMONTH & "</td>" & VbCrLf 
		        MailContents = MailContents & "                            </tr>" & VbCrLf 
		ELSEIF PayType = "B" THEN
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">결제수단</th>" & VbCrLf 
				MailContents = MailContents & "                                <td style=""padding: 10px 0;"">계좌이체</td>" & VbCrLf 
				MailContents = MailContents & "                            </tr>" & VbCrLf 
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">이체은행</th>" & VbCrLf 
		        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & LGD_FINANCENAME & "은행" & "</td>" & VbCrLf 
		        MailContents = MailContents & "                            </tr>" & VbCrLf 
		ELSEIF PayType = "V" THEN
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">결제수단</th>" & VbCrLf 
				MailContents = MailContents & "                                <td style=""padding: 10px 0;"">가상계좌(무통장)</td>" & VbCrLf 
				MailContents = MailContents & "                            </tr>" & VbCrLf 
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">입금은행</th>" & VbCrLf 
		        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & LGD_FINANCENAME & "은행" & "</td>" & VbCrLf 
		        MailContents = MailContents & "                            </tr>" & VbCrLf 
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">입금계좌</th>" & VbCrLf 
		        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & LGD_ACCOUNTNUM & "</td>" & VbCrLf 
		        MailContents = MailContents & "                            </tr>" & VbCrLf 
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">예금주</th>" & VbCrLf 
		        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & MALL_LGD_ACCOUNTOWNER & "</td>" & VbCrLf 
		        MailContents = MailContents & "                            </tr>" & VbCrLf 
		ELSEIF PayType = "M" THEN
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">결제수단</th>" & VbCrLf 
				MailContents = MailContents & "                                <td style=""padding: 10px 0;"">휴대폰결제</td>" & VbCrLf 
				MailContents = MailContents & "                            </tr>" & VbCrLf 
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">통신사</th>" & VbCrLf 
		        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & LGD_FINANCENAME & "</td>" & VbCrLf 
		        MailContents = MailContents & "                            </tr>" & VbCrLf 
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">휴대폰</th>" & VbCrLf 
		        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & LGD_TELNO & "</td>" & VbCrLf 
		        MailContents = MailContents & "                            </tr>" & VbCrLf 
		ELSEIF PayType = "N" THEN
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">결제수단</th>" & VbCrLf 
				MailContents = MailContents & "                                <td style=""padding: 10px 0;"">네이버페이</td>" & VbCrLf 
				MailContents = MailContents & "                            </tr>" & VbCrLf 
		ELSEIF PayType = "S" THEN
				MailContents = MailContents & "                            <tr>" & VbCrLf 
				MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">결제수단</th>" & VbCrLf 
				MailContents = MailContents & "                                <td style=""padding: 10px 0;"">슈마커페이</td>" & VbCrLf 
				MailContents = MailContents & "                            </tr>" & VbCrLf 
		END IF
        MailContents = MailContents & "                            <tr>" & VbCrLf 
        MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">결제금액</th>" & VbCrLf 
        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & FormatNumber(LGD_AMOUNT,0) & "원</td>" & VbCrLf 
        MailContents = MailContents & "                            </tr>" & VbCrLf 
        MailContents = MailContents & "                            <tr>" & VbCrLf 
        MailContents = MailContents & "                                <th style=""width: 100px;padding: 10px 0;font-size: 14px;color: #282828;font-weight: 700;"">결제일시</th>" & VbCrLf 
        MailContents = MailContents & "                                <td style=""padding: 10px 0;"">" & GetDateYMD(OrderDate) & "</td>" & VbCrLf 
        MailContents = MailContents & "                            </tr>" & VbCrLf 
        MailContents = MailContents & "                        </table>" & VbCrLf 
        MailContents = MailContents & "                        <div class=""ly-inform"">" & VbCrLf 
        MailContents = MailContents & "                            <p style=""margin: 0;font-size: 14px;color: #282828;line-height: 24px;font-weight: 500;-ms-word-break: keep-all;word-break: keep-all;"">상품은 주문/입금확인 후 1~2일 내에 출고가 진행됩니다.</p>" & VbCrLf 
        MailContents = MailContents & "                            <p style=""margin: 0;font-size: 14px;color: #282828;line-height: 24px;font-weight: 500;-ms-word-break: keep-all;word-break: keep-all;"">(단, 주말/공휴일 제외)</p>" & VbCrLf 
        MailContents = MailContents & "                            <p style=""margin: 0;font-size: 14px;color: #282828;line-height: 24px;font-weight: 500;-ms-word-break: keep-all;word-break: keep-all;"">상품 출고 후 배송번호를 통해 진행사항을 확인하실 수 있습니다</p>" & VbCrLf 
        MailContents = MailContents & "                        </div>" & VbCrLf 
        MailContents = MailContents & "                    </div>" & VbCrLf 
        MailContents = MailContents & "                </div>" & VbCrLf 
        MailContents = MailContents & "            </div>" & VbCrLf 
        MailContents = MailContents & "            <div style=""width: 100%;height: 15px;margin-top: -7px;background: url('" & FRONT_URL & "/Images/img/bg-pop-2.png')no-repeat;outline:none;""></div>" & VbCrLf 
        MailContents = MailContents & "        </div>" & VbCrLf 
        MailContents = MailContents & "    </div>" & VbCrLf 
        MailContents = MailContents & "</div>" & VbCrLf 
		'-----------------------------------------------------------------------------------------------------------'
		'주문정보 발송 END
		'-----------------------------------------------------------------------------------------------------------'
%>

		<!-- #include virtual = "/Common/Mail/MailForm.asp" -->
<%
		MailTitle = "고객님의 주문이 정상적으로 처리되었습니다."

		Call MailSend (CompName, MEmail, OrderName, OrderEmail, MailTitle, mail_con)
	

END IF


Set oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
 

