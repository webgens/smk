<%
DIM mail_con : mail_con = ""
DIM MailContents

mail_con = ""
mail_con = mail_con & "<!DOCTYPE html>" & VbCrLf 
mail_con = mail_con & "<html lang=""ko"" style=""width: 100%;height: 100%;font-family:  '돋움' ,'Dotum',  sans-serif"">" & VbCrLf 
mail_con = mail_con & "<head>" & VbCrLf 
mail_con = mail_con & "    <meta charset=""UTF-8"">" & VbCrLf 
mail_con = mail_con & "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no, target-densitydpi=medium-dpi"" />" & VbCrLf 
mail_con = mail_con & "    <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">" & VbCrLf 
mail_con = mail_con & "    <title>SHOEMARKER</title>" & VbCrLf 
mail_con = mail_con & "</head>" & VbCrLf 
mail_con = mail_con & "<body style=""width: 100%;height: 100%;padding: 0;margin: 0;font-size: 16px;color: #282828;letter-spacing: -0.5px;"">" & VbCrLf 
mail_con = mail_con & "    <div class=""wrap-ly-all"" style=""background-color: #f3f3f3;box-sizing: border-box;padding-bottom: 20px;"">" & VbCrLf 
mail_con = mail_con & "        <div class=""inner-all"" style=""max-width: 680px;margin: 0 auto;min-height: 100%;"">" & VbCrLf 
mail_con = mail_con & "            <div class=""wrap-header"" style=""height: 113px;border-top: 6px solid #e41d16;"">" & VbCrLf 
mail_con = mail_con & "                <img src=""" & FRONT_URL & "/Images/ico/logo-emd-shoemarker.png"" style=""display: block;width: 107px;height: 107px;margin: 0 auto;"" alt="""">" & VbCrLf 
mail_con = mail_con & "            </div>" & VbCrLf 


mail_con = mail_con & MailContents				  


mail_con = mail_con & "            <table style=""width: 100%;padding: 15px 0;margin-bottom: 15px;background-color: #4e4e4e;"">" & VbCrLf 
mail_con = mail_con & "                <tbody>" & VbCrLf 
mail_con = mail_con & "                    <tr>" & VbCrLf 
mail_con = mail_con & "                        <td style=""font-size: 12px;color: #fff;text-align: center;""><a href=""" & FRONT_URL & """ target=""_blank"" style=""font-size: inherit;color: inherit;text-decoration: none;"">온라인 쇼핑하기</a></td>" & VbCrLf 
mail_con = mail_con & "                        <td style=""font-size: 12px;color: #fff;text-align: center;""><a href=""" & FRONT_URL & "/ASP/Customer/Store.asp"" target=""_blank"" style=""font-size: inherit;color: inherit;text-decoration: none;"">슈마커 매장찾기</a></td>" & VbCrLf 
mail_con = mail_con & "                        <td style=""font-size: 12px;color: #fff;text-align: center;""><a href=""" & FRONT_URL & """ target=""_blank"" style=""font-size: inherit;color: inherit;text-decoration: none;"">슈마커 앱설치</a></td>" & VbCrLf 
mail_con = mail_con & "                    </tr>" & VbCrLf 
mail_con = mail_con & "                </tbody>" & VbCrLf 
mail_con = mail_con & "            </table>" & VbCrLf 
mail_con = mail_con & "            <table style=""width: 100%;margin: 0 25px 25px;"">" & VbCrLf 
mail_con = mail_con & "                <thead>" & VbCrLf 
mail_con = mail_con & "                    <tr>" & VbCrLf 
mail_con = mail_con & "                        <th colspan=""3"" style=""font-size: 14px;color: #282828;padding: 0 0 10px;font-weight: 600;text-align: left;-ms-word-break: keep-all;word-break: keep-all;"">본 메일은 발신전용입니다. 문의사항은 슈마커고객센터로 주시기 바랍니다.</th>" & VbCrLf 
mail_con = mail_con & "                    </tr>" & VbCrLf 
mail_con = mail_con & "                </thead>" & VbCrLf 
mail_con = mail_con & "                <tbody>" & VbCrLf 
mail_con = mail_con & "                    <tr>" & VbCrLf 
mail_con = mail_con & "                        <td style=""width: 135px;font-size: 12px;color: #767676;border-right: 1px solid #767676;line-height: 1;"">(주)에스엠케이티앤아이</td>" & VbCrLf 
mail_con = mail_con & "                        <td style=""width: 97px;font-size: 12px;color: #767676;border-right: 1px solid #767676;padding: 0 10px;line-height: 1;"">대표이사 : 안영환</td>" & VbCrLf 
mail_con = mail_con & "                        <td style=""font-size: 12px;color: #767676;padding: 0 10px;line-height: 1;"">고객센터 : 080-030-2809</td>" & VbCrLf 
mail_con = mail_con & "                    </tr>" & VbCrLf 
mail_con = mail_con & "                    <tr>" & VbCrLf 
mail_con = mail_con & "                        <td colspan=""3"" style=""font-size: 12px;color: #767676;padding:  6px 0 3px;"">주소 : 서울특별시 강남구 테헤란로 306 (역삼동 706-1,카이트타워) 7층</td>" & VbCrLf 
mail_con = mail_con & "                    </tr>" & VbCrLf 
mail_con = mail_con & "                    <tr>" & VbCrLf 
mail_con = mail_con & "                        <td colspan=""3"" style=""font-size: 12px;color: #767676;padding: 3px 0;"">&#9426; SHOEMARKER All RIGHTS RESERVED</td>" & VbCrLf 
mail_con = mail_con & "                    </tr>" & VbCrLf 
mail_con = mail_con & "                </tbody>" & VbCrLf 
mail_con = mail_con & "            </table>" & VbCrLf 
mail_con = mail_con & "        </div>" & VbCrLf 
mail_con = mail_con & "    </div>" & VbCrLf 
mail_con = mail_con & "</body>" & VbCrLf 
mail_con = mail_con & "</html>" & VbCrLf 
%> 
