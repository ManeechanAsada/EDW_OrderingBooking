using Avantik.Web.Service.Helpers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;

namespace Avantik.Web.Service.Infrastructrue
{
    public static class Authentication
    {
        public static Avantik.Web.Service.Entity.Authentication Authenticate(string token)
        {
            string userId = string.Empty;
            string agencyCode = string.Empty;
            string currencyRcd = string.Empty;
            string languageRcd = string.Empty;
            string allowSeat = string.Empty;
            string allowService = string.Empty;
            string allowPassInfo = string.Empty;
            string allowFlightChange = string.Empty;
            string allowCancelFlight = string.Empty;
            string allowNameChange = string.Empty;
            string allowChangeDetail= string.Empty;
            string strTime = string.Empty;
            string strTimeOut = ConfigHelper.ToString("strTimeOut");
            string strEncrypt = ConfigHelper.ToString("strKey");
            double timePoint = 0;

            Avantik.Web.Service.Entity.Authentication authen = new Entity.Authentication();

            try
            {
                string decryStr = SecurityHelper.DecryptString(token, strEncrypt);

                if (!string.IsNullOrEmpty(decryStr))
                {
                    userId = decryStr.Split('|')[0];
                    agencyCode = decryStr.Split('|')[1];
                    allowSeat = decryStr.Split('|')[2];
                    allowService = decryStr.Split('|')[3];
                    allowPassInfo = decryStr.Split('|')[4];
                    allowFlightChange = decryStr.Split('|')[5];
                    allowCancelFlight = decryStr.Split('|')[6];
                    allowNameChange = decryStr.Split('|')[7];
                    allowChangeDetail = decryStr.Split('|')[8];
                    currencyRcd = decryStr.Split('|')[9];
                    languageRcd = decryStr.Split('|')[10];
                    strTime = decryStr.Split('|')[11];

                    if (!String.IsNullOrEmpty(strTimeOut))
                        timePoint = Int32.Parse(strTimeOut);

                    authen.AgencyCode = agencyCode;
                    authen.UserId = new Guid(userId);
                    authen.B2bAllowSeat = allowSeat;
                    authen.B2bAllowService = allowService;
                    authen.B2bAllowInFoPassenger = allowPassInfo;
                    authen.B2bAllowFlightChange = allowFlightChange;
                    authen.B2bCancelFlight = allowCancelFlight;
                    authen.B2bAllowNameChange = allowNameChange;
                    authen.B2bAllowChangeDetail = allowChangeDetail;
                    authen.CurrencyRcd = currencyRcd;
                    authen.LanguageRcd = languageRcd;

                    DateTime dtValues;
                    dtValues = new DateTime();
                    dtValues = DateTime.ParseExact(strTime, "yyyy-MM-dd HH:mm", null);

                    DateTime dtNow = DateTime.Now;

                    // check time out
                    if (Math.Ceiling(dtNow.Subtract(dtValues).TotalSeconds) > timePoint)
                    {
                       // throw new TimeoutException();
                        authen.ResponseSuccess = false;
                        authen.ResponseMessage = "Security token timeout.";
                        authen.ResponseCode = "A003";
                    }
                    else
                    {
                        string strParams = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}", new string[] 
                        { userId, agencyCode, allowSeat, allowService, allowPassInfo, 
                            allowFlightChange,allowCancelFlight,allowNameChange,allowChangeDetail,currencyRcd,languageRcd,strTime });
                        
                        string hashing = SecurityHelper.EncryptString(strParams, strEncrypt);

                        // valid token
                        if (token == hashing)
                        {
                            authen.ResponseSuccess = true; 
                        }
                        else
                        {
                            authen.ResponseSuccess = false; 
                            authen.ResponseMessage = "Invalid security token.";
                            authen.ResponseCode = "A005";
                        }
                    }
                }
            }
            catch
            {
                authen.ResponseSuccess = false; 
                authen.ResponseMessage = "Invalid security token.";
                authen.ResponseCode = "A005";
            }

            return authen;
        }

    }
}
