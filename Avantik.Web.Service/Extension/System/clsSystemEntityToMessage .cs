using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Message.Agency;
using Avantik.Web.Service.Message.Fee;
using Avantik.Web.Service.Message.FormOfPayment;
using Avantik.Web.Service.Message.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Extension.System
{
    public static class SystemEntityToMessage 
    {
        public static IList<Message.Country> ToObjMessage(this IList<Entity.Country> objEntityList)
        {
            List<Message.Country> objMessageList = null;
            if (objEntityList != null)
            {
                objMessageList = new List<Message.Country>();
                for (int i = 0; i < objEntityList.Count; i++)
                {
                    objMessageList.Add(objEntityList[i].ToObjMessage());
                }
            }
            return objMessageList;
        }

        public static Message.Country ToObjMessage(this Entity.Country obj)
        {
            Message.Country objResponse = null;

            if(obj != null)
            {
                objResponse = new Message.Country();
                objResponse.AddressCountryFlag = obj.AddressCountryFlag;
                objResponse.AgencyCountryFlag = obj.AgencyCountryFlag;
                objResponse.CountryCodeLong = obj.CountryCodeLong;
                objResponse.CountryNumber = obj.CountryNumber;
                objResponse.CountryRcd = obj.CountryRcd;
                objResponse.CurrencyRcd = obj.CurrencyRcd;
                objResponse.DisplayName = obj.DisplayName;
                objResponse.IssueCountryFlag = obj.IssueCountryFlag;
                objResponse.NationalityCountryFlag = obj.NationalityCountryFlag;
                objResponse.PhonePrefix = obj.PhonePrefix;
                objResponse.ResidenceCountryFlag = obj.ResidenceCountryFlag;
                objResponse.StatusCode = obj.StatusCode;
                objResponse.VatDisplay = obj.VatDisplay;
            }

            return objResponse;
        }

        public static IList<Message.Language> ToObjMessage(this IList<Entity.Language> objEntityList)
        {
            List<Message.Language> objMessageList = null;
            if (objEntityList != null)
            {
                objMessageList = new List<Message.Language>();
                for (int i = 0; i < objEntityList.Count; i++)
                {
                    objMessageList.Add(objEntityList[i].ToObjMessage());
                }
            }
            return objMessageList;
        }

        public static Message.Language ToObjMessage(this Entity.Language obj)
        {
            Message.Language objResponse = null;

            if (obj != null)
            {
                objResponse = new Message.Language();
                objResponse.CharacterSet = obj.CharacterSet;
                objResponse.DisplayName = obj.DisplayName;
                objResponse.LanguageRcd = obj.LanguageRcd;
            }

            return objResponse;
        }

        public static IList<Message.Title> ToObjMessage(this IList<Entity.Title> objEntityList)
        {
            List<Message.Title> objMessageList = null;
            if (objEntityList != null)
            {
                objMessageList = new List<Message.Title>();
                for (int i = 0; i < objEntityList.Count; i++)
                {
                    objMessageList.Add(objEntityList[i].ToObjMessage());
                }
            }
            return objMessageList;
        }

        public static Message.Title ToObjMessage(this Entity.Title obj)
        {
            Message.Title objResponse = null;

            if (obj != null)
            {
                objResponse = new Message.Title();
                objResponse.DisplayName = obj.DisplayName;
                objResponse.GenderTypeRcd = obj.GenderTypeRcd;
                objResponse.TitleRcd = obj.TitleRcd;
            }

            return objResponse;
        }

        public static IList<Message.Currency> ToObjMessage(this IList<Entity.Currency> objEntityList)
        {
            List<Message.Currency> objMessageList = null;
            if (objEntityList != null)
            {
                objMessageList = new List<Message.Currency>();
                for (int i = 0; i < objEntityList.Count; i++)
                {
                    objMessageList.Add(objEntityList[i].ToObjMessage());
                }
            }
            return objMessageList;
        }

        public static Message.Currency ToObjMessage(this Entity.Currency obj)
        {
            Message.Currency objResponse = null;

            if (obj != null)
            {
                objResponse = new Message.Currency();
                objResponse.CurrencyRcd = obj.CurrencyRcd;
                objResponse.DisplayName = obj.DisplayName;
                objResponse.MaxVoucherValue = obj.MaxVoucherValue;
                objResponse.NumberOfDecimals = obj.NumberOfDecimals;
                objResponse.RoundingRule = obj.RoundingRule;

            }

            return objResponse;
        }

        public static IList<Message.SpecialService> ToObjMessage(this IList<Entity.SpecialService> objEntityList)
        {
            List<Message.SpecialService> objMessageList = null;
            if (objEntityList != null)
            {
                objMessageList = new List<Message.SpecialService>();
                for (int i = 0; i < objEntityList.Count; i++)
                {
                    objMessageList.Add(objEntityList[i].ToObjMessage());
                }
            }
            return objMessageList;
        }

        public static Message.SpecialService ToObjMessage(this Entity.SpecialService obj)
        {
            Message.SpecialService objResponse = null;

            if (obj != null)
            {
                objResponse = new Message.SpecialService();
                objResponse.BookingSegmentId = obj.BookingSegmentId;
                objResponse.CreateBy = obj.CreateBy;
                objResponse.CreateDateTime = obj.CreateDateTime;
                objResponse.CutOffTime = obj.CutOffTime;
                objResponse.DestinationRcd = obj.DestinationRcd;
                objResponse.DisplayName = obj.DisplayName;
                objResponse.FeeId = obj.FeeId;
                objResponse.FlightId = obj.FlightId;
                objResponse.HelpText = obj.HelpText;
                objResponse.IncludeActionCodeFlag = obj.IncludeActionCodeFlag;
                objResponse.IncludeCateringFlag = obj.IncludeCateringFlag;
                objResponse.IncludeFlightSegmentFlag = obj.IncludeFlightSegmentFlag;
                objResponse.IncludeNumberOfServiceFlag = obj.IncludeNumberOfServiceFlag;
                objResponse.IncludePassengerAssistanceFlag = obj.IncludePassengerAssistanceFlag;
                objResponse.IncludePassengerNameFlag = obj.IncludePassengerNameFlag;
                objResponse.InventoryControlFlag = obj.InventoryControlFlag;
                objResponse.ManifestFlag = obj.ManifestFlag;
                objResponse.NewRecord = obj.NewRecord;
                objResponse.NumberOfUnits = obj.NumberOfUnits;
                objResponse.OriginRcd = obj.OriginRcd;
                objResponse.PassengerId = obj.PassengerId;
                objResponse.PassengerSegmentServiceId = obj.PassengerSegmentServiceId;
                objResponse.SendInterlineReplyFlag = obj.SendInterlineReplyFlag;
                objResponse.ServiceOnRequestFlag = obj.ServiceOnRequestFlag;
                objResponse.ServiceSupportedFlag = obj.ServiceSupportedFlag;
                objResponse.ServiceText = obj.ServiceText;
                objResponse.SpecialServiceChangeStatusRcd = obj.SpecialServiceChangeStatusRcd;
                objResponse.SpecialServiceGroupRcd = obj.SpecialServiceGroupRcd;
                objResponse.SpecialServiceRcd = obj.SpecialServiceRcd;
                objResponse.SpecialServiceStatusRcd = obj.SpecialServiceStatusRcd;
                objResponse.StatusCode = obj.StatusCode;
                objResponse.TextAllowedFlag = obj.TextAllowedFlag;
                objResponse.TextRequiredFlag = obj.TextRequiredFlag;
                objResponse.UpdateBy = obj.UpdateBy;
                objResponse.UpdateDateTime = obj.UpdateDateTime;

            }

            return objResponse;
        }

        public static Message.FormOfPayment.FormOfPayment ToObjMessage(this Entity.FormOfPayment.FormOfPayment objEntityList)
        {
            Message.FormOfPayment.FormOfPayment objResponse = new Message.FormOfPayment.FormOfPayment();

            if (objEntityList != null && objEntityList.FormOfPaymentSubType != null)
            {
                List<FormOfPaymentSubType> FormOfPaymentSubTypes = null;

                if (objEntityList != null)
                {
                    FormOfPaymentSubTypes = new List<FormOfPaymentSubType>();
                    for (int i = 0; i < objEntityList.FormOfPaymentSubType.Count; i++)
                    {
                        FormOfPaymentSubTypes.Add(objEntityList.FormOfPaymentSubType[i].ToObjMessage());
                    }
                }

                objResponse.FormOfPaymentSubType = FormOfPaymentSubTypes;
            }

            return objResponse;
        }


        public static FormOfPaymentSubType ToObjMessage(this Entity.FormOfPayment.FormOfPaymentSubType obj)
        {
            FormOfPaymentSubType objResponse = null;

            if (obj != null)
            {
                objResponse = new FormOfPaymentSubType();
                objResponse.FormOfPaymentSubtypeRcd = obj.FormOfPaymentSubtypeRcd;
                objResponse.DisplayName = obj.DisplayName;
                objResponse.FormOfPaymentRcd = obj.FormOfPaymentRcd;
                objResponse.ExpiryDays = obj.ExpiryDays;
                objResponse.CardCode = obj.CardCode;
                objResponse.VoucherReference = obj.VoucherReference;
                objResponse.CvvRequiredFlag = obj.CvvRequiredFlag;
                objResponse.ValidateDocumentNumberFlag = obj.ValidateDocumentNumberFlag;
                objResponse.DisplayCvvFlag = obj.DisplayCvvFlag;
                objResponse.MultiplePaymentFlag = obj.MultiplePaymentFlag;
                objResponse.ApprovalCodeRequiredFlag = obj.ApprovalCodeRequiredFlag;
                objResponse.DisplayApprovalCodeFlag = obj.DisplayApprovalCodeFlag;
                objResponse.DisplayExpiryDateFlag = obj.DisplayExpiryDateFlag;
                objResponse.ExpiryDateRequiredFlag = obj.ExpiryDateRequiredFlag;
                objResponse.TravelAgencyPaymentFlag = obj.TravelAgencyPaymentFlag;
                objResponse.AgencyPaymentFlag = obj.AgencyPaymentFlag;
                objResponse.InternetPaymentFlag = obj.InternetPaymentFlag;
                objResponse.RefundPaymentFlag = obj.RefundPaymentFlag;
                objResponse.AddressRequiredFlag = obj.AddressRequiredFlag;
                objResponse.DisplayAddressFlag = obj.DisplayAddressFlag;
                objResponse.ShowPosIndictorFlag = obj.ShowPosIndictorFlag;
                objResponse.RequirePosIndicatorFlag = obj.RequirePosIndicatorFlag;
                objResponse.DisplayIssueDateFlag = obj.DisplayIssueDateFlag;
                objResponse.DisplayIssueNumberFlag = obj.DisplayIssueNumberFlag;
            }

            return objResponse;
        }

        public static IEnumerable<DocumentView> MappingDocumentTypeView(this IEnumerable<Document> documents)
        {
            try
            {
                if (documents != null)
                {
                    IList<DocumentView> rts = new List<DocumentView>();
                    DocumentView obv = null;
                    foreach (Document a in documents)
                    {
                        obv = new DocumentView();


                        obv.document_type_rcd = a.DocumentTypeRcd;
                        obv.display_name = a.DisplayName;

                        rts.Add(obv);

                    }

                    return rts;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                throw;
            }
        }
    }
}
