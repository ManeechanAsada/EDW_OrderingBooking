using Avantik.Web.Service.Model.Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.COMHelper;
using System.Runtime.InteropServices;
using Avantik.Web.Service.Model.COM.Extension;
using Avantik.Web.Service.Exception.Booking;
using Avantik.Web.Service.Entity.FormOfPayment;
using Avantik.Web.Service.Entity.Booking;

namespace Avantik.Web.Service.Model.COM
{
    public class SystemService : RunComplus, ISystemModelService
    {
        string _server = string.Empty;
        public SystemService(string server, string user, string pass, string domain)
            :base(user,pass,domain)
        {
            _server = server;
        }

        public IList<Country> GetCountry(string language)
        {
            tikAeroProcess.Inventory objInventory = null;
            ADODB.Recordset rs = null;
            List<Country> countries = new List<Country>();

            try
            {
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Inventory", _server);
                    objInventory = (tikAeroProcess.Inventory)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objInventory = new tikAeroProcess.Inventory(); }

                //set default language to EN
                if (language.Length == 0)
                { language = "EN"; }

                rs = objInventory.GetCountry(ref language);
                //convert Recordset to Object
                if (rs != null && rs.RecordCount > 0)
                {
                    Country c;
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        c = new Country();

                        c.CountryRcd = RecordsetHelper.ToString(rs, "country_rcd");
                        c.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        c.StatusCode = RecordsetHelper.ToString(rs, "status_code");
                        c.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        c.VatDisplay = RecordsetHelper.ToString(rs, "vat_display");
                        c.CountryCodeLong = RecordsetHelper.ToString(rs, "country_code_long");
                        c.CountryNumber = RecordsetHelper.ToString(rs, "country_number");
                        c.PhonePrefix = RecordsetHelper.ToString(rs, "phone_prefix");
                        c.IssueCountryFlag = RecordsetHelper.ToByte(rs, "issue_country_flag");
                        c.ResidenceCountryFlag = RecordsetHelper.ToByte(rs, "residence_country_flag");
                        c.NationalityCountryFlag = RecordsetHelper.ToByte(rs, "nationality_country_flag");
                        c.AddressCountryFlag = RecordsetHelper.ToByte(rs, "address_country_flag");
                        c.AgencyCountryFlag = RecordsetHelper.ToByte(rs, "agency_country_flag");

                        countries.Add(c);
                        c = null;
                        rs.MoveNext();
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (objInventory != null)
                {
                    Marshal.FinalReleaseComObject(objInventory);
                    objInventory = null;
                }
                RecordsetHelper.ClearRecordset(ref rs);
            }

            return countries;

        }

        public IList<Language> GetLanguage(string language)
        {
            tikAeroProcess.Reservation objReservation = null;
            ADODB.Recordset rs = null;
            List<Language> languages = new List<Language>();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Reservation", _server);
                    objReservation = (tikAeroProcess.Reservation)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objReservation = new tikAeroProcess.Reservation(); }
                //set default language to EN
                if (language.Length == 0)
                { language = "EN"; }
                rs = objReservation.GetLanguages(ref language);

                //convert Recordset to Object
                if (rs != null && rs.RecordCount > 0)
                {
                    Language l;
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        l = new Language();

                        l.LanguageRcd = RecordsetHelper.ToString(rs, "language_rcd");
                        l.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        l.CharacterSet = RecordsetHelper.ToString(rs, "character_set");

                        languages.Add(l);
                        l = null;
                        rs.MoveNext();
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (objReservation != null)
                { Marshal.FinalReleaseComObject(objReservation); }
                objReservation = null;

                RecordsetHelper.ClearRecordset(ref rs);
            }

            return languages;

        }

        public IList<Title> GetTitle(string language)
        {
            tikAeroProcess.Reservation objReservation = null;
            ADODB.Recordset rs = null;
            List<Title> titles = new List<Title>();

            try
            {
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Reservation", _server);
                    objReservation = (tikAeroProcess.Reservation)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objReservation = new tikAeroProcess.Reservation(); }

                if (string.IsNullOrEmpty(language) == true)
                { language = "EN"; }
                rs = objReservation.GetPassengerTitles(ref language);

                //convert Recordset to Object
                if (rs != null && rs.RecordCount > 0)
                {
                    Title t;
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        t = new Title();

                        t.TitleRcd = RecordsetHelper.ToString(rs, "title_rcd");
                        t.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        t.GenderTypeRcd = RecordsetHelper.ToString(rs, "gender_type_rcd");

                        titles.Add(t);
                        t = null;
                        rs.MoveNext();
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (objReservation != null)
                {
                    Marshal.FinalReleaseComObject(objReservation);
                    objReservation = null;
                }

                RecordsetHelper.ClearRecordset(ref rs);
            }

            return titles;

        }

        public IList<Currency> GetCurrency(string language)
        {
            tikAeroProcess.Payment objPayment = null;
            ADODB.Recordset rs = null;
            List<Currency> currencies = new List<Currency>();
            try
            {

                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objPayment = new tikAeroProcess.Payment(); }

                rs = objPayment.GetCurrencies(language);

                //convert Recordset to Object
                if (rs != null && rs.RecordCount > 0)
                {
                    Currency c;
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        c = new Currency();
                        c.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        c.CurrencyNumber = RecordsetHelper.ToString(rs, "currency_number");
                        c.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        c.MaxVoucherValue = RecordsetHelper.ToDecimal(rs, "max_voucher_value");
                        c.RoundingRule = RecordsetHelper.ToDecimal(rs, "rounding_rule");
                        c.NumberOfDecimals = RecordsetHelper.ToInt16(rs, "number_of_decimals");
                        currencies.Add(c);
                        c = null;
                        rs.MoveNext();
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }
                RecordsetHelper.ClearRecordset(ref rs);
            }
            return currencies;
        }

        public IList<SpecialService> GetSpecialService(string language)
        {
            tikAeroProcess.Reservation objReservation = null;
            ADODB.Recordset rs = null;
            try
            {
                List<SpecialService> specialServices = new List<SpecialService>();
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Reservation", _server);
                    objReservation = (tikAeroProcess.Reservation)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objReservation = new tikAeroProcess.Reservation(); }

                rs = objReservation.GetSpecialServices(language);
                //convert Recordset to Object
                if (rs != null && rs.RecordCount > 0)
                {
                    SpecialService s;
                    rs.Filter = 0;
                    while (!rs.EOF)
                    {
                        s = new SpecialService();

                        s.SpecialServiceRcd = RecordsetHelper.ToString(rs, "special_service_rcd");
                        s.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        s.TextAllowedFlag = RecordsetHelper.ToByte(rs, "text_allowed_flag");
                        s.ManifestFlag = RecordsetHelper.ToByte(rs, "manifest_flag");
                        s.TextRequiredFlag = RecordsetHelper.ToByte(rs, "text_required_flag");
                        s.ServiceOnRequestFlag = RecordsetHelper.ToByte(rs, "service_on_request_flag");
                        s.IncludePassengerNameFlag = RecordsetHelper.ToByte(rs, "include_passenger_name_flag");
                        s.IncludeFlightSegmentFlag = RecordsetHelper.ToByte(rs, "include_flight_segment_flag");
                        s.IncludeActionCodeFlag = RecordsetHelper.ToByte(rs, "include_action_code_flag");
                        s.IncludeNumberOfServiceFlag = RecordsetHelper.ToByte(rs, "include_number_of_service_flag");
                        s.IncludeCateringFlag = RecordsetHelper.ToByte(rs, "include_catering_flag");
                        s.IncludePassengerAssistanceFlag = RecordsetHelper.ToByte(rs, "include_passenger_assistance_flag");
                        s.ServiceSupportedFlag = RecordsetHelper.ToByte(rs, "service_supported_flag");
                        s.SendInterlineReplyFlag = RecordsetHelper.ToByte(rs, "send_interline_reply_flag");
                        s.CutOffTime = RecordsetHelper.ToInt32(rs, "cut_off_time");
                        s.StatusCode = RecordsetHelper.ToString(rs, "status_code");
                        s.PassengerSegmentServiceId = RecordsetHelper.ToGuid(rs, "passenger_segment_service_id");
                        s.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                        s.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                        s.ServiceText = RecordsetHelper.ToString(rs, "service_text");
                        s.SpecialServiceStatusRcd = RecordsetHelper.ToString(rs, "special_service_status_rcd");

                        specialServices.Add(s);
                        s = null;
                        rs.MoveNext();
                    }
                }

                return specialServices;
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (objReservation != null)
                {
                    Marshal.FinalReleaseComObject(objReservation);
                    objReservation = null;
                }
                RecordsetHelper.ClearRecordset(ref rs);
            }

        }

        public FormOfPayment GetFormOfPaymentSubType(string type, string language)
        {
            tikAeroProcess.Payment objPayment = null;
            ADODB.Recordset rs = null;
            FormOfPayment formOfPayments = new FormOfPayment();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objPayment = new tikAeroProcess.Payment(); }

                //set default language to EN
                if (string.IsNullOrEmpty(language))
                { language = "EN"; }
                rs = objPayment.GetFormOfPaymentSubTypes(ref type, ref language);

                if (rs != null && rs.RecordCount > 0)
                {
                    FormOfPaymentSubType formOfPaymentSubType;
                    formOfPayments.FormOfPaymentSubType = new List<FormOfPaymentSubType>();

                    rs.Filter = 0;
                    while (!rs.EOF)
                    {
                        formOfPaymentSubType = new FormOfPaymentSubType();

                        formOfPaymentSubType.VoucherReference = RecordsetHelper.ToInt32(rs, "voucher_reference");

                        formOfPaymentSubType.ExpiryDays = RecordsetHelper.ToInt16(rs, "expiry_days");

                        formOfPaymentSubType.CvvRequiredFlag = RecordsetHelper.ToByte(rs, "cvv_required_flag");
                        formOfPaymentSubType.ValidateDocumentNumberFlag = RecordsetHelper.ToByte(rs, "validate_document_number_flag");
                        formOfPaymentSubType.DisplayCvvFlag = RecordsetHelper.ToByte(rs, "display_cvv_flag");
                        formOfPaymentSubType.MultiplePaymentFlag = RecordsetHelper.ToByte(rs, "multiple_payment_flag");
                        formOfPaymentSubType.ApprovalCodeRequiredFlag = RecordsetHelper.ToByte(rs, "approval_code_required_flag");
                        formOfPaymentSubType.DisplayApprovalCodeFlag = RecordsetHelper.ToByte(rs, "display_approval_code_flag");
                        formOfPaymentSubType.DisplayExpiryDateFlag = RecordsetHelper.ToByte(rs, "display_expiry_date_flag");
                        formOfPaymentSubType.ExpiryDateRequiredFlag = RecordsetHelper.ToByte(rs, "expiry_date_required_flag");
                        formOfPaymentSubType.TravelAgencyPaymentFlag = RecordsetHelper.ToByte(rs, "travel_agency_payment_flag");
                        formOfPaymentSubType.AgencyPaymentFlag = RecordsetHelper.ToByte(rs, "agency_payment_flag");
                        formOfPaymentSubType.InternetPaymentFlag = RecordsetHelper.ToByte(rs, "internet_payment_flag");
                        formOfPaymentSubType.RefundPaymentFlag = RecordsetHelper.ToByte(rs, "refund_payment_flag");
                        formOfPaymentSubType.AddressRequiredFlag = RecordsetHelper.ToByte(rs, "address_required_flag");
                        formOfPaymentSubType.DisplayAddressFlag = RecordsetHelper.ToByte(rs, "display_address_flag");
                        formOfPaymentSubType.ShowPosIndictorFlag = RecordsetHelper.ToByte(rs, "show_pos_indictor_flag");
                        formOfPaymentSubType.RequirePosIndicatorFlag = RecordsetHelper.ToByte(rs, "require_pos_indicator_flag");
                        formOfPaymentSubType.DisplayIssueDateFlag = RecordsetHelper.ToByte(rs, "display_issue_date_flag");
                        formOfPaymentSubType.DisplayIssueNumberFlag = RecordsetHelper.ToByte(rs, "display_issue_number_flag");

                        formOfPaymentSubType.FormOfPaymentSubtypeRcd = RecordsetHelper.ToString(rs, "form_of_payment_subtype_rcd");
                        formOfPaymentSubType.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        formOfPaymentSubType.FormOfPaymentRcd = RecordsetHelper.ToString(rs, "form_of_payment_rcd");
                        formOfPaymentSubType.CardCode = RecordsetHelper.ToString(rs, "card_code");

                        formOfPayments.FormOfPaymentSubType.Add(formOfPaymentSubType);
                        formOfPaymentSubType = null;
                        rs.MoveNext();
                    }
                }

            }
            catch
            { }
            finally
            {
                if (objPayment != null)
                { Marshal.FinalReleaseComObject(objPayment); }
                objPayment = null;

                RecordsetHelper.ClearRecordset(ref rs);
            }

            return formOfPayments;
        }

        public List<Document> GetDocumentType(string language)
        {
            tikAeroProcess.Inventory objInventory = null;
            ADODB.Recordset rs = null;
            try
            {
                List<Document> docs = new List<Document>();
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Inventory", _server);
                    objInventory = (tikAeroProcess.Inventory)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objInventory = new tikAeroProcess.Inventory(); }

                if (string.IsNullOrEmpty(language) == true)
                { language = "EN"; }
                rs = objInventory.GetDocumentType(ref language);

                //convert Recordset to Object
                if (rs != null && rs.RecordCount > 0)
                {
                    docs.FillDocumentType(ref rs);
                }
                return docs;
            }
            catch (SystemException ex)
            {
                throw ex;
            }
            finally
            {
                if (objInventory != null)
                {
                    Marshal.FinalReleaseComObject(objInventory);
                    objInventory = null;
                }

                RecordsetHelper.ClearRecordset(ref rs);
            }

            
        }
    }
}
