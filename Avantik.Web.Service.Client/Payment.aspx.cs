using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.Booking;
using Avantik.Web.Service.Message.SeatMap;
using Avantik.Web.Service.Proxy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Avantik.Web.Service.Client
{
    public partial class Payment : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (HttpContext.Current.Request.QueryString.Count > 0)
            {
                switch(HttpContext.Current.Request.QueryString["type"])
                {
                    case "pcc" : 
                        PaymentCreditCard();
                        break;
                    case "pvc" :
                        PaymentVoucher();
                        break;
                    case "rvc":
                        ReadVoucher();
                        break;
                    case "pca":
                        PaymentCreditAgency();
                        break;
                    case "gsm" :
                        GetSeatMap();
                        break;
                    default :
                        break;
                }
            }
        }

        private void GetSeatMap()
        {
            FlightServiceProxy objFlight = new FlightServiceProxy();
            GetSeatMapRequest req = new GetSeatMapRequest();
            GetSeatMapResponse response = new GetSeatMapResponse();
            req.StrOrigin = "XCN";
            req.StrDestination = "LWY";
            req.StrFlightId = "fdb8b804-f007-460a-98f3-997ecf4becce";//997ecf4becce
            req.StrBookingClass = "Y";
            req.StrBoardingClass = "Y";
            response = objFlight.GetSeatMap(req);
            if (response != null)
            {
                HttpContext.Current.Response.Write(response.Success + " : " + response.Message + "\n");
                if (response.Success == true)
                {
                    foreach(SeatMap s in response.SeatMaps)
                    {
                        HttpContext.Current.Response.Write("row :" + s.LayoutRow + " column : " + s.LayoutColumn + "\n");
                    }
                }              
            }
        }

        private void PaymentCreditCard()
        {
            BookingServiceProxy objBooking = new BookingServiceProxy();
            BookingPaymentCreditCardRequest req = new BookingPaymentCreditCardRequest();
            BookingPaymentCreditCardResponse response = new BookingPaymentCreditCardResponse();
            string BookingId = System.Guid.NewGuid().ToString();
            string BookingSegmentId = System.Guid.NewGuid().ToString();
            string passId = System.Guid.NewGuid().ToString();

            #region Header
            if (req.Booking == null) req.Booking = new BookingRequest();
            if (req.Booking.Header == null) req.Booking.Header = new BookingHeader();

            req.Booking.Header.BookingId = new Guid(BookingId);
            req.Booking.Header.BookingSourceRcd = "bookingsourcercd1";
            req.Booking.Header.CurrencyRcd = "GBP";
            req.Booking.Header.ClientProfileId = System.Guid.NewGuid();
            req.Booking.Header.BookingNumber = 667590513;
            req.Booking.Header.RecordLocator = "LVN8UJ";
            req.Booking.Header.NumberOfAdults = 0;
            req.Booking.Header.NumberOfChildren = 0;
            req.Booking.Header.NumberOfInfants = 0;
            req.Booking.Header.LanguageRcd = "languagercd1";
            req.Booking.Header.AgencyCode = "B2C";
            req.Booking.Header.ContactName = "contactname1";
            req.Booking.Header.ContactEmail = "contactemail1";
            req.Booking.Header.PhoneMobile = "phonemobile1";
            req.Booking.Header.PhoneHome = "phonehome1";
            req.Booking.Header.PhoneBusiness = "phonebusiness1";
            req.Booking.Header.ReceivedFrom = "receivedfrom1";
            req.Booking.Header.PhoneFax = "phonefax1";
            req.Booking.Header.PhoneSearch = "phonesearch1";
            req.Booking.Header.Comment = "comment";

            req.Booking.Header.NotifyByEmailFlag = 0;
            req.Booking.Header.NotifyBySmsFlag = 0;
            req.Booking.Header.GroupName = "groupname1";
            req.Booking.Header.GroupBookingFlag = 0;
            req.Booking.Header.AgencyName = "agencyname1";
            req.Booking.Header.OwnAgencyFlag = 0;
            req.Booking.Header.WebAgencyFlag = 0;
            req.Booking.Header.ClientNumber = 0;

            req.Booking.Header.Lastname = "aaa";
            req.Booking.Header.Firstname = "bbb";
            req.Booking.Header.City = "bkk";

            req.Booking.Header.CreateName = "createname1";
            req.Booking.Header.Street = "bkk";
            req.Booking.Header.LockDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.Middlename = "midd";
            req.Booking.Header.AddressLine1 = "addressline11";
            req.Booking.Header.AddressLine2 = "addressline21";
            req.Booking.Header.State = "state";
            req.Booking.Header.District = "bkk";
            req.Booking.Header.Province = "province";
            req.Booking.Header.ZipCode = "zip_code1";
            req.Booking.Header.PoBox = "po_box1";
            req.Booking.Header.CountryRcd = "countryrcd1";
            req.Booking.Header.TitleRcd = "titlercd1";
            req.Booking.Header.ExternalPaymentReference = "externalpaymentreference1";
            req.Booking.Header.CreateBy = System.Guid.NewGuid();
            req.Booking.Header.UpdateBy = System.Guid.NewGuid();
            req.Booking.Header.CreateDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.UpdateDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.CostCenter = "cost_center1";
            req.Booking.Header.PurchaseOrder = "purchaseorder1";
            req.Booking.Header.ProjectNumber = "projectnumber1";
            req.Booking.Header.LockName = "lockname1";
            req.Booking.Header.IpAddress = "ipaddress1";
            req.Booking.Header.ApprovalFlag = 0;
            req.Booking.Header.InvoiceReceiver = "invoicereceiver1";
            req.Booking.Header.TaxId = "taxid1";
            req.Booking.Header.NewsletterFlag = 0;
            req.Booking.Header.ContactEmailCc = "contactemailcc1";
            req.Booking.Header.MobileEmail = "mobileemail1";
            req.Booking.Header.VendorRcd = "vendorrcd1";
            req.Booking.Header.BookingDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.NoVatFlag = 0;
            req.Booking.Header.CompanyName = "companyname1";
            req.Booking.Header.BusinessFlag = 0;

            #endregion

            #region FlightSegment
            FlightSegment Segment = new FlightSegment();

            Segment.OriginRcd = "ZRH";
            Segment.DestinationRcd = "AMS";
            Segment.CreateBy = System.Guid.NewGuid();
            Segment.UpdateBy = System.Guid.NewGuid();
            Segment.CreateDateTime = new DateTime(1900, 12, 12);
            Segment.UpdateDateTime = new DateTime(1900, 12, 12);
            Segment.RoutesTot = 0;
            Segment.RoutesAvl = 0;
            Segment.RoutesB2c = 0;
            Segment.RoutesB2b = 0;
            Segment.RoutesB2s = 0;
            Segment.RoutesApi = 0;
            Segment.RoutesB2t = 0;
            Segment.SegmentChangeFeeFlag = true;
            Segment.TransitFlag = true;
            Segment.DirectFlag = true;
            Segment.AvlFlag = true;
            Segment.B2cFlag = true;
            Segment.B2bFlag = true;
            Segment.B2tFlag = true;
            Segment.DayRange = 0;
            Segment.ShowRedressNumberFlag = true;
            Segment.RequirePassengerTitleFlag = true;
            Segment.RequirePassengerGenderFlag = true;
            Segment.RequireDateOfBirthFlag = true;
            Segment.RequireDocumentDetailsFlag = true;
            Segment.RequirePassengerWeightFlag = true;
            Segment.SpecialServiceFeeFlag = true;
            Segment.ShowInsuranceOnWebFlag = true;
            Segment.FlightId = System.Guid.NewGuid();
            Segment.ExchangedSegmentId = System.Guid.NewGuid();
            Segment.AirlineRcd = "GR";
            Segment.FlightNumber = "1234";
            Segment.RefundableFlag = 0;
            Segment.GroupFlag = 0;
            Segment.NonRevenueFlag = 0;
            Segment.EticketFlag = 0;
            Segment.FareReduction = 1;
            Segment.DepartureDate = new DateTime(1900, 12, 12);
            Segment.ArrivalDate = new DateTime(1900, 12, 12);
            Segment.OdOriginRcd = "od_origin_rcd1";
            Segment.OdDestinationRcd = "od_destination_rcd1";
            Segment.WaitlistFlag = 0;
            Segment.OverbookFlag = 0;
            Segment.FlightConnectionId = System.Guid.NewGuid();
            Segment.AdvancedPurchaseFlag = 0;
            Segment.JourneyTime = 1;
            Segment.DepartureTime = 0;
            Segment.ArrivalTime = 0;
            Segment.OriginName = "origin_name";
            Segment.DestinationName = "destination_name";
            Segment.TransitFlightId = System.Guid.NewGuid();
            Segment.TransitFareId = System.Guid.NewGuid();
            Segment.TransitDepartureDate = new DateTime(1900, 12, 12);
            Segment.TransitArrivalDate = new DateTime(1900, 12, 12);
            Segment.FareId = System.Guid.NewGuid();
            Segment.BookingClassRcd = "Y";
            Segment.BoardingClassRcd = "Y";
            Segment.AircraftTypeRcd = "aircraft_type_rcd1";
            Segment.PlannedDepartureTime = 0;
            Segment.PlannedArrivalTime = 0;
            Segment.DepartureDayofweek = 0;
            Segment.ArrivalDayofweek = 0;
            Segment.DepartureMonth = 0;
            Segment.BookingSegmentId = new Guid(BookingSegmentId);
            Segment.SegmentStatusRcd = "segment_status_rcd1";
            Segment.BookingId = new Guid(BookingId);
            Segment.NumberOfUnits = 0;
            Segment.SegmentChangeStatusRcd = "segment_change_status_rcd1";
            Segment.InfoSegmentFlag = 0;
            Segment.HighPriorityWaitlistFlag = 0;
            Segment.SegmentStatusName = "segment_status_name1";
            Segment.SeatmapFlag = 0;
            Segment.TempSeatmapFlag = 0;
            Segment.AllowWebCheckinFlag = 0;
            Segment.CloseWebSalesFlag = 0;
            Segment.ExcludeQuoteFlag = 0;
            Segment.CurrencyRate = 1;
            Segment.OpenSequence = 0;
            Segment.NumberOfStops = 0;

            IList<FlightSegment> segmentList = new List<FlightSegment>();
            segmentList.Add(Segment);
            req.Booking.FlightSegments = segmentList;

            #endregion

            #region Passenger

            Passenger p = new Passenger();
            p.TitleRcd = "title_rcd1";
            p.Lastname = "lastname1";
            p.Firstname = "firstname1";
            p.Middlename = "middlename1";
            p.NationalityRcd = "nationality_rcd1";
            p.PassengerWeight = 1;
            p.GenderTypeRcd = "gender_type_rcd1";
            p.PassengerTypeRcd = "passenger_type_rcd1";
            p.ClientProfileId = System.Guid.NewGuid();
            p.ClientNumber = 0;
            p.AddressLine1 = "address_line11";
            p.AddressLine2 = "address_line21";
            p.State = "state1";
            p.District = "district1";
            p.Province = "province1";
            p.ZipCode = "zip_code1";
            p.PoBox = "po_box1";
            p.CountryRcd = "country_rcd1";
            p.Street = "street1";
            p.City = "city1";
            p.DocumentTypeRcd = "document_type_rcd1";
            p.DocumentNumber = "document_number1";
            p.ResidenceCountryRcd = "residence_country_rcd1";
            p.PassportNumber = "passport_number1";
            p.PassportIssueDate = new DateTime(1900, 12, 12);
            p.PassportExpiryDate = new DateTime(1900, 12, 12);
            p.PassportIssuePlace = "passport_issue_place1";
            p.PassportBirthPlace = "passport_birth_place1";
            p.DateOfBirth = new DateTime(1900, 12, 12);
            p.PassportIssueCountryRcd = "passport_issue_country_rcd1";
            p.ContactName = "contact_name1";
            p.ContactEmail = "contact_email1";
            p.MobileEmail = "mobile_email1";
            p.PhoneMobile = "phone_mobile1";
            p.PhoneHome = "phone_home1";
            p.PhoneFax = "phone_fax1";
            p.PhoneBusiness = "phone_business1";
            p.ReceivedFrom = "received_from1";
            p.CreateBy = System.Guid.NewGuid();
            p.CreateDateTime = new DateTime(1900, 12, 12);
            p.UpdateBy = System.Guid.NewGuid();
            p.UpdateDateTime = new DateTime(1900, 12, 12);
            p.PassengerId = new Guid(passId);
            p.BookingId = new Guid(BookingId);
            p.PassengerProfileId = System.Guid.NewGuid();
            p.EmployeeNumber = "employee_number1";
            p.WheelchairFlag = 0;
            p.VipFlag = 0;
            p.MemberLevelRcd = "member_level_rcd1";
            p.MemberNumber = "member_number1";
            p.WindowSeatFlag = 0;
            p.RedressNumber = "redress_number1";
            p.PnrName = "pnr_name1";

            IList<Passenger> pass = new List<Passenger>();
            pass.Add(p);
            req.Booking.Passengers = pass;

            #endregion

            #region Fee
            Fee f = new Fee();
            f.BookingFeeId = System.Guid.NewGuid();
            f.FeeAmount = 1;
            f.BookingId = new Guid(BookingId);
            f.PassengerId = System.Guid.NewGuid();
            f.CurrencyRcd = "GBP";
            f.AcctFeeAmount = 1;
            f.FeeId = System.Guid.NewGuid();
            f.VatPercentage = 1;
            f.FeeAmountIncl = 1;
            f.AcctFeeAmountIncl = 1;
            f.FeeRcd = "fee_rcd1";
            f.DisplayName = "display_name1";
            f.AccountFeeBy = System.Guid.NewGuid();
            f.AccountFeeDateTime = new DateTime(1900, 12, 12);
            f.VoidDateTime = new DateTime(1900, 12, 12);
            f.VoidBy = System.Guid.NewGuid();
            f.PaymentAmount = 1;
            f.CreateBy = System.Guid.NewGuid();
            f.CreateDateTime = new DateTime(1900, 12, 12);
            f.UpdateBy = System.Guid.NewGuid();
            f.UpdateDateTime = new DateTime(1900, 12, 12);
            f.BookingSegmentId = System.Guid.NewGuid();
            f.AgencyCode = "B2C";
            f.PassengerSegmentServiceId = System.Guid.NewGuid();
            f.OriginRcd = "origin_rcd1";
            f.DestinationRcd = "destination_rcd1";
            f.OdOriginRcd = "od_origin_rcd1";
            f.OdDestinationRcd = "od_destination_rcd1";
            f.NumberOfUnits = 0;
            f.TotalAmount = 1;
            f.TotalAmountIncl = 1;
            f.ManualFeeFlag = 0;
            f.OdFlag = 0;
            f.SkipFareAllowanceFlag = 0;
            f.MinimumFeeAmountFlag = 0;
            f.FeePercentage = 1;
            f.TotalFeeAmount = 1;
            f.TotalFeeAmountIncl = 1;
            f.ChargeCurrencyRcd = "GBP";
            f.ChargeAmount = 1;
            f.ChargeAmountIncl = 1;
            f.DocumentDateTime = new DateTime(1900, 12, 12);
            f.BaggageFeeOptionId = System.Guid.NewGuid();

            IList<Fee> fee = new List<Fee>();
            fee.Add(f);
            req.Booking.Fees = fee;

            #endregion

            #region Remark
            Remark r = new Remark();
            r.BookingRemarkId = System.Guid.NewGuid();
            r.RemarkTypeRcd = "TKTL";
            r.TimelimitDateTime = new DateTime(1900, 12, 12);
            r.CompleteFlag = 0;
            r.RemarkText = "remark_text1";
            r.BookingId = new Guid(BookingId);
            r.AddedBy = "addby";
            r.ClientProfileId = System.Guid.NewGuid();
            r.AgencyCode = "B2C";
            r.CreateBy = System.Guid.NewGuid();
            r.CreateDateTime = new DateTime(1900, 12, 12);
            r.UpdateBy = System.Guid.NewGuid();
            r.UpdateDateTime = new DateTime(1900, 12, 12);
            r.SystemFlag = 0;
            r.UtcTimelimitDateTime = new DateTime(1900, 12, 12);
            r.ProtectedFlag = 0;
            r.WarningFlag = 0;
            r.ProcessMessageFlag = 0;

            IList<Remark> remark = new List<Remark>();
            remark.Add(r);
            req.Booking.Remarks = remark;

            #endregion

            #region Tax
            Tax t = new Tax();
            t.AcctAmountIncl = 1;
            t.SalesAmount = 1;
            t.SalesAmountIncl = 1;
            t.TaxRcd = "UB";
            t.PassengerId = new Guid(passId);
            t.TaxId = new Guid("3C3B5805-DE7C-11DE-9229-00145E3CC069");
            t.BookingSegmentId = new Guid(BookingSegmentId);
            t.TaxCurrencyRcd = "GBP";
            t.SalesCurrencyRcd = "GBP";
            t.DisplayName = "display_name1";
            t.SummarizeUp = "summarize_up1";
            t.CoverageType = "coverage_type1";
            t.CreateBy = System.Guid.NewGuid();
            t.UpdateBy = System.Guid.NewGuid();
            t.CreateDateTime = new DateTime(1900, 12, 12);
            t.UpdateDateTime = new DateTime(1900, 12, 12);
            t.VatPercentage = 1;
            IList<Tax> tax = new List<Tax>();
            tax.Add(t);
            req.Booking.Taxs = tax;

            #endregion

            #region Payment
            Message.Booking.Payment pay = new Message.Booking.Payment();

            pay.BookingPaymentId = System.Guid.NewGuid();
            pay.BookingSegmentId = new Guid(BookingSegmentId);
            pay.BookingId = new Guid(BookingId);
            pay.VoucherPaymentId = Guid.Empty;
            pay.FormOfPaymentRcd = "CC";
            pay.CurrencyRcd = "GBP";
            pay.ReceiveCurrencyRcd = "GBP";
            pay.AgencyCode = "B2C";
            pay.DebitAgencyCode = "B2C";
            pay.PaymentAmount = 100;
            pay.ReceivePaymentAmount = 10;
            pay.AcctPaymentAmount = 100;
            pay.PaymentBy = System.Guid.NewGuid();
            pay.PaymentDateTime = DateTime.Now;
            pay.PaymentDueDateTime = DateTime.Now;
            pay.DocumentAmount = 10;
            pay.VoidBy = Guid.Empty;
            pay.ExpiryMonth = 6;
            pay.ExpiryYear = 2019;
            pay.VoidDateTime = DateTime.MinValue;
            pay.RecordLocator = "LVN8UJ";
            pay.NameOnCard = "sss ssss";
            pay.DocumentNumber = "4444333322221111";
            pay.FormOfPaymentSubtypeRcd = "VISA";
            pay.City = "city1";
            pay.State = "state1";
            pay.Street = "street1";
            pay.AddressLine1 = "address_line11";
            pay.ZipCode = "zip_code1";
            pay.CountryRcd = "country_rcd1";
            pay.CreateBy = System.Guid.NewGuid();
            pay.CreateDateTime = DateTime.Now;
            pay.UpdateBy = System.Guid.NewGuid();
            pay.UpdateDateTime = DateTime.Now;
            pay.PosIndicator = "pos_indicator1";
            pay.IssueMonth = 5;
            pay.IssueYear = 2014;
            pay.IssueNumber = "4444333322221111";
            pay.ClientProfileId = Guid.Empty;
            pay.TransactionReference = "123456789";
            pay.ApprovalCode = "123";
            pay.BankName = "name";
            pay.BankCode = "code";
            pay.BankIban = "iban";
            pay.IpAddress = "194.66.82.11";
            pay.PaymentReference = "123456789";
            pay.AllocatedAmount = 10;
            pay.PaymentTypeRcd = "CC";
            pay.RefundedAmount = 0;
            

            IList<Message.Booking.Payment> pp = new List<Message.Booking.Payment>();
            pp.Add(pay);
            req.Booking.Payments = pp;

            #endregion

            #region Service
            PassengerService s = new PassengerService();
            s.PassengerSegmentServiceId = System.Guid.NewGuid();
            s.PassengerId = new Guid(passId);
            s.BookingSegmentId = new Guid(BookingSegmentId);
            s.SpecialServiceStatusRcd = "special_service_status_rcd1";
            s.SpecialServiceChangeStatusRcd = "special_service_change_status_rcd1";
            s.SpecialServiceRcd = "special_service_rcd1";
            s.ServiceText = "service_text1";
            s.CreateBy = System.Guid.NewGuid();
            s.CreateDateTime = new DateTime(1900, 12, 12);
            s.UpdateBy = System.Guid.NewGuid();
            s.UpdateDateTime = new DateTime(1900, 12, 12);
            s.NumberOfUnits = 0;
            s.OriginRcd = "origin_rcd1";
            s.DestinationRcd = "destination_rcd1";
            s.DisplayName = "display_name1";
            s.CutOffTime = 0;
            s.ServiceOnRequestFlag = 0;
            s.TextAllowedFlag = 0;
            s.ManifestFlag = 0;
            s.TextRequiredFlag = 0;
            s.IncludePassengerNameFlag = 0;
            s.IncludeFlightSegmentFlag = 0;
            s.IncludeActionCodeFlag = 0;
            s.IncludeNumberOfServiceFlag = 0;
            s.IncludeCateringFlag = 0;
            s.IncludePassengerAssistanceFlag = 0;
            s.ServiceSupportedFlag = 0;
            s.SendInterlineReplyFlag = 0;
            s.InventoryControlFlag = 0;

            IList<PassengerService> ser = new List<PassengerService>();
            ser.Add(s);
            req.Booking.Services = ser;

            #endregion

            #region Mapping
            Mapping m = new Mapping();
            m.OriginRcd = "origin_rcd1";
            m.DestinationRcd = "destination_rcd1";
            m.CreateBy = System.Guid.NewGuid();
            m.UpdateBy = System.Guid.NewGuid();
            m.CreateDateTime = new DateTime(1900, 12, 12);
            m.UpdateDateTime = new DateTime(1900, 12, 12);
            m.CurrencyRcd = "GBP";
            m.RoutesTot = 0;
            m.RoutesAvl = 0;
            m.RoutesB2c = 0;
            m.RoutesB2b = 0;
            m.RoutesB2s = 0;
            m.RoutesApi = 0;
            m.RoutesB2t = 0;
            m.SegmentChangeFeeFlag = true;
            m.TransitFlag = true;
            m.DirectFlag = true;
            m.AvlFlag = true;
            m.B2cFlag = true;
            m.B2bFlag = true;
            m.B2tFlag = true;
            m.DayRange = 0;
            m.ShowRedressNumberFlag = true;
            m.RequirePassengerTitleFlag = true;
            m.RequirePassengerGenderFlag = true;
            m.RequireDateOfBirthFlag = true;
            m.RequireDocumentDetailsFlag = true;
            m.RequirePassengerWeightFlag = true;
            m.SpecialServiceFeeFlag = true;
            m.ShowInsuranceOnWebFlag = true;
            m.FlightId = System.Guid.NewGuid();
            m.ExchangedSegmentId = System.Guid.NewGuid();
            m.AirlineRcd = "GR";
            m.FlightNumber = "1234";
            m.RefundableFlag = 0;
            m.GroupFlag = 0;
            m.NonRevenueFlag = 0;
            m.EticketFlag = 0;
            m.FareReduction = 1;
            m.DepartureDate = new DateTime(1900, 12, 12);
            m.ArrivalDate = new DateTime(1900, 12, 12);
            m.OdOriginRcd = "od_origin_rcd1";
            m.OdDestinationRcd = "od_destination_rcd1";
            m.WaitlistFlag = 0;
            m.OverbookFlag = 0;
            m.FlightConnectionId = System.Guid.NewGuid();
            m.AdvancedPurchaseFlag = 0;
            m.JourneyTime = 0;
            m.DepartureTime = 0;
            m.ArrivalTime = 0;
            m.TransitFlightId = System.Guid.NewGuid();
            m.TransitFareId = System.Guid.NewGuid();
            m.TransitDepartureDate = new DateTime(1900, 12, 12);
            m.TransitArrivalDate = new DateTime(1900, 12, 12);
            m.FareId = System.Guid.NewGuid();
            m.BookingClassRcd = "Y";
            m.BoardingClassRcd = "Y";
            m.PlannedDepartureTime = 0;
            m.PlannedArrivalTime = 0;
            m.DepartureDayofweek = 0;
            m.ArrivalDayofweek = 0;
            m.DepartureMonth = 0;
            m.BookingSegmentId = new System.Guid(BookingSegmentId);
            m.BookingId = new System.Guid(BookingId);
            m.NumberOfUnits = 0;
            m.InfoSegmentFlag = 0;
            m.HighPriorityWaitlistFlag = 0;
            m.SeatmapFlag = 0;
            m.TempSeatmapFlag = 0;
            m.AllowWebCheckinFlag = 0;
            m.CloseWebSalesFlag = 0;
            m.ExcludeQuoteFlag = 0;
            m.CurrencyRate = 1;
            m.OpenSequence = 0;
            m.NumberOfStops = 0;
            m.PassengerId = new System.Guid(passId);
            m.Lastname = "lastname1";
            m.Firstname = "firstname1";
            m.GenderTypeRcd = "gender_type_rcd1";
            m.TitleRcd = "title_rcd1";
            m.PassengerTypeRcd = "passenger_type_rcd1";
            m.DateOfBirth = new DateTime(1900, 12, 12);
            m.PassengerStatusRcd = "passenger_status_rcd1";
            m.Middlename = "middlename1";
            m.DocumentTypeRcd = "document_type_rcd1";
            m.PassportNumber = "passport_number1";
            m.PassportIssuePlace = "passport_issue_place1";
            m.PassportIssueDate = new DateTime(1900, 12, 12);
            m.PassportExpiryDate = new DateTime(1900, 12, 12);
            m.Sendmail = "sendmail1";
            m.ClientNumber = 0;
            m.PassengerProfileId = System.Guid.NewGuid();
            m.ClientProfileId = System.Guid.NewGuid();
            m.NationalityRcd = "nationality_rcd1";
            m.ContactEmail = "contact_email1";
            m.InventoryClassRcd = "inventory_class_rcd1";
            m.SeatNumber = "seat_number1";
            m.SeatRow = 0;
            m.SeatColumn = "seat_column1";
            m.NetTotal = 1;
            m.TaxAmount = 1;
            m.FareAmount = 1;
            m.YqAmount = 1;
            m.BaseTicketingFeeAmount = 1;
            m.TicketingFeeAmount = 1;
            m.ReservationFeeAmount = 1;
            m.CommissionAmount = 1;
            m.FareVat = 1;
            m.TaxVat = 1;
            m.YqVat = 1;
            m.TicketingFeeVat = 1;
            m.ReservationFeeVat = 1;
            m.FareAmountIncl = 1;
            m.TaxAmountIncl = 1;
            m.YqAmountIncl = 1;
            m.PublicFareAmountIncl = 1;
            m.PublicFareAmount = 1;
            m.CommissionAmountIncl = 1;
            m.CommissionPercentage = 1;
            m.TicketingFeeAmountIncl = 1;
            m.ReservationFeeAmountIncl = 1;
            m.AcctNetTotal = 1;
            m.AcctTaxAmount = 1;
            m.AcctFareAmount = 1;
            m.AcctYqAmount = 1;
            m.AcctTicketingFeeAmount = 1;
            m.AcctReservationFeeAmount = 1;
            m.AcctCommissionAmount = 1;
            m.AcctFareAmountIncl = 1;
            m.AcctTaxAmountIncl = 1;
            m.AcctYqAmountIncl = 1;
            m.AcctCommissionAmountIncl = 1;
            m.AcctTicketingFeeAmountIncl = 1;
            m.AcctReservationFeeAmountIncl = 1;
            m.FareCode = "fare_code1";
            m.FareDateTime = new DateTime(1900, 12, 12);
            m.ETicketFlag = 0;
            m.StandbyFlag = 0;
            m.TransferableFareFlag = 0;
            m.AgencyCode = "B2C";
            m.TicketNumber = "0";
            m.TicketingDateTime = new DateTime(1900, 12, 12);
            m.TicketingBy = System.Guid.NewGuid();
            m.CheckInSequence = 0;
            m.GroupSequence = 0;
            m.UnloadBy = System.Guid.NewGuid();
            m.UnloadDateTime = new DateTime(1900, 12, 12);
            m.NumberOfPieces = 0;
            m.BaggageWeight = 1;
            m.ExcessBaggageWeight = 1;
            m.CheckInBaggageWeight = 1;
            m.CheckInBy = System.Guid.NewGuid();
            m.CancelBy = System.Guid.NewGuid();
            m.ExchangedDateTime = new DateTime(1900, 12, 12);
            m.CancelDateTime = new DateTime(1900, 12, 12);
            m.RestrictionText = "restriction_text1";
            m.EndorsementText = "endorsement_text1";
            m.ExcludePricingFlag = 0;
            m.CreateName = "create_name1";
            m.UpdateName = "update_name1";
            m.VoidDateTime = new DateTime(1900, 12, 12);
            m.RefundCharge = 1;
            m.RefundableAmount = 1;
            m.RefundDateTime = new DateTime(1900, 12, 12);
            m.PaymentAmount = 1;
            m.MappingSort = 0;
            m.CheckInDateTime = new DateTime(1900, 12, 12);
            m.OnwardDepartureDate = new DateTime(1900, 12, 12);
            m.ETicketStatus = "e_ticket_status1";
            m.HandLuggageFlag = 0;
            m.HandNumberOfPieces = 0;
            m.HandBaggageWeight = 1;
            m.FreeSeatingFlag = 0;
            m.FareTypeRcd = "fare_type_rcd1";
            m.RedemptionPoints = 0;
            m.SeatFeeRcd = "seat_fee_rcd1";
            m.RefundWithChargeHours = 0;
            m.RefundNotPossibleHours = 0;
            m.DutyTravelFlag = 0;
            m.NotValidAfterDate = new DateTime(1900, 12, 12);
            m.NotValidBeforeDate = new DateTime(1900, 12, 12);
            m.AdvancedSeatingFlag = 0;
            m.FareColumn = 0;
            m.PieceAllowance = 1;
            m.BoardingTime = 0;
            m.ItFareFlag = 0;

            IList<Mapping> map = new List<Mapping>();
            map.Add(m);
            req.Booking.Mappings = map;

            #endregion

            Fee paymentFee = new Fee();
            //paymentFee.BookingFeeId = System.Guid.NewGuid();
            //paymentFee.FeeAmount = 1;
            //paymentFee.BookingId = new Guid(BookingId);
            //paymentFee.PassengerId = System.Guid.NewGuid();
            //paymentFee.CurrencyRcd = "GBP";
            //paymentFee.AcctFeeAmount = 1;
            //paymentFee.FeeId = System.Guid.NewGuid();
            //paymentFee.VatPercentage = 1;
            //paymentFee.FeeAmountIncl = 1;
            //paymentFee.AcctFeeAmountIncl = 1;
            //paymentFee.FeeRcd = "fee_rcd1";
            //paymentFee.DisplayName = "display_name1";
            //paymentFee.AccountFeeBy = System.Guid.NewGuid();
            //paymentFee.AccountFeeDateTime = new DateTime(1900, 12, 12);
            //paymentFee.VoidDateTime = new DateTime(1900, 12, 12);
            //paymentFee.VoidBy = System.Guid.NewGuid();
            //paymentFee.PaymentAmount = 1;
            //paymentFee.CreateBy = System.Guid.NewGuid();
            //paymentFee.CreateDateTime = new DateTime(1900, 12, 12);
            //paymentFee.UpdateBy = System.Guid.NewGuid();
            //paymentFee.UpdateDateTime = new DateTime(1900, 12, 12);
            //paymentFee.BookingSegmentId = System.Guid.NewGuid();
            //paymentFee.AgencyCode = "B2C";
            //paymentFee.PassengerSegmentServiceId = System.Guid.NewGuid();
            //paymentFee.OriginRcd = "origin_rcd1";
            //paymentFee.DestinationRcd = "destination_rcd1";
            //paymentFee.OdOriginRcd = "od_origin_rcd1";
            //paymentFee.OdDestinationRcd = "od_destination_rcd1";
            //paymentFee.NumberOfUnits = 0;
            //paymentFee.TotalAmount = 1;
            //paymentFee.TotalAmountIncl = 1;
            //paymentFee.ManualFeeFlag = 0;
            //paymentFee.OdFlag = 0;
            //paymentFee.SkipFareAllowanceFlag = 0;
            //paymentFee.MinimumFeeAmountFlag = 0;
            //paymentFee.FeePercentage = 1;
            //paymentFee.TotalFeeAmount = 1;
            //paymentFee.TotalFeeAmountIncl = 1;
            //paymentFee.ChargeCurrencyRcd = "GBP";
            //paymentFee.ChargeAmount = 1;
            //paymentFee.ChargeAmountIncl = 1;
            //paymentFee.DocumentDateTime = new DateTime(1900, 12, 12);
            //paymentFee.BaggageFeeOptionId = System.Guid.NewGuid();
            req.PaymentFee = paymentFee;

            req.AuthenticationToken = "";
            req.BookingReference = "";
            req.CommerceIndicator = "";
            req.RequestSource = "";
            req.SecurityToken = "";

            response = objBooking.BookingPaymentCreditCard(req);

            if (response != null)
            {
                HttpContext.Current.Response.Write(response.Success + " : " + response.Message);
            }
        }

        private void ReadVoucher()
        {
            BookingServiceProxy objBooking = new BookingServiceProxy();
            //BookingPaymentVoucherRequest req = new BookingPaymentVoucherRequest();
            //BookingPaymentVoucherResponse response = new BookingPaymentVoucherResponse();
            //Voucher voucher = new Voucher();
            //voucher.VoucherNumber = "3700002738";
            //voucher.VoucherId = "";// "2eba503d-7f85-4ec0-9d36-bc71ac1aec62";
            //voucher.VoucherStatusRcd = "";// "OPEN";
            //voucher.RecipientName = "";// "soi";
            //voucher.FormOfPaymentSubtypeRcd = "";// "VOUCHER";
            //voucher.CurrencyRcd = "USD";
            //voucher.VoucherPassword = "2750";
            //voucher.AgencyCode = "";// "B2C";

            //req.Voucher = voucher;
            //req.RecordLocator = "";
            //req.ClientProfileId = "";
            //req.IncludeExpiredVoucher = false;
            //req.IncludeFareOnly = false;
            //req.IncludeOpenVoucher = true;
            //req.IncludeRefundable = false;
            //req.IncludeUsedVoucher = false;
            //req.IncludeVoidedVoucher = false;
            //req.Write = false;
            BookingReadVoucherRequest req = new BookingReadVoucherRequest();
            BookingReadVoucherResponse response = new BookingReadVoucherResponse();
            Voucher voucher = new Voucher();
            voucher.VoucherNumber = "3700002738";
            voucher.VoucherId = "";// "2eba503d-7f85-4ec0-9d36-bc71ac1aec62";
            voucher.VoucherStatusRcd = "";// "OPEN";
            voucher.RecipientName = "";// "soi";
            voucher.FormOfPaymentSubtypeRcd = "";// "VOUCHER";
            voucher.CurrencyRcd = "USD";
            voucher.VoucherPassword = "2750";
            voucher.AgencyCode = "";// "B2C";
            List<Voucher> vouchers = new List<Voucher>();
            vouchers.Add(voucher);
            req.Vouchers = vouchers;

            response = objBooking.BookingReadVoucher(req);
            
            if (response != null)
            {
                HttpContext.Current.Response.Write(response.Success + " : " + response.Message);
            }
        }

        private void PaymentVoucher()
        {
            BookingServiceProxy objBooking = new BookingServiceProxy();
            BookingPaymentVoucherRequest req = new BookingPaymentVoucherRequest();
            BookingPaymentVoucherResponse response = new BookingPaymentVoucherResponse();
            //BookingMessage booking = new BookingMessage();
            //Message.Booking.Payment p = new Message.Booking.Payment();
            //List<Message.Booking.Payment> ps = new List<Message.Booking.Payment>();
            //ps.Add(p);
            //booking.Payments = ps;
            //req.Booking = booking;

            string BookingId = System.Guid.NewGuid().ToString();
            string BookingSegmentId = System.Guid.NewGuid().ToString();
            string passId = System.Guid.NewGuid().ToString();

            #region Header
            if (req.Booking == null) req.Booking = new BookingMessage();
            if (req.Booking.Header == null) req.Booking.Header = new BookingHeader();

            req.Booking.Header.BookingId = new Guid(BookingId);
            req.Booking.Header.BookingSourceRcd = "bookingsourcercd1";
            req.Booking.Header.CurrencyRcd = "THB";
            req.Booking.Header.ClientProfileId = System.Guid.NewGuid();
            req.Booking.Header.BookingNumber = 0;
            req.Booking.Header.RecordLocator = "";
            req.Booking.Header.NumberOfAdults = 0;
            req.Booking.Header.NumberOfChildren = 0;
            req.Booking.Header.NumberOfInfants = 0;
            req.Booking.Header.LanguageRcd = "languagercd1";
            req.Booking.Header.AgencyCode = "B2C";
            req.Booking.Header.ContactName = "contactname1";
            req.Booking.Header.ContactEmail = "contactemail1";
            req.Booking.Header.PhoneMobile = "phonemobile1";
            req.Booking.Header.PhoneHome = "phonehome1";
            req.Booking.Header.PhoneBusiness = "phonebusiness1";
            req.Booking.Header.ReceivedFrom = "receivedfrom1";
            req.Booking.Header.PhoneFax = "phonefax1";
            req.Booking.Header.PhoneSearch = "phonesearch1";
            req.Booking.Header.Comment = "comment";

            req.Booking.Header.NotifyByEmailFlag = 0;
            req.Booking.Header.NotifyBySmsFlag = 0;
            req.Booking.Header.GroupName = "groupname1";
            req.Booking.Header.GroupBookingFlag = 0;
            req.Booking.Header.AgencyName = "agencyname1";
            req.Booking.Header.OwnAgencyFlag = 0;
            req.Booking.Header.WebAgencyFlag = 0;
            req.Booking.Header.ClientNumber = 0;

            req.Booking.Header.Lastname = "aaa";
            req.Booking.Header.Firstname = "bbb";
            req.Booking.Header.City = "bkk";

            req.Booking.Header.CreateName = "createname1";
            req.Booking.Header.Street = "bkk";
            req.Booking.Header.LockDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.Middlename = "midd";
            req.Booking.Header.AddressLine1 = "addressline11";
            req.Booking.Header.AddressLine2 = "addressline21";
            req.Booking.Header.State = "state";
            req.Booking.Header.District = "bkk";
            req.Booking.Header.Province = "province";
            req.Booking.Header.ZipCode = "zip_code1";
            req.Booking.Header.PoBox = "po_box1";
            req.Booking.Header.CountryRcd = "countryrcd1";
            req.Booking.Header.TitleRcd = "titlercd1";
            req.Booking.Header.ExternalPaymentReference = "externalpaymentreference1";
            req.Booking.Header.CreateBy = System.Guid.NewGuid();
            req.Booking.Header.UpdateBy = System.Guid.NewGuid();
            req.Booking.Header.CreateDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.UpdateDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.CostCenter = "cost_center1";
            req.Booking.Header.PurchaseOrder = "purchaseorder1";
            req.Booking.Header.ProjectNumber = "projectnumber1";
            req.Booking.Header.LockName = "lockname1";
            req.Booking.Header.IpAddress = "ipaddress1";
            req.Booking.Header.ApprovalFlag = 0;
            req.Booking.Header.InvoiceReceiver = "invoicereceiver1";
            req.Booking.Header.TaxId = "taxid1";
            req.Booking.Header.NewsletterFlag = 0;
            req.Booking.Header.ContactEmailCc = "contactemailcc1";
            req.Booking.Header.MobileEmail = "mobileemail1";
            req.Booking.Header.VendorRcd = "vendorrcd1";
            req.Booking.Header.BookingDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.NoVatFlag = 0;
            req.Booking.Header.CompanyName = "companyname1";
            req.Booking.Header.BusinessFlag = 0;

            #endregion

            #region FlightSegment
            FlightSegment Segment = new FlightSegment();

            Segment.OriginRcd = "ZRH";
            Segment.DestinationRcd = "AMS";
            Segment.CreateBy = System.Guid.NewGuid();
            Segment.UpdateBy = System.Guid.NewGuid();
            Segment.CreateDateTime = new DateTime(1900, 12, 12);
            Segment.UpdateDateTime = new DateTime(1900, 12, 12);
            Segment.RoutesTot = 0;
            Segment.RoutesAvl = 0;
            Segment.RoutesB2c = 0;
            Segment.RoutesB2b = 0;
            Segment.RoutesB2s = 0;
            Segment.RoutesApi = 0;
            Segment.RoutesB2t = 0;
            Segment.SegmentChangeFeeFlag = true;
            Segment.TransitFlag = true;
            Segment.DirectFlag = true;
            Segment.AvlFlag = true;
            Segment.B2cFlag = true;
            Segment.B2bFlag = true;
            Segment.B2tFlag = true;
            Segment.DayRange = 0;
            Segment.ShowRedressNumberFlag = true;
            Segment.RequirePassengerTitleFlag = true;
            Segment.RequirePassengerGenderFlag = true;
            Segment.RequireDateOfBirthFlag = true;
            Segment.RequireDocumentDetailsFlag = true;
            Segment.RequirePassengerWeightFlag = true;
            Segment.SpecialServiceFeeFlag = true;
            Segment.ShowInsuranceOnWebFlag = true;
            Segment.FlightId = System.Guid.NewGuid();
            Segment.ExchangedSegmentId = System.Guid.NewGuid();
            Segment.AirlineRcd = "GR";
            Segment.FlightNumber = "1234";
            Segment.RefundableFlag = 0;
            Segment.GroupFlag = 0;
            Segment.NonRevenueFlag = 0;
            Segment.EticketFlag = 0;
            Segment.FareReduction = 1;
            Segment.DepartureDate = new DateTime(1900, 12, 12);
            Segment.ArrivalDate = new DateTime(1900, 12, 12);
            Segment.OdOriginRcd = "od_origin_rcd1";
            Segment.OdDestinationRcd = "od_destination_rcd1";
            Segment.WaitlistFlag = 0;
            Segment.OverbookFlag = 0;
            Segment.FlightConnectionId = System.Guid.NewGuid();
            Segment.AdvancedPurchaseFlag = 0;
            Segment.JourneyTime = 1;
            Segment.DepartureTime = 0;
            Segment.ArrivalTime = 0;
            Segment.OriginName = "origin_name";
            Segment.DestinationName = "destination_name";
            Segment.TransitFlightId = System.Guid.NewGuid();
            Segment.TransitFareId = System.Guid.NewGuid();
            Segment.TransitDepartureDate = new DateTime(1900, 12, 12);
            Segment.TransitArrivalDate = new DateTime(1900, 12, 12);
            Segment.FareId = System.Guid.NewGuid();
            Segment.BookingClassRcd = "Y";
            Segment.BoardingClassRcd = "Y";
            Segment.AircraftTypeRcd = "aircraft_type_rcd1";
            Segment.PlannedDepartureTime = 0;
            Segment.PlannedArrivalTime = 0;
            Segment.DepartureDayofweek = 0;
            Segment.ArrivalDayofweek = 0;
            Segment.DepartureMonth = 0;
            Segment.BookingSegmentId = new Guid(BookingSegmentId);
            Segment.SegmentStatusRcd = "segment_status_rcd1";
            Segment.BookingId = new Guid(BookingId);
            Segment.NumberOfUnits = 0;
            Segment.SegmentChangeStatusRcd = "segment_change_status_rcd1";
            Segment.InfoSegmentFlag = 0;
            Segment.HighPriorityWaitlistFlag = 0;
            Segment.SegmentStatusName = "segment_status_name1";
            Segment.SeatmapFlag = 0;
            Segment.TempSeatmapFlag = 0;
            Segment.AllowWebCheckinFlag = 0;
            Segment.CloseWebSalesFlag = 0;
            Segment.ExcludeQuoteFlag = 0;
            Segment.CurrencyRate = 1;
            Segment.OpenSequence = 0;
            Segment.NumberOfStops = 0;

            IList<FlightSegment> segmentList = new List<FlightSegment>();
            segmentList.Add(Segment);
            req.Booking.FlightSegments = segmentList;

            #endregion

            #region Passenger

            Passenger p = new Passenger();
            p.TitleRcd = "title_rcd1";
            p.Lastname = "lastname1";
            p.Firstname = "firstname1";
            p.Middlename = "middlename1";
            p.NationalityRcd = "nationality_rcd1";
            p.PassengerWeight = 1;
            p.GenderTypeRcd = "gender_type_rcd1";
            p.PassengerTypeRcd = "passenger_type_rcd1";
            p.ClientProfileId = System.Guid.NewGuid();
            p.ClientNumber = 0;
            p.AddressLine1 = "address_line11";
            p.AddressLine2 = "address_line21";
            p.State = "state1";
            p.District = "district1";
            p.Province = "province1";
            p.ZipCode = "zip_code1";
            p.PoBox = "po_box1";
            p.CountryRcd = "country_rcd1";
            p.Street = "street1";
            p.City = "city1";
            p.DocumentTypeRcd = "document_type_rcd1";
            p.DocumentNumber = "document_number1";
            p.ResidenceCountryRcd = "residence_country_rcd1";
            p.PassportNumber = "passport_number1";
            p.PassportIssueDate = new DateTime(1900, 12, 12);
            p.PassportExpiryDate = new DateTime(1900, 12, 12);
            p.PassportIssuePlace = "passport_issue_place1";
            p.PassportBirthPlace = "passport_birth_place1";
            p.DateOfBirth = new DateTime(1900, 12, 12);
            p.PassportIssueCountryRcd = "passport_issue_country_rcd1";
            p.ContactName = "contact_name1";
            p.ContactEmail = "contact_email1";
            p.MobileEmail = "mobile_email1";
            p.PhoneMobile = "phone_mobile1";
            p.PhoneHome = "phone_home1";
            p.PhoneFax = "phone_fax1";
            p.PhoneBusiness = "phone_business1";
            p.ReceivedFrom = "received_from1";
            p.CreateBy = System.Guid.NewGuid();
            p.CreateDateTime = new DateTime(1900, 12, 12);
            p.UpdateBy = System.Guid.NewGuid();
            p.UpdateDateTime = new DateTime(1900, 12, 12);
            p.PassengerId = new Guid(passId);
            p.BookingId = new Guid(BookingId);
            p.PassengerProfileId = System.Guid.NewGuid();
            p.EmployeeNumber = "employee_number1";
            p.WheelchairFlag = 0;
            p.VipFlag = 0;
            p.MemberLevelRcd = "member_level_rcd1";
            p.MemberNumber = "member_number1";
            p.WindowSeatFlag = 0;
            p.RedressNumber = "redress_number1";
            p.PnrName = "pnr_name1";

            IList<Passenger> pass = new List<Passenger>();
            pass.Add(p);
            req.Booking.Passengers = pass;

            #endregion

            #region Fee
            Fee f = new Fee();
            f.BookingFeeId = System.Guid.NewGuid();
            f.FeeAmount = 1;
            f.BookingId = new Guid(BookingId);
            f.PassengerId = System.Guid.NewGuid();
            f.CurrencyRcd = "THB";
            f.AcctFeeAmount = 1;
            f.FeeId = System.Guid.NewGuid();
            f.VatPercentage = 1;
            f.FeeAmountIncl = 1;
            f.AcctFeeAmountIncl = 1;
            f.FeeRcd = "fee_rcd1";
            f.DisplayName = "display_name1";
            f.AccountFeeBy = System.Guid.NewGuid();
            f.AccountFeeDateTime = new DateTime(1900, 12, 12);
            f.VoidDateTime = new DateTime(1900, 12, 12);
            f.VoidBy = System.Guid.NewGuid();
            f.PaymentAmount = 1;
            f.CreateBy = System.Guid.NewGuid();
            f.CreateDateTime = new DateTime(1900, 12, 12);
            f.UpdateBy = System.Guid.NewGuid();
            f.UpdateDateTime = new DateTime(1900, 12, 12);
            f.BookingSegmentId = System.Guid.NewGuid();
            f.AgencyCode = "B2C";
            f.PassengerSegmentServiceId = System.Guid.NewGuid();
            f.OriginRcd = "origin_rcd1";
            f.DestinationRcd = "destination_rcd1";
            f.OdOriginRcd = "od_origin_rcd1";
            f.OdDestinationRcd = "od_destination_rcd1";
            f.NumberOfUnits = 0;
            f.TotalAmount = 1;
            f.TotalAmountIncl = 1;
            f.ManualFeeFlag = 0;
            f.OdFlag = 0;
            f.SkipFareAllowanceFlag = 0;
            f.MinimumFeeAmountFlag = 0;
            f.FeePercentage = 1;
            f.TotalFeeAmount = 1;
            f.TotalFeeAmountIncl = 1;
            f.ChargeCurrencyRcd = "THB";
            f.ChargeAmount = 1;
            f.ChargeAmountIncl = 1;
            f.DocumentDateTime = new DateTime(1900, 12, 12);
            f.BaggageFeeOptionId = System.Guid.NewGuid();

            IList<Fee> fee = new List<Fee>();
            fee.Add(f);
            req.Booking.Fees = fee;

            #endregion

            #region Remark
            Remark r = new Remark();
            r.BookingRemarkId = System.Guid.NewGuid();
            r.RemarkTypeRcd = "TKTL";
            r.TimelimitDateTime = new DateTime(1900, 12, 12);
            r.CompleteFlag = 0;
            r.RemarkText = "remark_text1";
            r.BookingId = new Guid(BookingId);
            r.AddedBy = "addby";
            r.ClientProfileId = System.Guid.NewGuid();
            r.AgencyCode = "B2C";
            r.CreateBy = System.Guid.NewGuid();
            r.CreateDateTime = new DateTime(1900, 12, 12);
            r.UpdateBy = System.Guid.NewGuid();
            r.UpdateDateTime = new DateTime(1900, 12, 12);
            r.SystemFlag = 0;
            r.UtcTimelimitDateTime = new DateTime(1900, 12, 12);
            r.ProtectedFlag = 0;
            r.WarningFlag = 0;
            r.ProcessMessageFlag = 0;

            IList<Remark> remark = new List<Remark>();
            remark.Add(r);
            req.Booking.Remarks = remark;

            #endregion

            #region Tax
            Tax t = new Tax();
            t.AcctAmountIncl = 1;
            t.SalesAmount = 1;
            t.SalesAmountIncl = 1;
            t.TaxRcd = "UB";
            t.PassengerId = new Guid(passId);
            t.TaxId = new Guid("3C3B5805-DE7C-11DE-9229-00145E3CC069");
            t.BookingSegmentId = new Guid(BookingSegmentId);
            t.TaxCurrencyRcd = "THB";
            t.SalesCurrencyRcd = "THB";
            t.DisplayName = "display_name1";
            t.SummarizeUp = "summarize_up1";
            t.CoverageType = "coverage_type1";
            t.CreateBy = System.Guid.NewGuid();
            t.UpdateBy = System.Guid.NewGuid();
            t.CreateDateTime = new DateTime(1900, 12, 12);
            t.UpdateDateTime = new DateTime(1900, 12, 12);
            t.VatPercentage = 1;
            IList<Tax> tax = new List<Tax>();
            tax.Add(t);
            req.Booking.Taxs = tax;

            #endregion

            #region Payment
            Message.Booking.Payment pay = new Message.Booking.Payment();

            pay.BookingPaymentId = System.Guid.NewGuid();
            pay.BookingSegmentId = System.Guid.NewGuid();
            pay.BookingId = new Guid(BookingId);
            pay.VoucherPaymentId = System.Guid.NewGuid();
            pay.FormOfPaymentRcd = "VOUCHER";
            pay.CurrencyRcd = "EUR";
            pay.ReceiveCurrencyRcd = "EUR";
            pay.AgencyCode = "B2C";
            pay.DebitAgencyCode = "B2C";
            pay.PaymentAmount = 30;
            pay.ReceivePaymentAmount = 1;
            pay.AcctPaymentAmount = 1;
            pay.PaymentBy = System.Guid.NewGuid();
            pay.PaymentDateTime = new DateTime(1900, 12, 12);
            pay.PaymentDueDateTime = new DateTime(1900, 12, 12);
            pay.DocumentAmount = 1000;
            pay.VoidBy = System.Guid.NewGuid();
            pay.ExpiryMonth = 0;
            pay.ExpiryYear = 0;
            pay.VoidDateTime = new DateTime(1900, 12, 12);
            pay.RecordLocator = "record_locator1";
            pay.NameOnCard = "name_on_card1";
            pay.DocumentNumber = "3700002738";
            pay.FormOfPaymentSubtypeRcd = "form_of_payment_subtype_rcd1";
            pay.City = "city1";
            pay.State = "state1";
            pay.Street = "street1";
            pay.AddressLine1 = "address_line11";
            pay.ZipCode = "zip_code1";
            pay.CountryRcd = "country_rcd1";
            pay.CreateBy = System.Guid.NewGuid();
            pay.CreateDateTime = new DateTime(1900, 12, 12);
            pay.UpdateBy = System.Guid.NewGuid();
            pay.UpdateDateTime = new DateTime(1900, 12, 12);
            pay.PosIndicator = "pos_indicator1";
            pay.IssueMonth = 0;
            pay.IssueYear = 0;
            pay.IssueNumber = "issue_number1";
            pay.ClientProfileId = System.Guid.NewGuid();
            pay.TransactionReference = "transaction_reference1";
            pay.ApprovalCode = "0";
            pay.BankName = "bank_name1";
            pay.BankCode = "bank_code1";
            pay.BankIban = "bank_iban1";
            pay.IpAddress = "ip_address1";
            pay.PaymentReference = "0";
            pay.AllocatedAmount = 1;
            pay.PaymentTypeRcd = "payment_type_rcd1";
            pay.RefundedAmount = 1;

            IList<Message.Booking.Payment> pp = new List<Message.Booking.Payment>();
            pp.Add(pay);
            req.Booking.Payments = pp;

            #endregion

            #region Service
            PassengerService s = new PassengerService();
            s.PassengerSegmentServiceId = System.Guid.NewGuid();
            s.PassengerId = new Guid(passId);
            s.BookingSegmentId = new Guid(BookingSegmentId);
            s.SpecialServiceStatusRcd = "special_service_status_rcd1";
            s.SpecialServiceChangeStatusRcd = "special_service_change_status_rcd1";
            s.SpecialServiceRcd = "special_service_rcd1";
            s.ServiceText = "service_text1";
            s.CreateBy = System.Guid.NewGuid();
            s.CreateDateTime = new DateTime(1900, 12, 12);
            s.UpdateBy = System.Guid.NewGuid();
            s.UpdateDateTime = new DateTime(1900, 12, 12);
            s.NumberOfUnits = 0;
            s.OriginRcd = "origin_rcd1";
            s.DestinationRcd = "destination_rcd1";
            s.DisplayName = "display_name1";
            s.CutOffTime = 0;
            s.ServiceOnRequestFlag = 0;
            s.TextAllowedFlag = 0;
            s.ManifestFlag = 0;
            s.TextRequiredFlag = 0;
            s.IncludePassengerNameFlag = 0;
            s.IncludeFlightSegmentFlag = 0;
            s.IncludeActionCodeFlag = 0;
            s.IncludeNumberOfServiceFlag = 0;
            s.IncludeCateringFlag = 0;
            s.IncludePassengerAssistanceFlag = 0;
            s.ServiceSupportedFlag = 0;
            s.SendInterlineReplyFlag = 0;
            s.InventoryControlFlag = 0;

            IList<PassengerService> ser = new List<PassengerService>();
            ser.Add(s);
            req.Booking.Services = ser;

            #endregion

            #region Mapping
            Mapping m = new Mapping();
            m.OriginRcd = "origin_rcd1";
            m.DestinationRcd = "destination_rcd1";
            m.CreateBy = System.Guid.NewGuid();
            m.UpdateBy = System.Guid.NewGuid();
            m.CreateDateTime = new DateTime(1900, 12, 12);
            m.UpdateDateTime = new DateTime(1900, 12, 12);
            m.CurrencyRcd = "THB";
            m.RoutesTot = 0;
            m.RoutesAvl = 0;
            m.RoutesB2c = 0;
            m.RoutesB2b = 0;
            m.RoutesB2s = 0;
            m.RoutesApi = 0;
            m.RoutesB2t = 0;
            m.SegmentChangeFeeFlag = true;
            m.TransitFlag = true;
            m.DirectFlag = true;
            m.AvlFlag = true;
            m.B2cFlag = true;
            m.B2bFlag = true;
            m.B2tFlag = true;
            m.DayRange = 0;
            m.ShowRedressNumberFlag = true;
            m.RequirePassengerTitleFlag = true;
            m.RequirePassengerGenderFlag = true;
            m.RequireDateOfBirthFlag = true;
            m.RequireDocumentDetailsFlag = true;
            m.RequirePassengerWeightFlag = true;
            m.SpecialServiceFeeFlag = true;
            m.ShowInsuranceOnWebFlag = true;
            m.FlightId = System.Guid.NewGuid();
            m.ExchangedSegmentId = System.Guid.NewGuid();
            m.AirlineRcd = "GR";
            m.FlightNumber = "1234";
            m.RefundableFlag = 0;
            m.GroupFlag = 0;
            m.NonRevenueFlag = 0;
            m.EticketFlag = 0;
            m.FareReduction = 1;
            m.DepartureDate = new DateTime(1900, 12, 12);
            m.ArrivalDate = new DateTime(1900, 12, 12);
            m.OdOriginRcd = "od_origin_rcd1";
            m.OdDestinationRcd = "od_destination_rcd1";
            m.WaitlistFlag = 0;
            m.OverbookFlag = 0;
            m.FlightConnectionId = System.Guid.NewGuid();
            m.AdvancedPurchaseFlag = 0;
            m.JourneyTime = 0;
            m.DepartureTime = 0;
            m.ArrivalTime = 0;
            m.TransitFlightId = System.Guid.NewGuid();
            m.TransitFareId = System.Guid.NewGuid();
            m.TransitDepartureDate = new DateTime(1900, 12, 12);
            m.TransitArrivalDate = new DateTime(1900, 12, 12);
            m.FareId = System.Guid.NewGuid();
            m.BookingClassRcd = "Y";
            m.BoardingClassRcd = "Y";
            m.PlannedDepartureTime = 0;
            m.PlannedArrivalTime = 0;
            m.DepartureDayofweek = 0;
            m.ArrivalDayofweek = 0;
            m.DepartureMonth = 0;
            m.BookingSegmentId = new System.Guid(BookingSegmentId);
            m.BookingId = new System.Guid(BookingId);
            m.NumberOfUnits = 0;
            m.InfoSegmentFlag = 0;
            m.HighPriorityWaitlistFlag = 0;
            m.SeatmapFlag = 0;
            m.TempSeatmapFlag = 0;
            m.AllowWebCheckinFlag = 0;
            m.CloseWebSalesFlag = 0;
            m.ExcludeQuoteFlag = 0;
            m.CurrencyRate = 1;
            m.OpenSequence = 0;
            m.NumberOfStops = 0;
            m.PassengerId = new System.Guid(passId);
            m.Lastname = "lastname1";
            m.Firstname = "firstname1";
            m.GenderTypeRcd = "gender_type_rcd1";
            m.TitleRcd = "title_rcd1";
            m.PassengerTypeRcd = "passenger_type_rcd1";
            m.DateOfBirth = new DateTime(1900, 12, 12);
            m.PassengerStatusRcd = "passenger_status_rcd1";
            m.Middlename = "middlename1";
            m.DocumentTypeRcd = "document_type_rcd1";
            m.PassportNumber = "passport_number1";
            m.PassportIssuePlace = "passport_issue_place1";
            m.PassportIssueDate = new DateTime(1900, 12, 12);
            m.PassportExpiryDate = new DateTime(1900, 12, 12);
            m.Sendmail = "sendmail1";
            m.ClientNumber = 0;
            m.PassengerProfileId = System.Guid.NewGuid();
            m.ClientProfileId = System.Guid.NewGuid();
            m.NationalityRcd = "nationality_rcd1";
            m.ContactEmail = "contact_email1";
            m.InventoryClassRcd = "inventory_class_rcd1";
            m.SeatNumber = "seat_number1";
            m.SeatRow = 0;
            m.SeatColumn = "seat_column1";
            m.NetTotal = 1;
            m.TaxAmount = 1;
            m.FareAmount = 1;
            m.YqAmount = 1;
            m.BaseTicketingFeeAmount = 1;
            m.TicketingFeeAmount = 1;
            m.ReservationFeeAmount = 1;
            m.CommissionAmount = 1;
            m.FareVat = 1;
            m.TaxVat = 1;
            m.YqVat = 1;
            m.TicketingFeeVat = 1;
            m.ReservationFeeVat = 1;
            m.FareAmountIncl = 1;
            m.TaxAmountIncl = 1;
            m.YqAmountIncl = 1;
            m.PublicFareAmountIncl = 1;
            m.PublicFareAmount = 1;
            m.CommissionAmountIncl = 1;
            m.CommissionPercentage = 1;
            m.TicketingFeeAmountIncl = 1;
            m.ReservationFeeAmountIncl = 1;
            m.AcctNetTotal = 1;
            m.AcctTaxAmount = 1;
            m.AcctFareAmount = 1;
            m.AcctYqAmount = 1;
            m.AcctTicketingFeeAmount = 1;
            m.AcctReservationFeeAmount = 1;
            m.AcctCommissionAmount = 1;
            m.AcctFareAmountIncl = 1;
            m.AcctTaxAmountIncl = 1;
            m.AcctYqAmountIncl = 1;
            m.AcctCommissionAmountIncl = 1;
            m.AcctTicketingFeeAmountIncl = 1;
            m.AcctReservationFeeAmountIncl = 1;
            m.FareCode = "fare_code1";
            m.FareDateTime = new DateTime(1900, 12, 12);
            m.ETicketFlag = 0;
            m.StandbyFlag = 0;
            m.TransferableFareFlag = 0;
            m.AgencyCode = "B2C";
            m.TicketNumber = "0";
            m.TicketingDateTime = new DateTime(1900, 12, 12);
            m.TicketingBy = System.Guid.NewGuid();
            m.CheckInSequence = 0;
            m.GroupSequence = 0;
            m.UnloadBy = System.Guid.NewGuid();
            m.UnloadDateTime = new DateTime(1900, 12, 12);
            m.NumberOfPieces = 0;
            m.BaggageWeight = 1;
            m.ExcessBaggageWeight = 1;
            m.CheckInBaggageWeight = 1;
            m.CheckInBy = System.Guid.NewGuid();
            m.CancelBy = System.Guid.NewGuid();
            m.ExchangedDateTime = new DateTime(1900, 12, 12);
            m.CancelDateTime = new DateTime(1900, 12, 12);
            m.RestrictionText = "restriction_text1";
            m.EndorsementText = "endorsement_text1";
            m.ExcludePricingFlag = 0;
            m.CreateName = "create_name1";
            m.UpdateName = "update_name1";
            m.VoidDateTime = new DateTime(1900, 12, 12);
            m.RefundCharge = 1;
            m.RefundableAmount = 1;
            m.RefundDateTime = new DateTime(1900, 12, 12);
            m.PaymentAmount = 1;
            m.MappingSort = 0;
            m.CheckInDateTime = new DateTime(1900, 12, 12);
            m.OnwardDepartureDate = new DateTime(1900, 12, 12);
            m.ETicketStatus = "e_ticket_status1";
            m.HandLuggageFlag = 0;
            m.HandNumberOfPieces = 0;
            m.HandBaggageWeight = 1;
            m.FreeSeatingFlag = 0;
            m.FareTypeRcd = "fare_type_rcd1";
            m.RedemptionPoints = 0;
            m.SeatFeeRcd = "seat_fee_rcd1";
            m.RefundWithChargeHours = 0;
            m.RefundNotPossibleHours = 0;
            m.DutyTravelFlag = 0;
            m.NotValidAfterDate = new DateTime(1900, 12, 12);
            m.NotValidBeforeDate = new DateTime(1900, 12, 12);
            m.AdvancedSeatingFlag = 0;
            m.FareColumn = 0;
            m.PieceAllowance = 1;
            m.BoardingTime = 0;
            m.ItFareFlag = 0;

            IList<Mapping> map = new List<Mapping>();
            map.Add(m);
            req.Booking.Mappings = map;

            #endregion

            response = objBooking.BookingPaymentVoucher(req);


            //ValidateVoucherRequest req = new ValidateVoucherRequest();
            //ValidateVoucherResponse response = new ValidateVoucherResponse();
            //List<Voucher> vouchers = new List<Voucher>();
            //Voucher v1 = new Voucher();
            //v1.VoucherNumber = "3700001234";
            
            //Voucher v2 = new Voucher();
            //v2.VoucherNumber = "3700001230";
            //vouchers.Add(v1);
            //vouchers.Add(v2);

            //req.Vouchers = vouchers;

            //response = objBooking.ValidateVoucherDuplicate(req);
            if (response != null)
            {
                HttpContext.Current.Response.Write(response.Success + " : " + response.Message);
            }
        }

        private void ValidateVoucherPaymentAmount()
        {
            //BookingServiceProxy objBooking = new BookingServiceProxy();
            //ValidateVoucherRequest req = new ValidateVoucherRequest();
            //ValidateVoucherResponse response = new ValidateVoucherResponse();
            //List<Voucher> vouchers = new List<Voucher>();
            //Voucher v1 = new Voucher();
            //v1.VoucherNumber = "3700001234";
            //v1.VoucherValue = 50;

            //Voucher v2 = new Voucher();
            //v2.VoucherNumber = "3700001235";
            //v2.VoucherValue = 60;

            //vouchers.Add(v1);
            //vouchers.Add(v2);

            //Message.Booking.Payment p1 = new Message.Booking.Payment();
            //p1.PaymentAmount = 100;

            //req.Vouchers = vouchers;
            //req.Payment = p1;

            ////response = objBooking.ValidateVoucherPaymentAmount(req);
            //if (response != null)
            //{
            //    HttpContext.Current.Response.Write(response.Success + " : " + response.Message);
            //}
        }

        private void PaymentCreditAgency()
        {
            BookingServiceProxy objBooking = new BookingServiceProxy();
            BookingPaymentCreditAgencyRequest req = new BookingPaymentCreditAgencyRequest();
            BookingPaymentCreditAgencyResponse response = new BookingPaymentCreditAgencyResponse();

            string BookingId = System.Guid.NewGuid().ToString();
            string BookingSegmentId = System.Guid.NewGuid().ToString();
            string passId = System.Guid.NewGuid().ToString();

            #region Header
            if (req.Booking == null) req.Booking = new BookingMessage();
            if (req.Booking.Header == null) req.Booking.Header = new BookingHeader();

            req.Booking.Header.BookingId = new Guid(BookingId);
            req.Booking.Header.BookingSourceRcd = "bookingsourcercd1";
            req.Booking.Header.CurrencyRcd = "THB";
            req.Booking.Header.ClientProfileId = System.Guid.NewGuid();
            req.Booking.Header.BookingNumber = 0;
            req.Booking.Header.RecordLocator = "";
            req.Booking.Header.NumberOfAdults = 0;
            req.Booking.Header.NumberOfChildren = 0;
            req.Booking.Header.NumberOfInfants = 0;
            req.Booking.Header.LanguageRcd = "languagercd1";
            req.Booking.Header.AgencyCode = "B2C";
            req.Booking.Header.ContactName = "contactname1";
            req.Booking.Header.ContactEmail = "contactemail1";
            req.Booking.Header.PhoneMobile = "phonemobile1";
            req.Booking.Header.PhoneHome = "phonehome1";
            req.Booking.Header.PhoneBusiness = "phonebusiness1";
            req.Booking.Header.ReceivedFrom = "receivedfrom1";
            req.Booking.Header.PhoneFax = "phonefax1";
            req.Booking.Header.PhoneSearch = "phonesearch1";
            req.Booking.Header.Comment = "comment";

            req.Booking.Header.NotifyByEmailFlag = 0;
            req.Booking.Header.NotifyBySmsFlag = 0;
            req.Booking.Header.GroupName = "groupname1";
            req.Booking.Header.GroupBookingFlag = 0;
            req.Booking.Header.AgencyName = "agencyname1";
            req.Booking.Header.OwnAgencyFlag = 0;
            req.Booking.Header.WebAgencyFlag = 0;
            req.Booking.Header.ClientNumber = 0;

            req.Booking.Header.Lastname = "aaa";
            req.Booking.Header.Firstname = "bbb";
            req.Booking.Header.City = "bkk";

            req.Booking.Header.CreateName = "createname1";
            req.Booking.Header.Street = "bkk";
            req.Booking.Header.LockDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.Middlename = "midd";
            req.Booking.Header.AddressLine1 = "addressline11";
            req.Booking.Header.AddressLine2 = "addressline21";
            req.Booking.Header.State = "state";
            req.Booking.Header.District = "bkk";
            req.Booking.Header.Province = "province";
            req.Booking.Header.ZipCode = "zip_code1";
            req.Booking.Header.PoBox = "po_box1";
            req.Booking.Header.CountryRcd = "countryrcd1";
            req.Booking.Header.TitleRcd = "titlercd1";
            req.Booking.Header.ExternalPaymentReference = "externalpaymentreference1";
            req.Booking.Header.CreateBy = System.Guid.NewGuid();
            req.Booking.Header.UpdateBy = System.Guid.NewGuid();
            req.Booking.Header.CreateDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.UpdateDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.CostCenter = "cost_center1";
            req.Booking.Header.PurchaseOrder = "purchaseorder1";
            req.Booking.Header.ProjectNumber = "projectnumber1";
            req.Booking.Header.LockName = "lockname1";
            req.Booking.Header.IpAddress = "ipaddress1";
            req.Booking.Header.ApprovalFlag = 0;
            req.Booking.Header.InvoiceReceiver = "invoicereceiver1";
            req.Booking.Header.TaxId = "taxid1";
            req.Booking.Header.NewsletterFlag = 0;
            req.Booking.Header.ContactEmailCc = "contactemailcc1";
            req.Booking.Header.MobileEmail = "mobileemail1";
            req.Booking.Header.VendorRcd = "vendorrcd1";
            req.Booking.Header.BookingDateTime = new DateTime(1900, 12, 12);
            req.Booking.Header.NoVatFlag = 0;
            req.Booking.Header.CompanyName = "companyname1";
            req.Booking.Header.BusinessFlag = 0;

            #endregion

            #region FlightSegment
            FlightSegment Segment = new FlightSegment();

            Segment.OriginRcd = "ZRH";
            Segment.DestinationRcd = "AMS";
            Segment.CreateBy = System.Guid.NewGuid();
            Segment.UpdateBy = System.Guid.NewGuid();
            Segment.CreateDateTime = new DateTime(1900, 12, 12);
            Segment.UpdateDateTime = new DateTime(1900, 12, 12);
            Segment.RoutesTot = 0;
            Segment.RoutesAvl = 0;
            Segment.RoutesB2c = 0;
            Segment.RoutesB2b = 0;
            Segment.RoutesB2s = 0;
            Segment.RoutesApi = 0;
            Segment.RoutesB2t = 0;
            Segment.SegmentChangeFeeFlag = true;
            Segment.TransitFlag = true;
            Segment.DirectFlag = true;
            Segment.AvlFlag = true;
            Segment.B2cFlag = true;
            Segment.B2bFlag = true;
            Segment.B2tFlag = true;
            Segment.DayRange = 0;
            Segment.ShowRedressNumberFlag = true;
            Segment.RequirePassengerTitleFlag = true;
            Segment.RequirePassengerGenderFlag = true;
            Segment.RequireDateOfBirthFlag = true;
            Segment.RequireDocumentDetailsFlag = true;
            Segment.RequirePassengerWeightFlag = true;
            Segment.SpecialServiceFeeFlag = true;
            Segment.ShowInsuranceOnWebFlag = true;
            Segment.FlightId = System.Guid.NewGuid();
            Segment.ExchangedSegmentId = System.Guid.NewGuid();
            Segment.AirlineRcd = "GR";
            Segment.FlightNumber = "1234";
            Segment.RefundableFlag = 0;
            Segment.GroupFlag = 0;
            Segment.NonRevenueFlag = 0;
            Segment.EticketFlag = 0;
            Segment.FareReduction = 1;
            Segment.DepartureDate = new DateTime(1900, 12, 12);
            Segment.ArrivalDate = new DateTime(1900, 12, 12);
            Segment.OdOriginRcd = "od_origin_rcd1";
            Segment.OdDestinationRcd = "od_destination_rcd1";
            Segment.WaitlistFlag = 0;
            Segment.OverbookFlag = 0;
            Segment.FlightConnectionId = System.Guid.NewGuid();
            Segment.AdvancedPurchaseFlag = 0;
            Segment.JourneyTime = 1;
            Segment.DepartureTime = 0;
            Segment.ArrivalTime = 0;
            Segment.OriginName = "origin_name";
            Segment.DestinationName = "destination_name";
            Segment.TransitFlightId = System.Guid.NewGuid();
            Segment.TransitFareId = System.Guid.NewGuid();
            Segment.TransitDepartureDate = new DateTime(1900, 12, 12);
            Segment.TransitArrivalDate = new DateTime(1900, 12, 12);
            Segment.FareId = System.Guid.NewGuid();
            Segment.BookingClassRcd = "Y";
            Segment.BoardingClassRcd = "Y";
            Segment.AircraftTypeRcd = "aircraft_type_rcd1";
            Segment.PlannedDepartureTime = 0;
            Segment.PlannedArrivalTime = 0;
            Segment.DepartureDayofweek = 0;
            Segment.ArrivalDayofweek = 0;
            Segment.DepartureMonth = 0;
            Segment.BookingSegmentId = new Guid(BookingSegmentId);
            Segment.SegmentStatusRcd = "segment_status_rcd1";
            Segment.BookingId = new Guid(BookingId);
            Segment.NumberOfUnits = 0;
            Segment.SegmentChangeStatusRcd = "segment_change_status_rcd1";
            Segment.InfoSegmentFlag = 0;
            Segment.HighPriorityWaitlistFlag = 0;
            Segment.SegmentStatusName = "segment_status_name1";
            Segment.SeatmapFlag = 0;
            Segment.TempSeatmapFlag = 0;
            Segment.AllowWebCheckinFlag = 0;
            Segment.CloseWebSalesFlag = 0;
            Segment.ExcludeQuoteFlag = 0;
            Segment.CurrencyRate = 1;
            Segment.OpenSequence = 0;
            Segment.NumberOfStops = 0;

            IList<FlightSegment> segmentList = new List<FlightSegment>();
            segmentList.Add(Segment);
            req.Booking.FlightSegments = segmentList;

            #endregion

            #region Passenger

            Passenger p = new Passenger();
            p.TitleRcd = "title_rcd1";
            p.Lastname = "lastname1";
            p.Firstname = "firstname1";
            p.Middlename = "middlename1";
            p.NationalityRcd = "nationality_rcd1";
            p.PassengerWeight = 1;
            p.GenderTypeRcd = "gender_type_rcd1";
            p.PassengerTypeRcd = "passenger_type_rcd1";
            p.ClientProfileId = System.Guid.NewGuid();
            p.ClientNumber = 0;
            p.AddressLine1 = "address_line11";
            p.AddressLine2 = "address_line21";
            p.State = "state1";
            p.District = "district1";
            p.Province = "province1";
            p.ZipCode = "zip_code1";
            p.PoBox = "po_box1";
            p.CountryRcd = "country_rcd1";
            p.Street = "street1";
            p.City = "city1";
            p.DocumentTypeRcd = "document_type_rcd1";
            p.DocumentNumber = "document_number1";
            p.ResidenceCountryRcd = "residence_country_rcd1";
            p.PassportNumber = "passport_number1";
            p.PassportIssueDate = new DateTime(1900, 12, 12);
            p.PassportExpiryDate = new DateTime(1900, 12, 12);
            p.PassportIssuePlace = "passport_issue_place1";
            p.PassportBirthPlace = "passport_birth_place1";
            p.DateOfBirth = new DateTime(1900, 12, 12);
            p.PassportIssueCountryRcd = "passport_issue_country_rcd1";
            p.ContactName = "contact_name1";
            p.ContactEmail = "contact_email1";
            p.MobileEmail = "mobile_email1";
            p.PhoneMobile = "phone_mobile1";
            p.PhoneHome = "phone_home1";
            p.PhoneFax = "phone_fax1";
            p.PhoneBusiness = "phone_business1";
            p.ReceivedFrom = "received_from1";
            p.CreateBy = System.Guid.NewGuid();
            p.CreateDateTime = new DateTime(1900, 12, 12);
            p.UpdateBy = System.Guid.NewGuid();
            p.UpdateDateTime = new DateTime(1900, 12, 12);
            p.PassengerId = new Guid(passId);
            p.BookingId = new Guid(BookingId);
            p.PassengerProfileId = System.Guid.NewGuid();
            p.EmployeeNumber = "employee_number1";
            p.WheelchairFlag = 0;
            p.VipFlag = 0;
            p.MemberLevelRcd = "member_level_rcd1";
            p.MemberNumber = "member_number1";
            p.WindowSeatFlag = 0;
            p.RedressNumber = "redress_number1";
            p.PnrName = "pnr_name1";

            IList<Passenger> pass = new List<Passenger>();
            pass.Add(p);
            req.Booking.Passengers = pass;

            #endregion

            #region Fee
            Fee f = new Fee();
            f.BookingFeeId = System.Guid.NewGuid();
            f.FeeAmount = 1;
            f.BookingId = new Guid(BookingId);
            f.PassengerId = System.Guid.NewGuid();
            f.CurrencyRcd = "THB";
            f.AcctFeeAmount = 1;
            f.FeeId = System.Guid.NewGuid();
            f.VatPercentage = 1;
            f.FeeAmountIncl = 1;
            f.AcctFeeAmountIncl = 1;
            f.FeeRcd = "fee_rcd1";
            f.DisplayName = "display_name1";
            f.AccountFeeBy = System.Guid.NewGuid();
            f.AccountFeeDateTime = new DateTime(1900, 12, 12);
            f.VoidDateTime = new DateTime(1900, 12, 12);
            f.VoidBy = System.Guid.NewGuid();
            f.PaymentAmount = 1;
            f.CreateBy = System.Guid.NewGuid();
            f.CreateDateTime = new DateTime(1900, 12, 12);
            f.UpdateBy = System.Guid.NewGuid();
            f.UpdateDateTime = new DateTime(1900, 12, 12);
            f.BookingSegmentId = System.Guid.NewGuid();
            f.AgencyCode = "B2C";
            f.PassengerSegmentServiceId = System.Guid.NewGuid();
            f.OriginRcd = "origin_rcd1";
            f.DestinationRcd = "destination_rcd1";
            f.OdOriginRcd = "od_origin_rcd1";
            f.OdDestinationRcd = "od_destination_rcd1";
            f.NumberOfUnits = 0;
            f.TotalAmount = 1;
            f.TotalAmountIncl = 1;
            f.ManualFeeFlag = 0;
            f.OdFlag = 0;
            f.SkipFareAllowanceFlag = 0;
            f.MinimumFeeAmountFlag = 0;
            f.FeePercentage = 1;
            f.TotalFeeAmount = 1;
            f.TotalFeeAmountIncl = 1;
            f.ChargeCurrencyRcd = "THB";
            f.ChargeAmount = 1;
            f.ChargeAmountIncl = 1;
            f.DocumentDateTime = new DateTime(1900, 12, 12);
            f.BaggageFeeOptionId = System.Guid.NewGuid();

            IList<Fee> fee = new List<Fee>();
            fee.Add(f);
            req.Booking.Fees = fee;

            #endregion

            #region Remark
            Remark r = new Remark();
            r.BookingRemarkId = System.Guid.NewGuid();
            r.RemarkTypeRcd = "TKTL";
            r.TimelimitDateTime = new DateTime(1900, 12, 12);
            r.CompleteFlag = 0;
            r.RemarkText = "remark_text1";
            r.BookingId = new Guid(BookingId);
            r.AddedBy = "addby";
            r.ClientProfileId = System.Guid.NewGuid();
            r.AgencyCode = "B2C";
            r.CreateBy = System.Guid.NewGuid();
            r.CreateDateTime = new DateTime(1900, 12, 12);
            r.UpdateBy = System.Guid.NewGuid();
            r.UpdateDateTime = new DateTime(1900, 12, 12);
            r.SystemFlag = 0;
            r.UtcTimelimitDateTime = new DateTime(1900, 12, 12);
            r.ProtectedFlag = 0;
            r.WarningFlag = 0;
            r.ProcessMessageFlag = 0;

            IList<Remark> remark = new List<Remark>();
            remark.Add(r);
            req.Booking.Remarks = remark;

            #endregion

            #region Tax
            Tax t = new Tax();
            t.AcctAmountIncl = 1;
            t.SalesAmount = 1;
            t.SalesAmountIncl = 1;
            t.TaxRcd = "UB";
            t.PassengerId = new Guid(passId);
            t.TaxId = new Guid("3C3B5805-DE7C-11DE-9229-00145E3CC069");
            t.BookingSegmentId = new Guid(BookingSegmentId);
            t.TaxCurrencyRcd = "THB";
            t.SalesCurrencyRcd = "THB";
            t.DisplayName = "display_name1";
            t.SummarizeUp = "summarize_up1";
            t.CoverageType = "coverage_type1";
            t.CreateBy = System.Guid.NewGuid();
            t.UpdateBy = System.Guid.NewGuid();
            t.CreateDateTime = new DateTime(1900, 12, 12);
            t.UpdateDateTime = new DateTime(1900, 12, 12);
            t.VatPercentage = 1;
            IList<Tax> tax = new List<Tax>();
            tax.Add(t);
            req.Booking.Taxs = tax;

            #endregion

            #region Payment
            Message.Booking.Payment pay = new Message.Booking.Payment();

            pay.BookingPaymentId = System.Guid.NewGuid();
            pay.BookingSegmentId = System.Guid.NewGuid();
            pay.BookingId = new Guid(BookingId);
            pay.VoucherPaymentId = System.Guid.NewGuid();
            pay.FormOfPaymentRcd = "CRAGT";
            pay.CurrencyRcd = "GBP";
            pay.ReceiveCurrencyRcd = "GBP";
            pay.AgencyCode = "B2C";
            pay.DebitAgencyCode = "B2C";
            pay.PaymentAmount = 30;
            pay.ReceivePaymentAmount = 1;
            pay.AcctPaymentAmount = 1;
            pay.PaymentBy = System.Guid.NewGuid();
            pay.PaymentDateTime = new DateTime(1900, 12, 12);
            pay.PaymentDueDateTime = new DateTime(1900, 12, 12);
            pay.DocumentAmount = 1000;
            pay.VoidBy = System.Guid.NewGuid();
            pay.ExpiryMonth = 0;
            pay.ExpiryYear = 0;
            pay.VoidDateTime = new DateTime(1900, 12, 12);
            pay.RecordLocator = "record_locator1";
            pay.NameOnCard = "name_on_card1";
            pay.DocumentNumber = "3700002738";
            pay.FormOfPaymentSubtypeRcd = "form_of_payment_subtype_rcd1";
            pay.City = "city1";
            pay.State = "state1";
            pay.Street = "street1";
            pay.AddressLine1 = "address_line11";
            pay.ZipCode = "zip_code1";
            pay.CountryRcd = "country_rcd1";
            pay.CreateBy = System.Guid.NewGuid();
            pay.CreateDateTime = new DateTime(1900, 12, 12);
            pay.UpdateBy = System.Guid.NewGuid();
            pay.UpdateDateTime = new DateTime(1900, 12, 12);
            pay.PosIndicator = "pos_indicator1";
            pay.IssueMonth = 0;
            pay.IssueYear = 0;
            pay.IssueNumber = "issue_number1";
            pay.ClientProfileId = System.Guid.NewGuid();
            pay.TransactionReference = "transaction_reference1";
            pay.ApprovalCode = "0";
            pay.BankName = "bank_name1";
            pay.BankCode = "bank_code1";
            pay.BankIban = "bank_iban1";
            pay.IpAddress = "ip_address1";
            pay.PaymentReference = "0";
            pay.AllocatedAmount = 1;
            pay.PaymentTypeRcd = "payment_type_rcd1";
            pay.RefundedAmount = 1;

            IList<Message.Booking.Payment> pp = new List<Message.Booking.Payment>();
            pp.Add(pay);
            req.Booking.Payments = pp;

            #endregion

            #region Service
            PassengerService s = new PassengerService();
            s.PassengerSegmentServiceId = System.Guid.NewGuid();
            s.PassengerId = new Guid(passId);
            s.BookingSegmentId = new Guid(BookingSegmentId);
            s.SpecialServiceStatusRcd = "special_service_status_rcd1";
            s.SpecialServiceChangeStatusRcd = "special_service_change_status_rcd1";
            s.SpecialServiceRcd = "special_service_rcd1";
            s.ServiceText = "service_text1";
            s.CreateBy = System.Guid.NewGuid();
            s.CreateDateTime = new DateTime(1900, 12, 12);
            s.UpdateBy = System.Guid.NewGuid();
            s.UpdateDateTime = new DateTime(1900, 12, 12);
            s.NumberOfUnits = 0;
            s.OriginRcd = "origin_rcd1";
            s.DestinationRcd = "destination_rcd1";
            s.DisplayName = "display_name1";
            s.CutOffTime = 0;
            s.ServiceOnRequestFlag = 0;
            s.TextAllowedFlag = 0;
            s.ManifestFlag = 0;
            s.TextRequiredFlag = 0;
            s.IncludePassengerNameFlag = 0;
            s.IncludeFlightSegmentFlag = 0;
            s.IncludeActionCodeFlag = 0;
            s.IncludeNumberOfServiceFlag = 0;
            s.IncludeCateringFlag = 0;
            s.IncludePassengerAssistanceFlag = 0;
            s.ServiceSupportedFlag = 0;
            s.SendInterlineReplyFlag = 0;
            s.InventoryControlFlag = 0;

            IList<PassengerService> ser = new List<PassengerService>();
            ser.Add(s);
            req.Booking.Services = ser;

            #endregion

            #region Mapping
            Mapping m = new Mapping();
            m.OriginRcd = "origin_rcd1";
            m.DestinationRcd = "destination_rcd1";
            m.CreateBy = System.Guid.NewGuid();
            m.UpdateBy = System.Guid.NewGuid();
            m.CreateDateTime = new DateTime(1900, 12, 12);
            m.UpdateDateTime = new DateTime(1900, 12, 12);
            m.CurrencyRcd = "THB";
            m.RoutesTot = 0;
            m.RoutesAvl = 0;
            m.RoutesB2c = 0;
            m.RoutesB2b = 0;
            m.RoutesB2s = 0;
            m.RoutesApi = 0;
            m.RoutesB2t = 0;
            m.SegmentChangeFeeFlag = true;
            m.TransitFlag = true;
            m.DirectFlag = true;
            m.AvlFlag = true;
            m.B2cFlag = true;
            m.B2bFlag = true;
            m.B2tFlag = true;
            m.DayRange = 0;
            m.ShowRedressNumberFlag = true;
            m.RequirePassengerTitleFlag = true;
            m.RequirePassengerGenderFlag = true;
            m.RequireDateOfBirthFlag = true;
            m.RequireDocumentDetailsFlag = true;
            m.RequirePassengerWeightFlag = true;
            m.SpecialServiceFeeFlag = true;
            m.ShowInsuranceOnWebFlag = true;
            m.FlightId = System.Guid.NewGuid();
            m.ExchangedSegmentId = System.Guid.NewGuid();
            m.AirlineRcd = "GR";
            m.FlightNumber = "1234";
            m.RefundableFlag = 0;
            m.GroupFlag = 0;
            m.NonRevenueFlag = 0;
            m.EticketFlag = 0;
            m.FareReduction = 1;
            m.DepartureDate = new DateTime(1900, 12, 12);
            m.ArrivalDate = new DateTime(1900, 12, 12);
            m.OdOriginRcd = "od_origin_rcd1";
            m.OdDestinationRcd = "od_destination_rcd1";
            m.WaitlistFlag = 0;
            m.OverbookFlag = 0;
            m.FlightConnectionId = System.Guid.NewGuid();
            m.AdvancedPurchaseFlag = 0;
            m.JourneyTime = 0;
            m.DepartureTime = 0;
            m.ArrivalTime = 0;
            m.TransitFlightId = System.Guid.NewGuid();
            m.TransitFareId = System.Guid.NewGuid();
            m.TransitDepartureDate = new DateTime(1900, 12, 12);
            m.TransitArrivalDate = new DateTime(1900, 12, 12);
            m.FareId = System.Guid.NewGuid();
            m.BookingClassRcd = "Y";
            m.BoardingClassRcd = "Y";
            m.PlannedDepartureTime = 0;
            m.PlannedArrivalTime = 0;
            m.DepartureDayofweek = 0;
            m.ArrivalDayofweek = 0;
            m.DepartureMonth = 0;
            m.BookingSegmentId = new System.Guid(BookingSegmentId);
            m.BookingId = new System.Guid(BookingId);
            m.NumberOfUnits = 0;
            m.InfoSegmentFlag = 0;
            m.HighPriorityWaitlistFlag = 0;
            m.SeatmapFlag = 0;
            m.TempSeatmapFlag = 0;
            m.AllowWebCheckinFlag = 0;
            m.CloseWebSalesFlag = 0;
            m.ExcludeQuoteFlag = 0;
            m.CurrencyRate = 1;
            m.OpenSequence = 0;
            m.NumberOfStops = 0;
            m.PassengerId = new System.Guid(passId);
            m.Lastname = "lastname1";
            m.Firstname = "firstname1";
            m.GenderTypeRcd = "gender_type_rcd1";
            m.TitleRcd = "title_rcd1";
            m.PassengerTypeRcd = "passenger_type_rcd1";
            m.DateOfBirth = new DateTime(1900, 12, 12);
            m.PassengerStatusRcd = "passenger_status_rcd1";
            m.Middlename = "middlename1";
            m.DocumentTypeRcd = "document_type_rcd1";
            m.PassportNumber = "passport_number1";
            m.PassportIssuePlace = "passport_issue_place1";
            m.PassportIssueDate = new DateTime(1900, 12, 12);
            m.PassportExpiryDate = new DateTime(1900, 12, 12);
            m.Sendmail = "sendmail1";
            m.ClientNumber = 0;
            m.PassengerProfileId = System.Guid.NewGuid();
            m.ClientProfileId = System.Guid.NewGuid();
            m.NationalityRcd = "nationality_rcd1";
            m.ContactEmail = "contact_email1";
            m.InventoryClassRcd = "inventory_class_rcd1";
            m.SeatNumber = "seat_number1";
            m.SeatRow = 0;
            m.SeatColumn = "seat_column1";
            m.NetTotal = 1;
            m.TaxAmount = 1;
            m.FareAmount = 1;
            m.YqAmount = 1;
            m.BaseTicketingFeeAmount = 1;
            m.TicketingFeeAmount = 1;
            m.ReservationFeeAmount = 1;
            m.CommissionAmount = 1;
            m.FareVat = 1;
            m.TaxVat = 1;
            m.YqVat = 1;
            m.TicketingFeeVat = 1;
            m.ReservationFeeVat = 1;
            m.FareAmountIncl = 1;
            m.TaxAmountIncl = 1;
            m.YqAmountIncl = 1;
            m.PublicFareAmountIncl = 1;
            m.PublicFareAmount = 1;
            m.CommissionAmountIncl = 1;
            m.CommissionPercentage = 1;
            m.TicketingFeeAmountIncl = 1;
            m.ReservationFeeAmountIncl = 1;
            m.AcctNetTotal = 1;
            m.AcctTaxAmount = 1;
            m.AcctFareAmount = 1;
            m.AcctYqAmount = 1;
            m.AcctTicketingFeeAmount = 1;
            m.AcctReservationFeeAmount = 1;
            m.AcctCommissionAmount = 1;
            m.AcctFareAmountIncl = 1;
            m.AcctTaxAmountIncl = 1;
            m.AcctYqAmountIncl = 1;
            m.AcctCommissionAmountIncl = 1;
            m.AcctTicketingFeeAmountIncl = 1;
            m.AcctReservationFeeAmountIncl = 1;
            m.FareCode = "fare_code1";
            m.FareDateTime = new DateTime(1900, 12, 12);
            m.ETicketFlag = 0;
            m.StandbyFlag = 0;
            m.TransferableFareFlag = 0;
            m.AgencyCode = "B2C";
            m.TicketNumber = "0";
            m.TicketingDateTime = new DateTime(1900, 12, 12);
            m.TicketingBy = System.Guid.NewGuid();
            m.CheckInSequence = 0;
            m.GroupSequence = 0;
            m.UnloadBy = System.Guid.NewGuid();
            m.UnloadDateTime = new DateTime(1900, 12, 12);
            m.NumberOfPieces = 0;
            m.BaggageWeight = 1;
            m.ExcessBaggageWeight = 1;
            m.CheckInBaggageWeight = 1;
            m.CheckInBy = System.Guid.NewGuid();
            m.CancelBy = System.Guid.NewGuid();
            m.ExchangedDateTime = new DateTime(1900, 12, 12);
            m.CancelDateTime = new DateTime(1900, 12, 12);
            m.RestrictionText = "restriction_text1";
            m.EndorsementText = "endorsement_text1";
            m.ExcludePricingFlag = 0;
            m.CreateName = "create_name1";
            m.UpdateName = "update_name1";
            m.VoidDateTime = new DateTime(1900, 12, 12);
            m.RefundCharge = 1;
            m.RefundableAmount = 1;
            m.RefundDateTime = new DateTime(1900, 12, 12);
            m.PaymentAmount = 1;
            m.MappingSort = 0;
            m.CheckInDateTime = new DateTime(1900, 12, 12);
            m.OnwardDepartureDate = new DateTime(1900, 12, 12);
            m.ETicketStatus = "e_ticket_status1";
            m.HandLuggageFlag = 0;
            m.HandNumberOfPieces = 0;
            m.HandBaggageWeight = 1;
            m.FreeSeatingFlag = 0;
            m.FareTypeRcd = "fare_type_rcd1";
            m.RedemptionPoints = 0;
            m.SeatFeeRcd = "seat_fee_rcd1";
            m.RefundWithChargeHours = 0;
            m.RefundNotPossibleHours = 0;
            m.DutyTravelFlag = 0;
            m.NotValidAfterDate = new DateTime(1900, 12, 12);
            m.NotValidBeforeDate = new DateTime(1900, 12, 12);
            m.AdvancedSeatingFlag = 0;
            m.FareColumn = 0;
            m.PieceAllowance = 1;
            m.BoardingTime = 0;
            m.ItFareFlag = 0;

            IList<Mapping> map = new List<Mapping>();
            map.Add(m);
            req.Booking.Mappings = map;

            #endregion

            response = objBooking.BookingPaymentCreditAgency(req);
            if (response != null)
            {
                HttpContext.Current.Response.Write(response.Success + " : " + response.Message);
            }
        }
    }
}