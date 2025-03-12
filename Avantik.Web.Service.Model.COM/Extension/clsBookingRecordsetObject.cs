using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Exception.Booking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class clsBookingRecordsetObject
    {
        public static BookingHeader FillBooking(this BookingHeader header, ADODB.Recordset rs)
        {
            header = new BookingHeader();
            try
            {
                header.BookingId = RecordsetHelper.ToGuid(rs, "booking_id");
                header.BookingSourceRcd = RecordsetHelper.ToString(rs, "booking_source_rcd");
                header.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                header.ClientProfileId = RecordsetHelper.ToGuid(rs, "client_profile_id");
                header.BookingNumber = RecordsetHelper.ToInt64(rs, "booking_number");
                header.RecordLocator = RecordsetHelper.ToString(rs, "record_locator");
                header.NumberOfAdults = RecordsetHelper.ToInt16(rs, "number_of_adults");
                header.NumberOfChildren = RecordsetHelper.ToInt32(rs, "number_of_children");
                header.NumberOfInfants = RecordsetHelper.ToInt32(rs, "number_of_infants");
                header.LanguageRcd = RecordsetHelper.ToString(rs, "language_rcd");
                header.AgencyCode = RecordsetHelper.ToString(rs, "agency_code");
                header.ContactName = RecordsetHelper.ToString(rs, "contact_name");
                header.ContactEmail = RecordsetHelper.ToString(rs, "contact_email");
                header.PhoneMobile = RecordsetHelper.ToString(rs, "phone_mobile");
                header.PhoneHome = RecordsetHelper.ToString(rs, "phone_home");
                header.PhoneBusiness = RecordsetHelper.ToString(rs, "phone_business");
                header.ReceivedFrom = RecordsetHelper.ToString(rs, "received_from");
                header.PhoneFax = RecordsetHelper.ToString(rs, "phone_fax");
                header.PhoneSearch = RecordsetHelper.ToString(rs, "phone_search");
                header.Comment = RecordsetHelper.ToString(rs, "comment");
                header.NotifyByEmailFlag = RecordsetHelper.ToByte(rs, "notify_by_email_flag");
                header.NotifyBySmsFlag = RecordsetHelper.ToByte(rs, "notify_by_sms_flag");
                header.GroupName = RecordsetHelper.ToString(rs, "group_name");
                header.GroupBookingFlag = RecordsetHelper.ToByte(rs, "group_booking_flag");
                header.AgencyName = RecordsetHelper.ToString(rs, "agency_name");
                header.OwnAgencyFlag = RecordsetHelper.ToByte(rs, "own_agency_flag");
                header.WebAgencyFlag = RecordsetHelper.ToByte(rs, "web_agency_flag");
                header.ClientNumber = RecordsetHelper.ToInt64(rs, "client_number");
                header.Lastname = RecordsetHelper.ToString(rs, "lastname");
                header.Firstname = RecordsetHelper.ToString(rs, "firstname");
                header.City = RecordsetHelper.ToString(rs, "city");
                header.CreateName = RecordsetHelper.ToString(rs, "create_name");
                header.Street = RecordsetHelper.ToString(rs, "street");
                header.LockDateTime = RecordsetHelper.ToDateTime(rs, "lock_date_time");
                header.Middlename = RecordsetHelper.ToString(rs, "middlename");
                header.AddressLine1 = RecordsetHelper.ToString(rs, "address_line1");
                header.AddressLine2 = RecordsetHelper.ToString(rs, "address_line2");
                header.State = RecordsetHelper.ToString(rs, "state");
                header.District = RecordsetHelper.ToString(rs, "district");
                header.Province = RecordsetHelper.ToString(rs, "province");
                header.ZipCode = RecordsetHelper.ToString(rs, "zip_code");
                header.PoBox = RecordsetHelper.ToString(rs, "po_box");
                header.CountryRcd = RecordsetHelper.ToString(rs, "country_rcd");
                header.TitleRcd = RecordsetHelper.ToString(rs, "title_rcd");
                header.ExternalPaymentReference = RecordsetHelper.ToString(rs, "external_payment_reference");
                header.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                header.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                header.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                header.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                header.CostCenter = RecordsetHelper.ToString(rs, "cost_center");
                header.PurchaseOrder = RecordsetHelper.ToString(rs, "purchase_order");
                header.ProjectNumber = RecordsetHelper.ToString(rs, "project_number");
                header.LockName = RecordsetHelper.ToString(rs, "lock_name");
                header.IpAddress = RecordsetHelper.ToString(rs, "ip_address");
                header.ApprovalFlag = RecordsetHelper.ToByte(rs, "approval_flag");
                header.InvoiceReceiver = RecordsetHelper.ToString(rs, "invoice_receiver");
                header.TaxId = RecordsetHelper.ToString(rs, "tax_id");
                header.NewsletterFlag = RecordsetHelper.ToByte(rs, "newsletter_flag");
                header.ContactEmailCc = RecordsetHelper.ToString(rs, "contact_email_cc");
                header.MobileEmail = RecordsetHelper.ToString(rs, "mobile_email");
                header.VendorRcd = RecordsetHelper.ToString(rs, "vendor_rcd");
                header.BookingDateTime = RecordsetHelper.ToDateTime(rs, "booking_date_time");
                header.NoVatFlag = RecordsetHelper.ToByte(rs, "no_vat_flag");
                header.CompanyName = RecordsetHelper.ToString(rs, "company_name");
                header.BusinessFlag = RecordsetHelper.ToByte(rs, "business_flag");
            }
            catch
            {
                throw;
            }
            return header;
        }

        public static IList<FlightSegment> FillBooking(this IList<FlightSegment> segmentList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                segmentList = new List<FlightSegment>();
                FlightSegment segment = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        segment = new FlightSegment();

                        segment.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        segment.DestinationRcd = RecordsetHelper.ToString(rs, "destination_rcd");
                        segment.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        segment.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        segment.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        segment.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        segment.RoutesTot = RecordsetHelper.ToInt16(rs, "routes_tot");
                        segment.RoutesAvl = RecordsetHelper.ToInt16(rs, "routes_avl");
                        segment.RoutesB2c = RecordsetHelper.ToInt16(rs, "routes_b2c");
                        segment.RoutesB2b = RecordsetHelper.ToInt16(rs, "routes_b2b");
                        segment.RoutesB2s = RecordsetHelper.ToInt16(rs, "routes_b2s");
                        segment.RoutesApi = RecordsetHelper.ToInt16(rs, "routes_api");
                        segment.RoutesB2t = RecordsetHelper.ToInt16(rs, "routes_b2t");
                        segment.SegmentChangeFeeFlag = RecordsetHelper.ToBoolean(rs, "segment_change_fee_flag");
                        segment.TransitFlag = RecordsetHelper.ToBoolean(rs, "transit_flag");
                        segment.DirectFlag = RecordsetHelper.ToBoolean(rs, "direct_flag");
                        segment.AvlFlag = RecordsetHelper.ToBoolean(rs, "avl_flag");
                        segment.B2cFlag = RecordsetHelper.ToBoolean(rs, "b2c_flag");
                        segment.B2bFlag = RecordsetHelper.ToBoolean(rs, "b2b_flag");
                        segment.B2tFlag = RecordsetHelper.ToBoolean(rs, "b2t_flag");
                        segment.DayRange = RecordsetHelper.ToInt16(rs, "day_range");
                        segment.ShowRedressNumberFlag = RecordsetHelper.ToBoolean(rs, "show_redressNumber_flag");
                        segment.RequirePassengerTitleFlag = RecordsetHelper.ToBoolean(rs, "require_passenger_title_flag");
                        segment.RequirePassengerGenderFlag = RecordsetHelper.ToBoolean(rs, "require_passenger_gender_flag");
                        segment.RequireDateOfBirthFlag = RecordsetHelper.ToBoolean(rs, "require_date_of_birth_flag");
                        segment.RequireDocumentDetailsFlag = RecordsetHelper.ToBoolean(rs, "require_document_details_flag");
                        segment.RequirePassengerWeightFlag = RecordsetHelper.ToBoolean(rs, "require_passenger_weight_flag");
                        segment.SpecialServiceFeeFlag = RecordsetHelper.ToBoolean(rs, "special_service_fee_flag");
                        segment.ShowInsuranceOnWebFlag = RecordsetHelper.ToBoolean(rs, "show_insurance_on_web_flag");
                        segment.FlightId = RecordsetHelper.ToGuid(rs, "flight_id");
                        segment.ExchangedSegmentId = RecordsetHelper.ToGuid(rs, "exchanged_segment_id");
                        segment.AirlineRcd = RecordsetHelper.ToString(rs, "airline_rcd");
                        segment.FlightNumber = RecordsetHelper.ToString(rs, "flight_number");
                        segment.RefundableFlag = RecordsetHelper.ToByte(rs, "refundable_flag");
                        segment.GroupFlag = RecordsetHelper.ToByte(rs, "group_flag");
                        segment.NonRevenueFlag = RecordsetHelper.ToByte(rs, "non_revenue_flag");
                        segment.EticketFlag = RecordsetHelper.ToByte(rs, "eticket_flag");
                        segment.FareReduction = RecordsetHelper.ToByte(rs, "fare_reduction");
                        segment.DepartureDate = RecordsetHelper.ToDateTime(rs, "departure_date");
                        segment.ArrivalDate = RecordsetHelper.ToDateTime(rs, "arrival_date");
                        segment.OdOriginRcd = RecordsetHelper.ToString(rs, "od_origin_rcd");
                        segment.OdDestinationRcd = RecordsetHelper.ToString(rs, "od_destination_rcd");
                        segment.WaitlistFlag = RecordsetHelper.ToByte(rs, "waitlist_flag");
                        segment.OverbookFlag = RecordsetHelper.ToByte(rs, "overbook_flag");
                        segment.FlightConnectionId = RecordsetHelper.ToGuid(rs, "flight_connection_id");
                        segment.AdvancedPurchaseFlag = RecordsetHelper.ToByte(rs, "advanced_purchase_flag");
                        segment.JourneyTime = RecordsetHelper.ToInt16(rs, "journey_time");
                        segment.DepartureTime = RecordsetHelper.ToInt16(rs, "departure_time");
                        segment.ArrivalTime = RecordsetHelper.ToInt16(rs, "arrival_time");

                        segment.OriginName = RecordsetHelper.ToString(rs, "origin_name");
                        segment.DestinationName = RecordsetHelper.ToString(rs, "destination_name");

                        segment.TransitFlightId = RecordsetHelper.ToGuid(rs, "transit_flight_id");
                        segment.TransitFareId = RecordsetHelper.ToGuid(rs, "transit_fare_id");
                        segment.TransitDepartureDate = RecordsetHelper.ToDateTime(rs, "transit_departure_date");
                        segment.TransitArrivalDate = RecordsetHelper.ToDateTime(rs, "transit_arrival_date");
                        segment.FareId = RecordsetHelper.ToGuid(rs, "fare_id");
                        segment.BookingClassRcd = RecordsetHelper.ToString(rs, "booking_class_rcd");
                        segment.BoardingClassRcd = RecordsetHelper.ToString(rs, "boarding_class_rcd");
                        segment.AircraftTypeRcd = RecordsetHelper.ToString(rs, "aircraft_type_rcd");
                        segment.PlannedDepartureTime = RecordsetHelper.ToInt16(rs, "planned_departure_time");
                        segment.PlannedArrivalTime = RecordsetHelper.ToInt16(rs, "planned_arrival_time");
                        segment.DepartureDayofweek = RecordsetHelper.ToByte(rs, "departure_dayOfWeek");
                        segment.ArrivalDayofweek = RecordsetHelper.ToByte(rs, "arrival_dayOfWeek");
                        segment.DepartureMonth = RecordsetHelper.ToInt16(rs, "departure_month");
                        segment.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");

                        segment.SegmentStatusRcd = RecordsetHelper.ToString(rs, "segment_status_rcd");
                        segment.BookingId = RecordsetHelper.ToGuid(rs, "booking_id");
                        segment.NumberOfUnits = RecordsetHelper.ToInt16(rs, "number_of_units");

                        segment.SegmentChangeStatusRcd = RecordsetHelper.ToString(rs, "segment_change_status_rcd");
                        segment.InfoSegmentFlag = RecordsetHelper.ToByte(rs, "info_segment_flag");
                        segment.HighPriorityWaitlistFlag = RecordsetHelper.ToByte(rs, "high_priority_waitlist_flag");
                        segment.SegmentStatusName = RecordsetHelper.ToString(rs, "segment_status_name");

                        segment.SeatmapFlag = RecordsetHelper.ToByte(rs, "seatmap_flag");
                        segment.TempSeatmapFlag = RecordsetHelper.ToByte(rs, "temp_seatmap_flag");

                        segment.AllowWebCheckinFlag = RecordsetHelper.ToByte(rs, "allow_web_checkin_flag");
                        segment.CloseWebSalesFlag = RecordsetHelper.ToByte(rs, "close_web_sales_flag");
                        segment.ExcludeQuoteFlag = RecordsetHelper.ToByte(rs, "exclude_quote_flag");
                        segment.CurrencyRate = RecordsetHelper.ToDouble(rs, "currency_rate");
                        segment.OpenSequence = RecordsetHelper.ToByte(rs, "open_sequence");
                        segment.NumberOfStops = RecordsetHelper.ToByte(rs, "number_of_stops");
                        segment.FlightCheckInStatusRcd = RecordsetHelper.ToString(rs, "flight_check_in_status_rcd");
                        segment.FlightFlownDateTime = RecordsetHelper.ToDateTime(rs, "flight_flown_date_time");

                        segment.FlightStatusRCD = RecordsetHelper.ToString(rs, "flight_status_rcd");

                     //   flight_status_rcd
                        segmentList.Add(segment);

                        rs.MoveNext();

                    }

                }
                catch
                {
                    throw;
                }
            }
            return segmentList;

        }

        public static IList<Entity.Booking.Flight> FillBooking(this IList<Entity.Booking.Flight> flightList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                flightList = new List<Entity.Booking.Flight>();
                Entity.Booking.Flight flight = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        flight = new Entity.Booking.Flight();

                        flight.FlightId = RecordsetHelper.ToGuid(rs, "flight_id");
                        flight.FairId = RecordsetHelper.ToGuid(rs, "fare_id");
                        flight.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        flight.DestinationRcd = RecordsetHelper.ToString(rs, "destination_rcd");
                        flight.BoardingClassRcd = RecordsetHelper.ToString(rs, "boarding_class_rcd");
                        flight.BookingClassRcd = RecordsetHelper.ToString(rs, "booking_class_rcd");
                        flight.DepartureDate = RecordsetHelper.ToDateTime(rs, "departure_date");
                        //flight.FareReduction = RecordsetHelper.ToDecimal(rs, "fare_reduction");
                        //flight.WaitlistFlag = RecordsetHelper.ToBoolean(rs, "waitlist_flag");
                        //flight.RefundableFlag = RecordsetHelper.ToBoolean(rs, "refundable_flag");
                        //flight.Non_revenueFlag = RecordsetHelper.ToBoolean(rs, "non_revenue_flag");
                        flight.EticketFlag = RecordsetHelper.ToByte(rs, "eticket_flag");
                        //flight.GroupFlag = RecordsetHelper.ToBoolean(rs, "group_flag");
                        flight.ExchangedSegmentId = RecordsetHelper.ToGuid(rs, "exchanged_segment_id");
                        //flight.ExcludeQuoteFlag = RecordsetHelper.ToBoolean(rs, "exclude_quote_flag");
                        //flight.AdvancedPurchaseFlag = RecordsetHelper.ToBoolean(rs, "advanced_purchase_flag");

                        flightList.Add(flight);

                        rs.MoveNext();

                    }

                }
                catch
                {
                    throw;
                }
            }
            return flightList;

        }

        public static IList<Passenger> FillBooking(this IList<Passenger> passengerList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                passengerList = new List<Passenger>();
                Passenger passenger = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        passenger = new Passenger();

                        passenger.TitleRcd = RecordsetHelper.ToString(rs, "title_rcd");
                        passenger.Lastname = RecordsetHelper.ToString(rs, "lastname");
                        passenger.Firstname = RecordsetHelper.ToString(rs, "firstname");
                        passenger.Middlename = RecordsetHelper.ToString(rs, "middlename");
                        passenger.NationalityRcd = RecordsetHelper.ToString(rs, "nationality_rcd");
                        passenger.PassengerWeight = RecordsetHelper.ToDecimal(rs, "passenger_weight");
                        passenger.GenderTypeRcd = RecordsetHelper.ToString(rs, "gender_type_rcd");
                        passenger.PassengerTypeRcd = RecordsetHelper.ToString(rs, "passenger_type_rcd");
                        passenger.ClientProfileId = RecordsetHelper.ToGuid(rs, "client_profile_id");
                        passenger.ClientNumber = RecordsetHelper.ToInt64(rs, "client_number");

                        passenger.AddressLine1 = RecordsetHelper.ToString(rs, "address_line1");
                        passenger.AddressLine2 = RecordsetHelper.ToString(rs, "address_line2");
                        passenger.State = RecordsetHelper.ToString(rs, "state");
                        passenger.District = RecordsetHelper.ToString(rs, "district");
                        passenger.Province = RecordsetHelper.ToString(rs, "province");
                        passenger.ZipCode = RecordsetHelper.ToString(rs, "zip_code");
                        passenger.PoBox = RecordsetHelper.ToString(rs, "po_box");
                        passenger.CountryRcd = RecordsetHelper.ToString(rs, "country_rcd");
                        passenger.Street = RecordsetHelper.ToString(rs, "street");
                        passenger.City = RecordsetHelper.ToString(rs, "city");

                        passenger.DocumentTypeRcd = RecordsetHelper.ToString(rs, "document_type_rcd");
                        passenger.DocumentNumber = RecordsetHelper.ToString(rs, "document_number");
                        passenger.ResidenceCountryRcd = RecordsetHelper.ToString(rs, "residence_country_rcd");
                        passenger.PassportNumber = RecordsetHelper.ToString(rs, "passport_number");
                        passenger.PassportIssueDate = RecordsetHelper.ToDateTime(rs, "passport_issue_date");
                        passenger.PassportExpiryDate = RecordsetHelper.ToDateTime(rs, "passport_expiry_date");
                        passenger.PassportIssuePlace = RecordsetHelper.ToString(rs, "passport_issue_place");
                        passenger.PassportBirthPlace = RecordsetHelper.ToString(rs, "passport_birth_place");
                        passenger.DateOfBirth = RecordsetHelper.ToDateTime(rs, "date_of_birth");
                        passenger.PassportIssueCountryRcd = RecordsetHelper.ToString(rs, "passport_issue_country_rcd");
                        passenger.CountryCodeLong = RecordsetHelper.ToString(rs, "country_code_long");

                        passenger.ContactName = RecordsetHelper.ToString(rs, "contact_name");
                        passenger.ContactEmail = RecordsetHelper.ToString(rs, "contact_email");
                        passenger.MobileEmail = RecordsetHelper.ToString(rs, "mobile_email");
                        passenger.PhoneMobile = RecordsetHelper.ToString(rs, "phone_mobile");
                        passenger.PhoneHome = RecordsetHelper.ToString(rs, "phone_home");
                        passenger.PhoneFax = RecordsetHelper.ToString(rs, "phone_fax");
                        passenger.PhoneBusiness = RecordsetHelper.ToString(rs, "phone_business");
                        passenger.ReceivedFrom = RecordsetHelper.ToString(rs, "received_from");

                        passenger.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        passenger.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        passenger.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        passenger.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");

                        passenger.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                        passenger.BookingId = RecordsetHelper.ToGuid(rs, "booking_id");
                        passenger.PassengerProfileId = RecordsetHelper.ToGuid(rs, "passenger_profile_id");
                        passenger.EmployeeNumber = RecordsetHelper.ToString(rs, "employee_number");
                        passenger.PassengerRoleRcd = RecordsetHelper.ToString(rs, "passenger_role_rcd");
                        passenger.MemberLevelRcd = RecordsetHelper.ToString(rs, "member_level_rcd");
                        passenger.MemberNumber = RecordsetHelper.ToString(rs, "member_number");
                        passenger.RedressNumber = RecordsetHelper.ToString(rs, "redress_number");
                        passenger.PnrName = RecordsetHelper.ToString(rs, "pnr_name");
                        passenger.WheelchairFlag = RecordsetHelper.ToByte(rs, "wheelchair_flag");
                        passenger.VipFlag = RecordsetHelper.ToByte(rs, "vip_flag");
                        passenger.WindowSeatFlag = RecordsetHelper.ToByte(rs, "window_seat_flag");
                        passenger.GuardianPassengerId = RecordsetHelper.ToGuid(rs, "guardian_passenger_id");
                        passenger.MemberAirlineRcd = RecordsetHelper.ToString(rs, "member_airline_rcd");

                        passengerList.Add(passenger);

                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }

            }
            return passengerList;

        }

        public static IList<Fee> FillBooking(this IList<Fee> feeList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                feeList = new List<Fee>();
                Fee fee = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        fee = new Fee();
                        fee.BookingFeeId = RecordsetHelper.ToGuid(rs, "booking_fee_id");
                        fee.FeeAmount = RecordsetHelper.ToDecimal(rs, "fee_amount");
                        fee.BookingId = RecordsetHelper.ToGuid(rs, "booking_id");
                        fee.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                        fee.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        fee.AcctFeeAmount = RecordsetHelper.ToDecimal(rs, "acct_fee_amount");
                        fee.FeeId = RecordsetHelper.ToGuid(rs, "fee_id");
                        fee.VatPercentage = RecordsetHelper.ToDecimal(rs, "vat_percentage");
                        fee.FeeAmountIncl = RecordsetHelper.ToDecimal(rs, "fee_amount_incl");
                        fee.AcctFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_fee_amount_incl");
                        fee.FeeRcd = RecordsetHelper.ToString(rs, "fee_rcd");
                        fee.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        fee.AccountFeeBy = RecordsetHelper.ToGuid(rs, "account_fee_by");
                        fee.AccountFeeDateTime = RecordsetHelper.ToDateTime(rs, "account_fee_date_time");
                        fee.VoidDateTime = RecordsetHelper.ToDateTime(rs, "void_date_time");
                        fee.VoidBy = RecordsetHelper.ToGuid(rs, "void_by");
                        fee.PaymentAmount = RecordsetHelper.ToDecimal(rs, "payment_amount");
                        fee.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        fee.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        fee.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        fee.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        fee.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                        fee.AgencyCode = RecordsetHelper.ToString(rs, "agency_code");
                        fee.PassengerSegmentServiceId = RecordsetHelper.ToGuid(rs, "passenger_segment_service_id");
                        fee.FeeCategoryRcd = RecordsetHelper.ToString(rs, "fee_category_rcd");
                        fee.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        fee.DestinationRcd = RecordsetHelper.ToString(rs, "destination_rcd");
                        fee.OdOriginRcd = RecordsetHelper.ToString(rs, "od_origin_rcd");
                        fee.OdDestinationRcd = RecordsetHelper.ToString(rs, "od_destination_rcd");
                        fee.NumberOfUnits = RecordsetHelper.ToDecimal(rs, "number_of_units");
                        fee.TotalAmount = RecordsetHelper.ToDecimal(rs, "total_amount");
                        fee.TotalAmountIncl = RecordsetHelper.ToDecimal(rs, "total_amount_incl");
                        fee.ManualFeeFlag = RecordsetHelper.ToByte(rs, "manual_fee_flag");
                        fee.OdFlag = RecordsetHelper.ToByte(rs, "od_flag");
                        fee.SkipFareAllowanceFlag = RecordsetHelper.ToByte(rs, "skip_fare_allowance_flag");
                        fee.FeeLevel = RecordsetHelper.ToString(rs, "fee_level");
                        fee.FeeCalculationRcd = RecordsetHelper.ToString(rs, "fee_calculation_rcd");
                        fee.MinimumFeeAmountFlag = RecordsetHelper.ToByte(rs, "minimum_fee_amount_flag");
                        fee.FeePercentage = RecordsetHelper.ToDecimal(rs, "fee_percentage");
                        fee.ChangeComment = RecordsetHelper.ToString(rs, "change_comment");
                        fee.Comment = RecordsetHelper.ToString(rs, "comment");
                        fee.TotalFeeAmount = RecordsetHelper.ToDecimal(rs, "total_fee_amount");
                        fee.TotalFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "total_fee_amount_incl");
                        fee.Units = RecordsetHelper.ToString(rs, "units");
                        fee.ChargeCurrencyRcd = RecordsetHelper.ToString(rs, "charge_currency_rcd");
                        fee.ChargeAmount = RecordsetHelper.ToDecimal(rs, "charge_amount");
                        fee.ChargeAmountIncl = RecordsetHelper.ToDecimal(rs, "charge_amount_incl");
                        fee.DocumentNumber = RecordsetHelper.ToString(rs, "document_number");
                        fee.DocumentDateTime = RecordsetHelper.ToDateTime(rs, "document_date_time");
                        fee.ExternalReference = RecordsetHelper.ToString(rs, "external_reference");
                        fee.VendorRcd = RecordsetHelper.ToString(rs, "vendor_rcd");
                        fee.BaggageFeeOptionId = RecordsetHelper.ToGuid(rs, "baggage_fee_option_id");
                        fee.NewRecord = RecordsetHelper.ToBoolean(rs, "new_record");
                        fee.SelectedFee = RecordsetHelper.ToBoolean(rs, "selected_fee");

                        fee.WeightLbs = RecordsetHelper.ToDecimal(rs, "weight_lbs");
                        fee.WeightKgs = RecordsetHelper.ToDecimal(rs, "weight_kgs");
                        fee.MpdNumber = RecordsetHelper.ToString(rs, "mpd_number");
                        fee.UnitRcd = RecordsetHelper.ToString(rs, "unit_rcd");


                        feeList.Add(fee);

                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }

            }
            return feeList;

        }
        public static IList<Fee> FillNameChangeFee(this IList<Fee> feeList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                feeList = new List<Fee>();
                Fee fee = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        fee = new Fee();
                        fee.BookingFeeId = RecordsetHelper.ToGuid(rs, "booking_fee_id");
                        fee.FeeAmount = RecordsetHelper.ToDecimal(rs, "fee_amount");
                        fee.BookingId = RecordsetHelper.ToGuid(rs, "booking_id");
                        fee.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                        fee.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        fee.AcctFeeAmount = RecordsetHelper.ToDecimal(rs, "acct_fee_amount");
                        fee.FeeId = RecordsetHelper.ToGuid(rs, "fee_id");
                        fee.VatPercentage = RecordsetHelper.ToDecimal(rs, "vat_percentage");
                        fee.FeeAmountIncl = RecordsetHelper.ToDecimal(rs, "fee_amount_incl");
                        fee.AcctFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_fee_amount_incl");
                        fee.FeeRcd = RecordsetHelper.ToString(rs, "fee_rcd");
                        fee.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        fee.AccountFeeBy = RecordsetHelper.ToGuid(rs, "account_fee_by");
                        fee.AccountFeeDateTime = RecordsetHelper.ToDateTime(rs, "account_fee_date_time");
                        fee.VoidDateTime = RecordsetHelper.ToDateTime(rs, "void_date_time");
                        fee.VoidBy = RecordsetHelper.ToGuid(rs, "void_by");
                        fee.PaymentAmount = RecordsetHelper.ToDecimal(rs, "payment_amount");
                        fee.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        fee.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        fee.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        fee.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        fee.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                        fee.AgencyCode = RecordsetHelper.ToString(rs, "agency_code");
                        fee.PassengerSegmentServiceId = RecordsetHelper.ToGuid(rs, "passenger_segment_service_id");
                        fee.FeeCategoryRcd = RecordsetHelper.ToString(rs, "fee_category_rcd");
                        fee.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        fee.DestinationRcd = RecordsetHelper.ToString(rs, "destination_rcd");
                        fee.OdOriginRcd = RecordsetHelper.ToString(rs, "od_origin_rcd");
                        fee.OdDestinationRcd = RecordsetHelper.ToString(rs, "od_destination_rcd");
                        fee.NumberOfUnits = RecordsetHelper.ToDecimal(rs, "number_of_units");
                        fee.TotalAmount = RecordsetHelper.ToDecimal(rs, "total_amount");
                        fee.TotalAmountIncl = RecordsetHelper.ToDecimal(rs, "total_amount_incl");
                        fee.ManualFeeFlag = RecordsetHelper.ToByte(rs, "manual_fee_flag");
                        fee.OdFlag = RecordsetHelper.ToByte(rs, "od_flag");
                        fee.SkipFareAllowanceFlag = RecordsetHelper.ToByte(rs, "skip_fare_allowance_flag");
                        fee.FeeLevel = RecordsetHelper.ToString(rs, "fee_level");
                        fee.FeeCalculationRcd = RecordsetHelper.ToString(rs, "fee_calculation_rcd");
                        fee.MinimumFeeAmountFlag = RecordsetHelper.ToByte(rs, "minimum_fee_amount_flag");
                        fee.FeePercentage = RecordsetHelper.ToDecimal(rs, "fee_percentage");
                        fee.ChangeComment = RecordsetHelper.ToString(rs, "change_comment");
                        fee.Comment = RecordsetHelper.ToString(rs, "comment");
                        fee.TotalFeeAmount = RecordsetHelper.ToDecimal(rs, "total_fee_amount");
                        fee.TotalFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "total_fee_amount_incl");
                        fee.Units = RecordsetHelper.ToString(rs, "units");
                        fee.ChargeCurrencyRcd = RecordsetHelper.ToString(rs, "charge_currency_rcd");
                        fee.ChargeAmount = RecordsetHelper.ToDecimal(rs, "charge_amount");
                        fee.ChargeAmountIncl = RecordsetHelper.ToDecimal(rs, "charge_amount_incl");
                        fee.DocumentNumber = RecordsetHelper.ToString(rs, "document_number");
                        fee.DocumentDateTime = RecordsetHelper.ToDateTime(rs, "document_date_time");
                        fee.ExternalReference = RecordsetHelper.ToString(rs, "external_reference");
                        fee.VendorRcd = RecordsetHelper.ToString(rs, "vendor_rcd");
                        fee.BaggageFeeOptionId = RecordsetHelper.ToGuid(rs, "baggage_fee_option_id");
                        fee.NewRecord = RecordsetHelper.ToBoolean(rs, "new_record");
                        fee.SelectedFee = RecordsetHelper.ToBoolean(rs, "selected_fee");

                        fee.WeightLbs = RecordsetHelper.ToDecimal(rs, "weight_lbs");
                        fee.WeightKgs = RecordsetHelper.ToDecimal(rs, "weight_kgs");
                        fee.MpdNumber = RecordsetHelper.ToString(rs, "mpd_number");
                        fee.UnitRcd = RecordsetHelper.ToString(rs, "unit_rcd");

                        if (fee.PaymentAmount == 0 && fee.VoidBy == Guid.Empty)
                            feeList.Add(fee);

                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }

            }
            return feeList;
        }

        public static IList<Remark> FillBooking(this IList<Remark> remarkList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                remarkList = new List<Remark>();
                Remark remark = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        remark = new Remark();
                        remark.BookingRemarkId = RecordsetHelper.ToGuid(rs, "booking_remark_id");
                        remark.RemarkTypeRcd = RecordsetHelper.ToString(rs, "remark_type_rcd");
                        remark.TimelimitDateTime = RecordsetHelper.ToDateTime(rs, "timelimit_date_time");
                        remark.CompleteFlag = RecordsetHelper.ToByte(rs, "complete_flag");
                        remark.RemarkText = RecordsetHelper.ToString(rs, "remark_text");
                        remark.BookingId = RecordsetHelper.ToGuid(rs, "booking_id");
                        remark.AddedBy = RecordsetHelper.ToString(rs, "added_by");
                        remark.ClientProfileId = RecordsetHelper.ToGuid(rs, "client_profile_id");
                        remark.AgencyCode = RecordsetHelper.ToString(rs, "agency_code");
                        remark.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        remark.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        remark.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        remark.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        remark.SystemFlag = RecordsetHelper.ToByte(rs, "system_flag");
                        remark.UtcTimelimitDateTime = RecordsetHelper.ToDateTime(rs, "utc_timelimit_date_time");
                        remark.Nickname = RecordsetHelper.ToString(rs, "nickname");
                        remark.ProtectedFlag = RecordsetHelper.ToByte(rs, "protected_flag");
                        remark.WarningFlag = RecordsetHelper.ToByte(rs, "warning_flag");
                        remark.ProcessMessageFlag = RecordsetHelper.ToByte(rs, "process_message_flag");

                        remarkList.Add(remark);

                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }


            }
            return remarkList;

        }

        public static IList<Quote> FillBooking(this IList<Quote> quoteList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                quoteList = new List<Quote>();
                Quote quote = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        quote = new Quote();
                        quote.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                        quote.PassengerTypeRcd = RecordsetHelper.ToString(rs, "passenger_type_rcd");
                        quote.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        quote.ChargeType = RecordsetHelper.ToString(rs, "charge_type");
                        quote.ChargeName = RecordsetHelper.ToString(rs, "charge_name");
                        quote.ChargeAmount = RecordsetHelper.ToDecimal(rs, "charge_amount");
                        quote.TotalAmount = RecordsetHelper.ToDecimal(rs, "total_amount");
                        quote.TaxAmount = RecordsetHelper.ToDecimal(rs, "tax_amount");
                        quote.PassengerCount = RecordsetHelper.ToInt16(rs, "passenger_count");
                        quote.SortSequence = RecordsetHelper.ToInt16(rs, "sort_sequence");
                        quote.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        quote.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        quote.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        quote.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        quote.RedemptionPoints = RecordsetHelper.ToDecimal(rs, "redemption_points");
                        quote.ChargeAmountIncl = RecordsetHelper.ToDecimal(rs, "charge_amount_incl");

                        quoteList.Add(quote);

                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }


            }
            return quoteList;

        }

        public static IList<Tax> FillBooking(this IList<Tax> taxList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                taxList = new List<Tax>();
                Tax tax = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        tax = new Tax();
                        tax.PassengerSegmentTaxId = RecordsetHelper.ToGuid(rs, "passenger_segment_tax_id");
                        tax.TaxAmount = RecordsetHelper.ToDecimal(rs, "tax_amount");
                        tax.TaxAmountIncl = RecordsetHelper.ToDecimal(rs, "tax_amount_incl");
                        tax.AcctAmount = RecordsetHelper.ToDecimal(rs, "acct_amount");
                        tax.AcctAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_amount_incl");
                        tax.SalesAmount = RecordsetHelper.ToDecimal(rs, "sales_amount");
                        tax.SalesAmountIncl = RecordsetHelper.ToDecimal(rs, "sales_amount_incl");
                        tax.TaxRcd = RecordsetHelper.ToString(rs, "tax_rcd");
                        tax.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                        tax.TaxId = RecordsetHelper.ToGuid(rs, "tax_id");
                        tax.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                        tax.TaxCurrencyRcd = RecordsetHelper.ToString(rs, "tax_currency_rcd");
                        tax.SalesCurrencyRcd = RecordsetHelper.ToString(rs, "sales_currency_rcd");
                        tax.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        tax.SummarizeUp = RecordsetHelper.ToString(rs, "summarize_up");
                        tax.CoverageType = RecordsetHelper.ToString(rs, "coverage_type");
                        tax.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        tax.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        tax.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        tax.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        tax.VatPercentage = RecordsetHelper.ToDecimal(rs, "vat_percentage");

                        taxList.Add(tax);

                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }


            }
            return taxList;

        }

        public static IList<PassengerService> FillBooking(this IList<PassengerService> serviceList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                serviceList = new List<PassengerService>();
                PassengerService service = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        service = new PassengerService();
                        service.PassengerSegmentServiceId = RecordsetHelper.ToGuid(rs, "passenger_segment_service_id");
                        service.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                        service.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                        service.SpecialServiceStatusRcd = RecordsetHelper.ToString(rs, "special_service_status_rcd");
                        service.SpecialServiceChangeStatusRcd = RecordsetHelper.ToString(rs, "special_service_change_status_rcd");
                        service.SpecialServiceRcd = RecordsetHelper.ToString(rs, "special_service_rcd");
                        service.ServiceText = RecordsetHelper.ToString(rs, "service_text");
                        service.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        service.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        service.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        service.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        service.FlightId = RecordsetHelper.ToString(rs, "flight_id");
                        service.FeeId = RecordsetHelper.ToString(rs, "fee_id");
                        service.NumberOfUnits = RecordsetHelper.ToInt16(rs, "number_of_units");
                        service.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        service.DestinationRcd = RecordsetHelper.ToString(rs, "destination_rcd");
                        service.NewRecord = RecordsetHelper.ToBoolean(rs, "new_record");
                        service.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        service.CutOffTime = RecordsetHelper.ToInt16(rs, "cut_off_time");
                        service.StatusCode = RecordsetHelper.ToString(rs, "status_code");
                        service.HelpText = RecordsetHelper.ToString(rs, "help_text");
                        service.SpecialServiceGroupRcd = RecordsetHelper.ToString(rs, "special_service_group_rcd");
                        service.ServiceOnRequestFlag = RecordsetHelper.ToByte(rs, "service_on_request_flag");
                        service.TextAllowedFlag = RecordsetHelper.ToByte(rs, "text_allowed_flag");
                        service.ManifestFlag = RecordsetHelper.ToByte(rs, "manifest_flag");
                        service.TextRequiredFlag = RecordsetHelper.ToByte(rs, "text_required_flag");
                        service.IncludePassengerNameFlag = RecordsetHelper.ToByte(rs, "include_passenger_name_flag");
                        service.IncludeFlightSegmentFlag = RecordsetHelper.ToByte(rs, "include_flight_segment_flag");
                        service.IncludeActionCodeFlag = RecordsetHelper.ToByte(rs, "include_action_code_flag");
                        service.IncludeNumberOfServiceFlag = RecordsetHelper.ToByte(rs, "include_number_of_service_flag");
                        service.IncludeCateringFlag = RecordsetHelper.ToByte(rs, "include_catering_flag");
                        service.IncludePassengerAssistanceFlag = RecordsetHelper.ToByte(rs, "include_passenger_assistance_flag");
                        service.ServiceSupportedFlag = RecordsetHelper.ToByte(rs, "service_supported_flag");
                        service.SendInterlineReplyFlag = RecordsetHelper.ToByte(rs, "send_interline_reply_flag");
                        service.InventoryControlFlag = RecordsetHelper.ToByte(rs, "inventory_control_flag");

                        service.UnitRcd = RecordsetHelper.ToString(rs, "unit_rcd");

                        serviceList.Add(service);

                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }


            }
            return serviceList;

        }

        public static IList<Payment> FillBooking(this IList<Payment> paymentList, ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                paymentList = new List<Payment>();
                Payment payment = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        payment = new Payment();

                        payment.BookingPaymentId = RecordsetHelper.ToGuid(rs, "booking_payment_id");
                        payment.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                        payment.BookingId = RecordsetHelper.ToGuid(rs, "booking_id");
                        payment.VoucherPaymentId = RecordsetHelper.ToGuid(rs, "voucher_payment_id");
                        payment.FormOfPaymentRcd = RecordsetHelper.ToString(rs, "form_of_payment_rcd");
                        payment.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        payment.ReceiveCurrencyRcd = RecordsetHelper.ToString(rs, "receive_currency_rcd");
                        payment.AgencyPaymentTypeRcd = RecordsetHelper.ToString(rs, "agency_payment_type_rcd");
                        payment.AgencyCode = RecordsetHelper.ToString(rs, "agency_code");
                        payment.DebitAgencyCode = RecordsetHelper.ToString(rs, "debit_agency_code");
                        payment.PaymentAmount = RecordsetHelper.ToDecimal(rs, "payment_amount");
                        payment.ReceivePaymentAmount = RecordsetHelper.ToDecimal(rs, "receive_payment_amount");
                        payment.AcctPaymentAmount = RecordsetHelper.ToDecimal(rs, "acct_payment_amount");
                        payment.PaymentBy = RecordsetHelper.ToGuid(rs, "payment_by");
                        payment.PaymentDateTime = RecordsetHelper.ToDateTime(rs, "payment_date_time");
                        payment.PaymentDueDateTime = RecordsetHelper.ToDateTime(rs, "payment_due_date_time");
                        payment.DocumentAmount = RecordsetHelper.ToDecimal(rs, "document_amount");
                        payment.VoidBy = RecordsetHelper.ToGuid(rs, "void_by");
                        payment.ExpiryMonth = RecordsetHelper.ToInt16(rs, "expiry_month");
                        payment.ExpiryYear = RecordsetHelper.ToInt16(rs, "expiry_year");
                        payment.VoidDateTime = RecordsetHelper.ToDateTime(rs, "void_date_time");
                        payment.RecordLocator = RecordsetHelper.ToString(rs, "record_locator");
                        payment.CvvCode = RecordsetHelper.ToString(rs, "cvv_code");
                        payment.NameOnCard = RecordsetHelper.ToString(rs, "name_on_card");
                        payment.DocumentNumber = RecordsetHelper.ToString(rs, "document_number");
                        payment.FormOfPaymentSubtypeRcd = RecordsetHelper.ToString(rs, "form_of_payment_subtype_rcd");
                        payment.City = RecordsetHelper.ToString(rs, "city");
                        payment.State = RecordsetHelper.ToString(rs, "state");
                        payment.Street = RecordsetHelper.ToString(rs, "street");
                        payment.PoBox = RecordsetHelper.ToString(rs, "po_box");
                        payment.AddressLine1 = RecordsetHelper.ToString(rs, "address_line1");
                        payment.AddressLine2 = RecordsetHelper.ToString(rs, "address_line2");
                        payment.District = RecordsetHelper.ToString(rs, "district");
                        payment.Province = RecordsetHelper.ToString(rs, "province");
                        payment.ZipCode = RecordsetHelper.ToString(rs, "zip_code");
                        payment.CountryRcd = RecordsetHelper.ToString(rs, "country_rcd");
                        payment.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        payment.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        payment.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        payment.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        payment.PosIndicator = RecordsetHelper.ToString(rs, "pos_indicator");
                        payment.IssueMonth = RecordsetHelper.ToInt16(rs, "issue_month");
                        payment.IssueYear = RecordsetHelper.ToInt16(rs, "issue_year");
                        payment.IssueNumber = RecordsetHelper.ToString(rs, "issue_number");
                        payment.ClientProfileId = RecordsetHelper.ToGuid(rs, "client_profile_id");
                        payment.TransactionReference = RecordsetHelper.ToString(rs, "transaction_reference");
                        payment.ApprovalCode = RecordsetHelper.ToString(rs, "approval_code");
                        payment.BankName = RecordsetHelper.ToString(rs, "bank_name");
                        payment.BankCode = RecordsetHelper.ToString(rs, "bank_code");
                        payment.BankIban = RecordsetHelper.ToString(rs, "bank_iban");
                        payment.IpAddress = RecordsetHelper.ToString(rs, "ip_address");
                        payment.PaymentReference = RecordsetHelper.ToString(rs, "payment_reference");
                        payment.AllocatedAmount = RecordsetHelper.ToDecimal(rs, "allocated_amount");
                        payment.PaymentTypeRcd = RecordsetHelper.ToString(rs, "payment_type_rcd");
                        payment.RefundedAmount = RecordsetHelper.ToDecimal(rs, "refunded_amount");

                        payment.PaymentNumber = RecordsetHelper.ToString(rs, "payment_number");
                        payment.PaymentRemark = RecordsetHelper.ToString(rs, "payment_remark");

                        paymentList.Add(payment);

                        rs.MoveNext();
                    }
                }
                catch
                {
                    throw;
                }
            }
            return paymentList;

        }

        public static IList<Mapping> FillBooking(this IList<Mapping> mappingList, ADODB.Recordset rs)
        {
            try
            {
                if (rs != null && rs.RecordCount > 0)
                {
                    mappingList = new List<Mapping>();
                    Mapping mapping = null;

                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        mapping = new Mapping();

                        mapping.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                        mapping.DestinationRcd = RecordsetHelper.ToString(rs, "destination_rcd");
                        mapping.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        mapping.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                        mapping.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                        mapping.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                        mapping.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                        #region Origin Flag
                        mapping.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        mapping.RoutesTot = RecordsetHelper.ToInt16(rs, "routes_tot");
                        mapping.RoutesAvl = RecordsetHelper.ToInt16(rs, "routes_avl");
                        mapping.RoutesB2c = RecordsetHelper.ToInt16(rs, "routes_b2c");
                        mapping.RoutesB2b = RecordsetHelper.ToInt16(rs, "routes_b2b");
                        mapping.RoutesB2s = RecordsetHelper.ToInt16(rs, "routes_b2s");
                        mapping.RoutesApi = RecordsetHelper.ToInt16(rs, "routes_api");
                        mapping.RoutesB2t = RecordsetHelper.ToInt16(rs, "routes_b2t");
                        #endregion
                        #region Destination Flag
                        mapping.SegmentChangeFeeFlag = RecordsetHelper.ToBoolean(rs, "segment_change_fee_flag");
                        mapping.TransitFlag = RecordsetHelper.ToBoolean(rs, "transit_flag");
                        mapping.DirectFlag = RecordsetHelper.ToBoolean(rs, "direct_flag");
                        mapping.AvlFlag = RecordsetHelper.ToBoolean(rs, "avl_flag");
                        mapping.B2cFlag = RecordsetHelper.ToBoolean(rs, "b2c_flag");
                        mapping.B2bFlag = RecordsetHelper.ToBoolean(rs, "b2b_flag");
                        mapping.B2tFlag = RecordsetHelper.ToBoolean(rs, "b2t_flag");
                        mapping.DayRange = RecordsetHelper.ToInt16(rs, "day_range");
                        mapping.ShowRedressNumberFlag = RecordsetHelper.ToBoolean(rs, "show_redress_number_flag");
                        mapping.RequirePassengerTitleFlag = RecordsetHelper.ToBoolean(rs, "require_passenger_title_flag");
                        mapping.RequirePassengerGenderFlag = RecordsetHelper.ToBoolean(rs, "require_passenger_gender_flag");
                        mapping.RequireDateOfBirthFlag = RecordsetHelper.ToBoolean(rs, "require_date_of_birth_flag");
                        mapping.RequireDocumentDetailsFlag = RecordsetHelper.ToBoolean(rs, "require_document_details_flag");
                        mapping.RequirePassengerWeightFlag = RecordsetHelper.ToBoolean(rs, "require_passenger_weight_flag");
                        mapping.SpecialServiceFeeFlag = RecordsetHelper.ToBoolean(rs, "special_service_fee_flag");
                        mapping.ShowInsuranceOnWebFlag = RecordsetHelper.ToBoolean(rs, "show_insurance_on_web_flag");
                        #endregion

                        #region "Property"
                        mapping.FlightId = RecordsetHelper.ToGuid(rs, "flight_id");
                        mapping.ExchangedSegmentId = RecordsetHelper.ToGuid(rs, "exchanged_segment_id");
                        mapping.AirlineRcd = RecordsetHelper.ToString(rs, "airline_rcd");
                        mapping.FlightNumber = RecordsetHelper.ToString(rs, "flight_number");
                        mapping.RefundableFlag = RecordsetHelper.ToByte(rs, "refundable_flag");
                        mapping.GroupFlag = RecordsetHelper.ToByte(rs, "group_flag");
                        mapping.NonRevenueFlag = RecordsetHelper.ToByte(rs, "non_revenue_flag");
                        mapping.EticketFlag = RecordsetHelper.ToByte(rs, "eticket_flag");
                        mapping.FareReduction = RecordsetHelper.ToByte(rs, "fare_reduction");
                        mapping.DepartureDate = RecordsetHelper.ToDateTime(rs, "departure_date");
                        mapping.ArrivalDate = RecordsetHelper.ToDateTime(rs, "arrival_date");
                        mapping.OdOriginRcd = RecordsetHelper.ToString(rs, "od_origin_rcd");
                        mapping.OdDestinationRcd = RecordsetHelper.ToString(rs, "od_destination_rcd");
                        mapping.WaitlistFlag = RecordsetHelper.ToByte(rs, "waitlist_flag");
                        mapping.OverbookFlag = RecordsetHelper.ToByte(rs, "overbook_flag");
                        mapping.FlightConnectionId = RecordsetHelper.ToGuid(rs, "flight_connection_id");
                        mapping.AdvancedPurchaseFlag = RecordsetHelper.ToByte(rs, "advanced_purchase_flag");
                        mapping.JourneyTime = RecordsetHelper.ToInt16(rs, "journey_time");
                        mapping.DepartureTime = RecordsetHelper.ToInt16(rs, "departure_time");
                        mapping.ArrivalTime = RecordsetHelper.ToInt16(rs, "arrival_time");

                        mapping.TransitFlightId = RecordsetHelper.ToGuid(rs, "transit_flight_id");
                        mapping.TransitFareId = RecordsetHelper.ToGuid(rs, "transit_fare_id");
                        mapping.TransitDepartureDate = RecordsetHelper.ToDateTime(rs, "transit_departure_date");
                        mapping.TransitArrivalDate = RecordsetHelper.ToDateTime(rs, "transit_arrival_date");
                        mapping.FareId = RecordsetHelper.ToGuid(rs, "fare_id");
                        mapping.BookingClassRcd = RecordsetHelper.ToString(rs, "booking_class_rcd");
                        mapping.BoardingClassRcd = RecordsetHelper.ToString(rs, "boarding_class_rcd");

                        mapping.PlannedDepartureTime = RecordsetHelper.ToInt16(rs, "planned_departure_time");
                        mapping.PlannedArrivalTime = RecordsetHelper.ToInt16(rs, "planned_arrival_time");
                        mapping.DepartureDayofweek = RecordsetHelper.ToByte(rs, "departure_dayOfWeek");
                        mapping.ArrivalDayofweek = RecordsetHelper.ToByte(rs, "arrival_dayOfWeek");
                        mapping.DepartureMonth = RecordsetHelper.ToInt16(rs, "departure_month");
                        #endregion

                        #region "Property"

                        mapping.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                        mapping.BookingId = RecordsetHelper.ToGuid(rs, "booking_id");
                        mapping.NumberOfUnits = RecordsetHelper.ToInt16(rs, "number_of_units");
                        mapping.InfoSegmentFlag = RecordsetHelper.ToByte(rs, "info_segment_flag");
                        mapping.HighPriorityWaitlistFlag = RecordsetHelper.ToByte(rs, "high_priority_waitlist_flag");
                        mapping.SeatmapFlag = RecordsetHelper.ToByte(rs, "seatmap_flag");
                        mapping.TempSeatmapFlag = RecordsetHelper.ToByte(rs, "temp_seatmap_flag");
                        mapping.AllowWebCheckinFlag = RecordsetHelper.ToByte(rs, "allow_web_checkin_flag");
                        mapping.CloseWebSalesFlag = RecordsetHelper.ToInt16(rs, "close_web_sales_flag");
                        mapping.ExcludeQuoteFlag = RecordsetHelper.ToInt16(rs, "exclude_quote_flag");
                        mapping.CurrencyRate = RecordsetHelper.ToDouble(rs, "currency_rate");
                        mapping.OpenSequence = RecordsetHelper.ToByte(rs, "open_sequence");
                        mapping.NumberOfStops = RecordsetHelper.ToByte(rs, "number_of_stops");
                        #endregion


                        #region Passenger information
                        mapping.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                        mapping.Lastname = RecordsetHelper.ToString(rs, "lastname");
                        mapping.Firstname = RecordsetHelper.ToString(rs, "firstname");
                        mapping.GenderTypeRcd = RecordsetHelper.ToString(rs, "gender_type_rcd");
                        mapping.TitleRcd = RecordsetHelper.ToString(rs, "title_rcd");
                        mapping.PassengerTypeRcd = RecordsetHelper.ToString(rs, "passenger_type_rcd");
                        mapping.DateOfBirth = RecordsetHelper.ToDateTime(rs, "date_of_birth");
                        mapping.PassengerStatusRcd = RecordsetHelper.ToString(rs, "passenger_status_rcd");
                        mapping.Middlename = RecordsetHelper.ToString(rs, "middlename");
                        mapping.DocumentTypeRcd = RecordsetHelper.ToString(rs, "document_type_rcd");
                        mapping.PassportNumber = RecordsetHelper.ToString(rs, "passport_number");
                        mapping.PassportIssuePlace = RecordsetHelper.ToString(rs, "passport_issue_place");
                        mapping.PassportIssueDate = RecordsetHelper.ToDateTime(rs, "passport_issue_date");
                        mapping.PassportExpiryDate = RecordsetHelper.ToDateTime(rs, "passport_expiry_date");
                        mapping.Sendmail = RecordsetHelper.ToString(rs, "sendmail");
                        mapping.ClientNumber = RecordsetHelper.ToInt64(rs, "client_number");
                        mapping.PassengerProfileId = RecordsetHelper.ToGuid(rs, "passenger_profile_id");
                        mapping.ClientProfileId = RecordsetHelper.ToGuid(rs, "client_profile_id");
                        mapping.EmployeeNumber = RecordsetHelper.ToString(rs, "employee_number");
                        mapping.NationalityRcd = RecordsetHelper.ToString(rs, "nationality_rcd");
                        mapping.ContactEmail = RecordsetHelper.ToString(rs, "contact_email");

                        #endregion
                        #region Ticket information
                        mapping.InventoryClassRcd = RecordsetHelper.ToString(rs, "inventory_class_rcd");
                        mapping.SeatNumber = RecordsetHelper.ToString(rs, "seat_number");
                        mapping.SeatRow = RecordsetHelper.ToInt16(rs, "seat_row");
                        mapping.SeatColumn = RecordsetHelper.ToString(rs, "seat_column");
                        mapping.NetTotal = RecordsetHelper.ToDecimal(rs, "net_total");
                        mapping.TaxAmount = RecordsetHelper.ToDecimal(rs, "tax_amount");
                        mapping.FareAmount = RecordsetHelper.ToDecimal(rs, "fare_amount");
                        mapping.YqAmount = RecordsetHelper.ToDecimal(rs, "yq_amount");
                        mapping.BaseTicketingFeeAmount = RecordsetHelper.ToDecimal(rs, "base_ticketing_fee_amount");
                        mapping.TicketingFeeAmount = RecordsetHelper.ToDecimal(rs, "ticketing_fee_amount");
                        mapping.ReservationFeeAmount = RecordsetHelper.ToDecimal(rs, "reservation_fee_amount");
                        mapping.CommissionAmount = RecordsetHelper.ToDecimal(rs, "commission_amount");
                        mapping.FareVat = RecordsetHelper.ToDecimal(rs, "fare_vat");
                        mapping.TaxVat = RecordsetHelper.ToDecimal(rs, "tax_vat");
                        mapping.YqVat = RecordsetHelper.ToDecimal(rs, "yq_vat");
                        mapping.TicketingFeeVat = RecordsetHelper.ToDecimal(rs, "ticketing_fee_vat");
                        mapping.ReservationFeeVat = RecordsetHelper.ToDecimal(rs, "reservation_fee_vat");
                        mapping.FareAmountIncl = RecordsetHelper.ToDecimal(rs, "fare_amount_incl");
                        mapping.TaxAmountIncl = RecordsetHelper.ToDecimal(rs, "tax_amount_incl");
                        mapping.YqAmountIncl = RecordsetHelper.ToDecimal(rs, "yq_amount_incl");
                        mapping.PublicFareAmountIncl = RecordsetHelper.ToDecimal(rs, "public_fare_amount_incl");
                        mapping.PublicFareAmount = RecordsetHelper.ToDecimal(rs, "public_fare_amount");
                        mapping.CommissionAmountIncl = RecordsetHelper.ToDecimal(rs, "commission_amount_incl");
                        mapping.CommissionPercentage = RecordsetHelper.ToDecimal(rs, "commission_percentage");
                        mapping.TicketingFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "ticketing_fee_amount_incl");
                        mapping.ReservationFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "reservation_fee_amount_incl");
                        mapping.AcctNetTotal = RecordsetHelper.ToDecimal(rs, "acct_net_total");
                        mapping.AcctTaxAmount = RecordsetHelper.ToDecimal(rs, "acct_tax_amount");
                        mapping.AcctFareAmount = RecordsetHelper.ToDecimal(rs, "acct_fare_amount");
                        mapping.AcctYqAmount = RecordsetHelper.ToDecimal(rs, "acct_yq_amount");
                        mapping.AcctTicketingFeeAmount = RecordsetHelper.ToDecimal(rs, "acct_ticketing_fee_amount");
                        mapping.AcctReservationFeeAmount = RecordsetHelper.ToDecimal(rs, "acct_reservation_fee_amount");
                        mapping.AcctCommissionAmount = RecordsetHelper.ToDecimal(rs, "acct_commission_amount");
                        mapping.AcctFareAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_fare_amount_incl");
                        mapping.AcctTaxAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_tax_amount_incl");
                        mapping.AcctYqAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_yq_amount_incl");
                        mapping.AcctCommissionAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_commission_amount_incl");
                        mapping.AcctTicketingFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_ticketing_fee_amount_incl");
                        mapping.AcctReservationFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_reservation_fee_amount_incl");
                        mapping.FareCode = RecordsetHelper.ToString(rs, "fare_code");
                        mapping.FareDateTime = RecordsetHelper.ToDateTime(rs, "fare_date_time");
                        mapping.ETicketFlag = RecordsetHelper.ToByte(rs, "e_ticket_flag");
                        mapping.StandbyFlag = RecordsetHelper.ToByte(rs, "standby_flag");
                        mapping.TransferableFareFlag = RecordsetHelper.ToByte(rs, "transferable_fare_flag");
                        mapping.AgencyCode = RecordsetHelper.ToString(rs, "agency_code");
                        mapping.TicketNumber = RecordsetHelper.ToString(rs, "ticket_number");
                        mapping.TicketingDateTime = RecordsetHelper.ToDateTime(rs, "ticketing_date_time");
                        mapping.TicketingBy = RecordsetHelper.ToGuid(rs, "ticketing_by");
                        mapping.CheckInSequence = RecordsetHelper.ToInt16(rs, "check_in_sequence");
                        mapping.GroupSequence = RecordsetHelper.ToInt16(rs, "group_sequence");
                        mapping.UnloadBy = RecordsetHelper.ToGuid(rs, "unload_by");
                        mapping.UnloadDateTime = RecordsetHelper.ToDateTime(rs, "unload_date_time");
                        mapping.NumberOfPieces = RecordsetHelper.ToInt16(rs, "number_of_pieces");
                        mapping.BaggageWeight = RecordsetHelper.ToDecimal(rs, "baggage_weight");
                        mapping.ExcessBaggageWeight = RecordsetHelper.ToDecimal(rs, "excess_baggage_weight");
                        mapping.CheckInBaggageWeight = RecordsetHelper.ToDecimal(rs, "check_in_baggage_weight");
                        mapping.CheckInBy = RecordsetHelper.ToGuid(rs, "check_in_by");
                        mapping.CancelBy = RecordsetHelper.ToGuid(rs, "cancel_by");
                        mapping.ExchangedDateTime = RecordsetHelper.ToDateTime(rs, "exchanged_date_time");
                        mapping.CancelDateTime = RecordsetHelper.ToDateTime(rs, "cancel_date_time");
                        mapping.RestrictionText = RecordsetHelper.ToString(rs, "restriction_text");
                        mapping.EndorsementText = RecordsetHelper.ToString(rs, "endorsement_text");
                        mapping.ExcludePricingFlag = RecordsetHelper.ToByte(rs, "exclude_pricing_flag");
                        mapping.CreateName = RecordsetHelper.ToString(rs, "create_name");
                        mapping.UpdateName = RecordsetHelper.ToString(rs, "update_name");
                        mapping.VoidDateTime = RecordsetHelper.ToDateTime(rs, "void_date_time");
                        mapping.RefundCharge = RecordsetHelper.ToDecimal(rs, "refund_charge");
                        mapping.RefundableAmount = RecordsetHelper.ToDecimal(rs, "refundable_amount");
                        mapping.RefundDateTime = RecordsetHelper.ToDateTime(rs, "refund_date_time");
                        mapping.PaymentAmount = RecordsetHelper.ToDecimal(rs, "payment_amount");
                        mapping.MappingSort = RecordsetHelper.ToInt16(rs, "mapping_sort");
                        mapping.CheckInDateTime = RecordsetHelper.ToDateTime(rs, "check_in_date_time");
                        mapping.OnwardDepartureDate = RecordsetHelper.ToDateTime(rs, "onward_departure_date");
                        mapping.ETicketStatus = RecordsetHelper.ToString(rs, "e_ticket_status");
                        mapping.HandLuggageFlag = RecordsetHelper.ToByte(rs, "hand_luggage_flag");
                        mapping.HandNumberOfPieces = RecordsetHelper.ToInt16(rs, "hand_number_of_pieces");
                        mapping.HandBaggageWeight = RecordsetHelper.ToDouble(rs, "hand_baggage_weight");
                        mapping.FreeSeatingFlag = RecordsetHelper.ToByte(rs, "free_seating_flag");
                        mapping.FareTypeRcd = RecordsetHelper.ToString(rs, "fare_type_rcd");
                        mapping.RedemptionPoints = RecordsetHelper.ToDouble(rs, "redemption_points");
                        mapping.SeatFeeRcd = RecordsetHelper.ToString(rs, "seat_fee_rcd");
                        mapping.RefundWithChargeHours = RecordsetHelper.ToInt16(rs, "refund_with_charge_hours");
                        mapping.RefundNotPossibleHours = RecordsetHelper.ToInt16(rs, "refund_not_possible_hours");
                        mapping.DutyTravelFlag = RecordsetHelper.ToByte(rs, "duty_travel_flag");
                        mapping.NotValidAfterDate = RecordsetHelper.ToDateTime(rs, "not_valid_after_date");
                        mapping.NotValidBeforeDate = RecordsetHelper.ToDateTime(rs, "not_valid_before_date");
                        mapping.AdvancedSeatingFlag = RecordsetHelper.ToByte(rs, "advanced_seating_flag");
                        mapping.FareColumn = RecordsetHelper.ToByte(rs, "fare_column");
                        #endregion

                        mapping.PieceAllowance = RecordsetHelper.ToInt16(rs, "piece_allowance");
                        mapping.BoardingTime = RecordsetHelper.ToInt16(rs, "boarding_time");
                        mapping.ItFareFlag = RecordsetHelper.ToByte(rs, "it_fare_flag");

                        mapping.ThroughFareFlag = RecordsetHelper.ToByte(rs, "through_fare_flag");
                        mapping.PassengerCheckinStatusRcd = RecordsetHelper.ToString(rs, "passenger_check_in_status_rcd");

                        mappingList.Add(mapping);

                        rs.MoveNext();
                    }

                }
            }
            catch
            {
                throw;
            }
            
            return mappingList;

        }

        public static IList<Mapping> FillBookingAllocattion(this IList<Mapping> mappingList, ADODB.Recordset rs)
        {
            try
            {
                if (rs != null && rs.RecordCount > 0)
                {
                    mappingList = new List<Mapping>();
                    Mapping mapping = null;

                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        if (RecordsetHelper.ToDecimal(rs, "payment_amount") == 0)
                        {
                            mapping = new Mapping();

                            mapping.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                            mapping.OdOriginRcd = RecordsetHelper.ToString(rs, "od_origin_rcd");
                            mapping.OdDestinationRcd = RecordsetHelper.ToString(rs, "od_destination_rcd");
                            mapping.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                            mapping.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                            mapping.NetTotal = RecordsetHelper.ToDecimal(rs, "net_total");
                            mapping.ExcludePricingFlag = RecordsetHelper.ToByte(rs, "exclude_pricing_flag");
                            mapping.PaymentAmount = RecordsetHelper.ToDecimal(rs, "payment_amount");

                            mappingList.Add(mapping);
                        }

                        rs.MoveNext();
                    }

                }
            }
            catch
            {
                throw;
            }

            return mappingList;

        }

        public static IList<Fee> FillPaymentFee(this IList<Fee> feeList, ADODB.Recordset rs, string OriginRcd, string DestinationRcd)
        {
            if (rs != null && rs.RecordCount > 0)
            {
                feeList = new List<Fee>();
                Fee fee = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        if (RecordsetHelper.ToString(rs, "od_origin_rcd") == OriginRcd && RecordsetHelper.ToString(rs, "od_destination_rcd") == DestinationRcd)
                        {
                            fee = new Fee();
                            fee.BookingFeeId = RecordsetHelper.ToGuid(rs, "booking_fee_id");
                            fee.FeeAmount = RecordsetHelper.ToDecimal(rs, "fee_amount");
                            fee.BookingId = RecordsetHelper.ToGuid(rs, "booking_id");
                            fee.PassengerId = RecordsetHelper.ToGuid(rs, "passenger_id");
                            fee.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                            fee.AcctFeeAmount = RecordsetHelper.ToDecimal(rs, "acct_fee_amount");
                            fee.FeeId = RecordsetHelper.ToGuid(rs, "fee_id");
                            fee.VatPercentage = RecordsetHelper.ToDecimal(rs, "vat_percentage");
                            fee.FeeAmountIncl = RecordsetHelper.ToDecimal(rs, "fee_amount_incl");
                            fee.AcctFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "acct_fee_amount_incl");
                            fee.FeeRcd = RecordsetHelper.ToString(rs, "fee_rcd");
                            fee.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                            fee.AccountFeeBy = RecordsetHelper.ToGuid(rs, "account_fee_by");
                            fee.AccountFeeDateTime = RecordsetHelper.ToDateTime(rs, "account_fee_date_time");
                            fee.VoidDateTime = RecordsetHelper.ToDateTime(rs, "void_date_time");
                            fee.VoidBy = RecordsetHelper.ToGuid(rs, "void_by");
                            fee.PaymentAmount = RecordsetHelper.ToDecimal(rs, "payment_amount");
                            fee.CreateBy = RecordsetHelper.ToGuid(rs, "create_by");
                            fee.CreateDateTime = RecordsetHelper.ToDateTime(rs, "create_date_time");
                            fee.UpdateBy = RecordsetHelper.ToGuid(rs, "update_by");
                            fee.UpdateDateTime = RecordsetHelper.ToDateTime(rs, "update_date_time");
                            fee.BookingSegmentId = RecordsetHelper.ToGuid(rs, "booking_segment_id");
                            fee.AgencyCode = RecordsetHelper.ToString(rs, "agency_code");
                            fee.PassengerSegmentServiceId = RecordsetHelper.ToGuid(rs, "passenger_segment_service_id");
                            fee.FeeCategoryRcd = RecordsetHelper.ToString(rs, "fee_category_rcd");
                            fee.OriginRcd = RecordsetHelper.ToString(rs, "origin_rcd");
                            fee.DestinationRcd = RecordsetHelper.ToString(rs, "destination_rcd");
                            fee.OdOriginRcd = RecordsetHelper.ToString(rs, "od_origin_rcd");
                            fee.OdDestinationRcd = RecordsetHelper.ToString(rs, "od_destination_rcd");
                            fee.NumberOfUnits = RecordsetHelper.ToDecimal(rs, "number_of_units");
                            fee.TotalAmount = RecordsetHelper.ToDecimal(rs, "total_amount");
                            fee.TotalAmountIncl = RecordsetHelper.ToDecimal(rs, "total_amount_incl");
                            fee.ManualFeeFlag = RecordsetHelper.ToByte(rs, "manual_fee_flag");
                            fee.OdFlag = RecordsetHelper.ToByte(rs, "od_flag");
                            fee.SkipFareAllowanceFlag = RecordsetHelper.ToByte(rs, "skip_fare_allowance_flag");
                            fee.FeeLevel = RecordsetHelper.ToString(rs, "fee_level");
                            fee.FeeCalculationRcd = RecordsetHelper.ToString(rs, "fee_calculation_rcd");
                            fee.MinimumFeeAmountFlag = RecordsetHelper.ToByte(rs, "minimum_fee_amount_flag");
                            fee.FeePercentage = RecordsetHelper.ToDecimal(rs, "fee_percentage");
                            fee.ChangeComment = RecordsetHelper.ToString(rs, "change_comment");
                            fee.Comment = RecordsetHelper.ToString(rs, "comment");
                            fee.TotalFeeAmount = RecordsetHelper.ToDecimal(rs, "total_fee_amount");
                            fee.TotalFeeAmountIncl = RecordsetHelper.ToDecimal(rs, "total_fee_amount_incl");
                            fee.Units = RecordsetHelper.ToString(rs, "units");
                            fee.ChargeCurrencyRcd = RecordsetHelper.ToString(rs, "charge_currency_rcd");
                            fee.ChargeAmount = RecordsetHelper.ToDecimal(rs, "charge_amount");
                            fee.ChargeAmountIncl = RecordsetHelper.ToDecimal(rs, "charge_amount_incl");
                            fee.DocumentNumber = RecordsetHelper.ToString(rs, "document_number");
                            fee.DocumentDateTime = RecordsetHelper.ToDateTime(rs, "document_date_time");
                            fee.ExternalReference = RecordsetHelper.ToString(rs, "external_reference");
                            fee.VendorRcd = RecordsetHelper.ToString(rs, "vendor_rcd");
                            fee.BaggageFeeOptionId = RecordsetHelper.ToGuid(rs, "baggage_fee_option_id");
                            fee.NewRecord = RecordsetHelper.ToBoolean(rs, "new_record");
                            fee.SelectedFee = RecordsetHelper.ToBoolean(rs, "selected_fee");

                            fee.WeightLbs = RecordsetHelper.ToDecimal(rs, "weight_lbs");
                            fee.WeightKgs = RecordsetHelper.ToDecimal(rs, "weight_kgs");
                            fee.MpdNumber = RecordsetHelper.ToString(rs, "mpd_number");
                            fee.UnitRcd = RecordsetHelper.ToString(rs, "unit_rcd");

                            feeList.Add(fee);

                            rs.MoveNext();
                        }
                    }

                }
                catch
                {
                    throw;
                }

            }
            return feeList;

        }

    }
}
