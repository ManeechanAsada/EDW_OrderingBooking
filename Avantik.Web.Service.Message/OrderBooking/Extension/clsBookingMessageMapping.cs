using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking.Extension
{
    [MessageContract]
    public static class BookingMessageMapping
    {
        public static BookingHeaderResponse MapBookingHeaderResponse(this Message.Booking.BookingHeader header)
        {
            try
            {
                BookingHeaderResponse response = new BookingHeaderResponse();

                   // response.PhoneMobile = header.PhoneMobile;
                   // response.PhoneHome = header.PhoneHome;
                   // response.PhoneBusiness = header.PhoneBusiness;
                   // response.PhoneFax = header.PhoneFax;

                   // response.AddressLine1 = header.AddressLine1;
                   // response.AddressLine2 = header.AddressLine2;
                   // response.Street = header.Street;
                   // response.PoBox = header.PoBox;
                   // response.City = header.City;
                   // response.State = header.State;
                   // response.District = header.District;
                   // response.Province = header.Province;
                   // response.ZipCode = header.ZipCode;
                   // response.CountryRcd = header.CountryRcd;

                   // response.AgencyCode = header.AgencyCode;
                   // response.CurrencyRcd = header.CurrencyRcd;
                   // response.BookingNumber = header.BookingNumber;
                   // response.LanguageRcd = header.LanguageRcd;
                   // response.ContactName = header.ContactName;
                   // response.ContactEmail = header.ContactEmail;
                   // response.ReceivedFrom = header.ReceivedFrom;
                   // response.PhoneSearch = header.PhoneSearch;
                   // response.Comment = header.Comment;
                   // response.TitleRcd = header.TitleRcd;
                   // response.Lastname = header.Lastname;
                   // response.Firstname = header.Firstname;
                   // response.Middlename = header.Middlename;
                   //// response.CountryRcd = header.CountryRcd;
                   // response.GroupBookingFlag = header.GroupBookingFlag;

                response.booking_id = header.BookingId;
                response.record_locator = header.RecordLocator;
                response.agency_code = header.AgencyCode;
                response.currency_rcd = header.CurrencyRcd;
                response.booking_number = Convert.ToString(header.BookingNumber);
                response.language_rcd = header.LanguageRcd;
                response.contact_name = header.ContactName;
                response.contact_email = header.ContactEmail;
                response.received_from = header.ReceivedFrom;
                response.phone_search = header.PhoneSearch;
                response.comment = header.Comment;
                response.title_rcd = header.TitleRcd;
                response.lastname = header.Lastname;
                response.firstname = header.Firstname;
                response.middlename = header.Middlename;
                response.country_rcd = header.CountryRcd;
                response.group_booking_flag = header.GroupBookingFlag;

                response.address = new Address();
                response.address.address_line1 = header.AddressLine1;
                response.address.address_line2 = header.AddressLine2;
                response.address.street = header.Street;
                response.address.po_box = header.PoBox;
                response.address.city = header.City;
                response.address.state = header.State;
                response.address.district = header.District;
                response.address.province = header.Province;
                response.address.zip_code = header.ZipCode;
                response.address.country_rcd = header.CountryRcd;

                response.telephone = new Telephone();
                response.telephone.phone_mobile = header.PhoneMobile;
                response.telephone.phone_home = header.PhoneHome;
                response.telephone.phone_business = header.PhoneBusiness;
                response.telephone.phone_fax = header.PhoneFax;

                return response;
            }
            catch
            {
                throw;
            }
        }

        public static IList<FlightSegmentResponse> MapBookingSegmentsResponse(this  IList<Message.Booking.FlightSegment> objBooking)
        {
            IList<FlightSegmentResponse> objResponseList = new List<FlightSegmentResponse>();
            if (objBooking != null)
            {
                for (int i = 0; i < objBooking.Count; i++)
                {
                    objResponseList.Add(objBooking[i].MapBookingSegmentsResponse());
                }
            }
            return objResponseList;
        }
        public static FlightSegmentResponse MapBookingSegmentsResponse(this  Avantik.Web.Service.Message.Booking.FlightSegment objBookingResponse)
        {
            FlightSegmentResponse objReadResponse = new FlightSegmentResponse();

            if (objBookingResponse != null)
            {
                objReadResponse.booking_segment_id = objBookingResponse.BookingSegmentId;
                objReadResponse.flight_connection_id = objBookingResponse.FlightConnectionId;
                objReadResponse.flight_id = objBookingResponse.FlightId;

                objReadResponse.departure_date = objBookingResponse.DepartureDate;
                objReadResponse.departure_time = objBookingResponse.DepartureTime;
                objReadResponse.arrival_time = objBookingResponse.ArrivalTime;

                // add Arrival Date
                objReadResponse.arrival_date = objBookingResponse.ArrivalDate;

                objReadResponse.airline_rcd = objBookingResponse.AirlineRcd;
                objReadResponse.flight_number = objBookingResponse.FlightNumber;
                objReadResponse.origin_rcd = objBookingResponse.OriginRcd;
                objReadResponse.destination_rcd = objBookingResponse.DestinationRcd;
                objReadResponse.od_origin_rcd = objBookingResponse.OdOriginRcd;
                objReadResponse.od_destination_rcd = objBookingResponse.OdDestinationRcd;
                objReadResponse.booking_class_rcd = objBookingResponse.BookingClassRcd;
                objReadResponse.boarding_class_rcd = objBookingResponse.BoardingClassRcd;
                objReadResponse.segment_status_rcd = objBookingResponse.SegmentStatusRcd;

                //Use for passenger count in the flight.
                objReadResponse.number_of_units = objBookingResponse.NumberOfUnits;

            }

            return objReadResponse;
        }

        public static IList<PassengerResponse> MapPassengersResponse(this  IList<Avantik.Web.Service.Message.Booking.Passenger> objBooking)
        {
            IList<PassengerResponse> objResponseList = new List<PassengerResponse>();
            if (objBooking != null)
            {
                for (int i = 0; i < objBooking.Count; i++)
                {
                    objResponseList.Add(objBooking[i].MapPassengersResponse());
                }
            }
            return objResponseList;
        }
        public static PassengerResponse MapPassengersResponse(this  Avantik.Web.Service.Message.Booking.Passenger objBookingResponse)
        {
            PassengerResponse objReadResponse = new PassengerResponse();

            if (objBookingResponse != null)
            {
                objReadResponse.passenger_id = objBookingResponse.PassengerId;
                //optional use only when want to defind guardian of passenger type CHD
                objReadResponse.guardian_passenger_id = objBookingResponse.GuardianPassengerId;

                objReadResponse.date_of_birth = objBookingResponse.DateOfBirth;

                objReadResponse.passenger_type_rcd = objBookingResponse.PassengerTypeRcd;
                objReadResponse.lastname = objBookingResponse.Lastname;
                objReadResponse.firstname = objBookingResponse.Firstname;
                objReadResponse.middlename = objBookingResponse.Middlename;
                objReadResponse.title_rcd = objBookingResponse.TitleRcd;


            }

            return objReadResponse;
        }

        public static IList<MappingResponse> MapMappingsResponse(this  IList<Avantik.Web.Service.Message.Booking.Mapping> objBooking)
        {
            IList<MappingResponse> objResponseList = new List<MappingResponse>();
            if (objBooking != null)
            {
                for (int i = 0; i < objBooking.Count; i++)
                {
                    objResponseList.Add(objBooking[i].MapMappingsResponse());
                }
            }
            return objResponseList;
        }
        public static MappingResponse MapMappingsResponse(this  Avantik.Web.Service.Message.Booking.Mapping objBookingResponse)
        {
            MappingResponse objReadResponse = new MappingResponse();

            if (objBookingResponse != null)
            {
                objReadResponse.passenger_id = objBookingResponse.PassengerId;
                objReadResponse.booking_segment_id = objBookingResponse.BookingSegmentId;

                objReadResponse.not_valid_before_date = objBookingResponse.NotValidBeforeDate;
                objReadResponse.not_valid_after_date = objBookingResponse.NotValidAfterDate;

                objReadResponse.refund_with_charge_hours = objBookingResponse.RefundWithChargeHours;
                objReadResponse.refund_not_possible_hours = objBookingResponse.RefundNotPossibleHours;

                objReadResponse.piece_allowance = objBookingResponse.PieceAllowance;

                objReadResponse.fare_code = objBookingResponse.FareCode;
                objReadResponse.seat_number = objBookingResponse.SeatNumber;
                objReadResponse.currency_rcd = objBookingResponse.CurrencyRcd;
                objReadResponse.endorsement_text = objBookingResponse.EndorsementText;
                objReadResponse.restriction_text = objBookingResponse.RestrictionText;

                objReadResponse.fare_amount = objBookingResponse.FareAmount;
                objReadResponse.fare_amount_incl = objBookingResponse.FareAmountIncl;
                objReadResponse.fare_vat = objBookingResponse.FareVat;
                objReadResponse.net_total = objBookingResponse.NetTotal;
                //Optinal
                objReadResponse.commission_amount = objBookingResponse.CommissionAmount;
                //Optinal
                objReadResponse.commission_amount_incl = objBookingResponse.CommissionAmountIncl;
                //Optinal
                objReadResponse.commission_percentage = objBookingResponse.CommissionPercentage;
                //Optinal
                objReadResponse.public_fare_amount = objBookingResponse.PublicFareAmount;
                //Optinal
                objReadResponse.public_fare_amount_incl = objBookingResponse.PublicFareAmountIncl;
                objReadResponse.refund_charge = objBookingResponse.RefundCharge;
                objReadResponse.baggage_weight = objBookingResponse.BaggageWeight;
                objReadResponse.redemption_points = objBookingResponse.RedemptionPoints;

                objReadResponse.e_ticket_flag = objBookingResponse.ETicketFlag;
                objReadResponse.refundable_flag = objBookingResponse.RefundableFlag;
                objReadResponse.transferable_fare_flag = objBookingResponse.TransferableFareFlag;
                objReadResponse.through_fare_flag = objBookingResponse.ThroughFareFlag;
                objReadResponse.it_fare_flag = objBookingResponse.ItFareFlag;
                objReadResponse.duty_travel_flag = objBookingResponse.DutyTravelFlag;
                objReadResponse.standby_flag = objBookingResponse.StandbyFlag;
                objReadResponse.exclude_pricing_flag = objBookingResponse.ExcludePricingFlag;

            }

            return objReadResponse;
        }

        public static IList<FeeResponse> MapFeesResponse(this  IList<Avantik.Web.Service.Message.Booking.Fee> objBooking)
        {
            IList<FeeResponse> objResponseList = new List<FeeResponse>();
            if (objBooking != null)
            {
                for (int i = 0; i < objBooking.Count; i++)
                {
                    objResponseList.Add(objBooking[i].MapFeesResponse());
                }
            }
            return objResponseList;
        }
        public static FeeResponse MapFeesResponse(this  Avantik.Web.Service.Message.Booking.Fee objBookingResponse)
        {
            FeeResponse objReadResponse = new FeeResponse();

            if (objBookingResponse != null)
            {
                objReadResponse.passenger_id = objBookingResponse.PassengerId;
                objReadResponse.booking_segment_id = objBookingResponse.BookingSegmentId;
                objReadResponse.passenger_segment_service_id = objBookingResponse.PassengerSegmentServiceId;

                objReadResponse.fee_id = objBookingResponse.BookingFeeId;

                objReadResponse.fee_rcd = objBookingResponse.FeeRcd;
                objReadResponse.vendor_rcd = objBookingResponse.VendorRcd;
                objReadResponse.od_origin_rcd = objBookingResponse.OdOriginRcd;
                objReadResponse.od_destination_rcd = objBookingResponse.OdDestinationRcd;
                objReadResponse.currency_rcd = objBookingResponse.CurrencyRcd;
                objReadResponse.charge_currency_rcd = objBookingResponse.ChargeCurrencyRcd;
                objReadResponse.unit_rcd = objBookingResponse.UnitRcd;
                objReadResponse.mpd_number = objBookingResponse.MpdNumber;
                objReadResponse.comment = objBookingResponse.Comment;

                objReadResponse.external_reference = objBookingResponse.ExternalReference;

                objReadResponse.fee_amount = objBookingResponse.FeeAmount;
                objReadResponse.fee_amount_incl = objBookingResponse.FeeAmountIncl;
                objReadResponse.vat_percentage = objBookingResponse.VatPercentage;
                objReadResponse.charge_amount = objBookingResponse.ChargeAmount;
                objReadResponse.charge_amount_incl = objBookingResponse.ChargeAmountIncl;
                objReadResponse.weight_lbs = objBookingResponse.WeightLbs;
                objReadResponse.weight_kgs = objBookingResponse.WeightKgs;
                objReadResponse.number_of_units = objBookingResponse.NumberOfUnits;

            }

            return objReadResponse;
        }
        public static IList<TaxResponse> MapTaxsResponse(this  IList<Avantik.Web.Service.Message.Booking.Tax> objBooking)
        {
            IList<TaxResponse> objResponseList = new List<TaxResponse>();
            if (objBooking != null)
            {
                for (int i = 0; i < objBooking.Count; i++)
                {
                    objResponseList.Add(objBooking[i].MapTaxsResponse());
                }
            }
            return objResponseList;
        }
        public static TaxResponse MapTaxsResponse(this  Avantik.Web.Service.Message.Booking.Tax objBookingResponse)
        {
            TaxResponse objReadResponse = new TaxResponse();

            if (objBookingResponse != null)
            {
                objReadResponse.passenger_id = objBookingResponse.PassengerId;
                objReadResponse.booking_segment_id = objBookingResponse.BookingSegmentId;

                objReadResponse.tax_rcd = objBookingResponse.TaxRcd;
                objReadResponse.tax_currency_rcd = objBookingResponse.TaxCurrencyRcd;
                objReadResponse.sales_currency_rcd = objBookingResponse.SalesCurrencyRcd;

                objReadResponse.sales_amount = objBookingResponse.SalesAmount;
                objReadResponse.sales_amount_incl = objBookingResponse.SalesAmountIncl;
                objReadResponse.tax_amount = objBookingResponse.TaxAmount;
                objReadResponse.tax_amount_incl = objBookingResponse.TaxAmountIncl;
            }

            return objReadResponse;
        }

        public static IList<RemarkResponse> MapRemarksResponse(this  IList<Avantik.Web.Service.Message.Booking.Remark> objBooking)
        {
            IList<RemarkResponse> objResponseList = new List<RemarkResponse>();
            if (objBooking != null)
            {
                for (int i = 0; i < objBooking.Count; i++)
                {
                    objResponseList.Add(objBooking[i].MapRemarksResponse());
                }
            }
            return objResponseList;
        }
        public static RemarkResponse MapRemarksResponse(this  Avantik.Web.Service.Message.Booking.Remark objBookingResponse)
        {
            RemarkResponse objReadResponse = new RemarkResponse();

            if (objBookingResponse != null)
            {
                objReadResponse.timelimit_date_time = objBookingResponse.TimelimitDateTime;
                objReadResponse.utc_timelimit_date_time = objBookingResponse.UtcTimelimitDateTime;

                objReadResponse.remark_type_rcd = objBookingResponse.RemarkTypeRcd;
                objReadResponse.remark_text = objBookingResponse.RemarkText;
                objReadResponse.nickname = objBookingResponse.Nickname;
            }

            return objReadResponse;
        }

        public static IList<ServiceResponse> MapServicesResponse(this  IList<Avantik.Web.Service.Message.Booking.PassengerService> objBooking)
        {
            IList<ServiceResponse> objResponseList = new List<ServiceResponse>();
            if (objBooking != null)
            {
                for (int i = 0; i < objBooking.Count; i++)
                {
                    objResponseList.Add(objBooking[i].MapServicesResponse());
                }
            }
            return objResponseList;
        }
        public static ServiceResponse MapServicesResponse(this  Avantik.Web.Service.Message.Booking.PassengerService objBookingResponse)
        {
            ServiceResponse objReadResponse = new ServiceResponse();

            if (objBookingResponse != null)
            {
               // objReadResponse.p = objBookingResponse.PassengerSegmentServiceId;
                objReadResponse.passenger_id = objBookingResponse.PassengerId;
                objReadResponse.booking_segment_id = objBookingResponse.BookingSegmentId;

                objReadResponse.number_of_units = objBookingResponse.NumberOfUnits;

                objReadResponse.special_service_rcd = objBookingResponse.SpecialServiceRcd;
                objReadResponse.service_text = objBookingResponse.ServiceText;
                //The value should be SS or NN
                objReadResponse.special_service_rcd = objBookingResponse.SpecialServiceStatusRcd;
                objReadResponse.unit_rcd = objBookingResponse.UnitRcd;

            }

            return objReadResponse;
        }

        public static IList<PaymentResponse> MapPaymentsResponse(this  IList<Avantik.Web.Service.Message.Booking.Payment> objBooking)
        {
            IList<PaymentResponse> objResponseList = new List<PaymentResponse>();
            if (objBooking != null)
            {
                for (int i = 0; i < objBooking.Count; i++)
                {
                    objResponseList.Add(objBooking[i].MapPaymentsResponse());
                }
            }
            return objResponseList;
        }
        public static PaymentResponse MapPaymentsResponse(this  Avantik.Web.Service.Message.Booking.Payment objBookingResponse)
        {
            PaymentResponse objReadResponse = new PaymentResponse();

            if (objBookingResponse != null)
            {
                objReadResponse.payment_date_time = objBookingResponse.PaymentDateTime;

                objReadResponse.payment_amount = objBookingResponse.PaymentAmount;
                objReadResponse.receive_payment_amount = objBookingResponse.ReceivePaymentAmount;

                //Value will be SALE or REFUND
                objReadResponse.payment_type_rcd = objBookingResponse.PaymentTypeRcd;
                objReadResponse.form_of_payment_rcd = objBookingResponse.FormOfPaymentRcd;
                objReadResponse.form_of_payment_subtype_rcd = objBookingResponse.FormOfPaymentSubtypeRcd;
                objReadResponse.debit_agency_code = objBookingResponse.DebitAgencyCode;
                objReadResponse.currency_rcd = objBookingResponse.CurrencyRcd;
                objReadResponse.receive_currency_rcd = objBookingResponse.ReceiveCurrencyRcd;
                objReadResponse.approval_code = objBookingResponse.ApprovalCode;
                objReadResponse.transaction_reference = objBookingResponse.TransactionReference;
                objReadResponse.payment_number = objBookingResponse.PaymentNumber;
                objReadResponse.payment_reference = objBookingResponse.PaymentReference;
                objReadResponse.payment_remark = objBookingResponse.PaymentRemark;

                // credit card
                //  objReadResponse.credit_card = new CreditCard();
                //objReadResponse.c.DocumentNumber = objBookingResponse.DocumentNumber;
                //objReadResponse.NameOnCard = objBookingResponse.NameOnCard;
                //objReadResponse.CvvCode = objBookingResponse.CvvCode;
                //objReadResponse.IssueNumber = objBookingResponse.IssueNumber;

                //objReadResponse.ExpiryMonth = objBookingResponse.ExpiryMonth;
                //objReadResponse.ExpiryYear = objBookingResponse.ExpiryYear;
                //objReadResponse.IssueMonth = objBookingResponse.IssueMonth;
                //objReadResponse.IssueYear = objBookingResponse.IssueYear;

                // address
                //  objReadResponse.credit_card.Address = new Message.Address();
                //objReadResponse. = objBookingResponse.AddressLine1;
                //objReadResponse.AddressLine2 = objBookingResponse.AddressLine2;
                //objReadResponse.Street = objBookingResponse.Street;
                //objReadResponse.PoBox = objBookingResponse.PoBox;
                //objReadResponse.City = objBookingResponse.City;
                //objReadResponse.State = objBookingResponse.State;
                //objReadResponse.District = objBookingResponse.District;
                //objReadResponse.Province = objBookingResponse.Province;
                //objReadResponse.ZipCode = objBookingResponse.ZipCode;
                //objReadResponse.CountryRcd = objBookingResponse.CountryRcd;
            }

            return objReadResponse;
        }



    }
}
