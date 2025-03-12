using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Avantik.Web.Service.Message;

namespace Avantik.Web.Service.Message.OrderBooking.Extension
{
    [MessageContract]
    public static class OrderToBookingMessage
    {
        public static IList<Message.Booking.FlightSegment> FillOrderToBookingMessage(this  IList<Message.OrderBooking.FlightSegment> objMessage, Guid bookingId, Guid userId)
        {
            IList<Avantik.Web.Service.Message.Booking.FlightSegment> segmentList = null;
            if (objMessage != null)
            {
                segmentList = new List<Avantik.Web.Service.Message.Booking.FlightSegment>();
                for (int i = 0; i < objMessage.Count; i++)
                {
                    segmentList.Add(objMessage[i].FillOrderToBookingMessage(bookingId, userId));
                }
            }
            return segmentList;
        }
        public static Avantik.Web.Service.Message.Booking.FlightSegment FillOrderToBookingMessage(this  Message.OrderBooking.FlightSegment segment, Guid bookingId, Guid userId)
        {
            Avantik.Web.Service.Message.Booking.FlightSegment flightSegment = null;

            if (segment != null)
            {
                flightSegment = new Avantik.Web.Service.Message.Booking.FlightSegment();
                flightSegment.BookingId = bookingId;

                flightSegment.BookingSegmentId = segment.booking_segment_id;
                flightSegment.FlightConnectionId = segment.flight_connection_id;

                flightSegment.DepartureDate = segment.departure_date;
                flightSegment.DepartureTime = segment.departure_time;

                flightSegment.AirlineRcd = segment.airline_rcd;
                flightSegment.FlightNumber = segment.flight_number;
                flightSegment.OriginRcd = segment.origin_rcd;
                flightSegment.DestinationRcd = segment.destination_rcd;
                flightSegment.OdOriginRcd = segment.od_origin_rcd;
                flightSegment.OdDestinationRcd = segment.od_destination_rcd;
                flightSegment.BookingClassRcd = segment.booking_class_rcd;
                flightSegment.BoardingClassRcd = segment.boarding_class_rcd;
                flightSegment.SegmentStatusRcd = segment.segment_status_rcd;

                //Use for passenger count in the flight.
                flightSegment.NumberOfUnits = segment.number_of_units;

                flightSegment.CreateBy = userId;
                flightSegment.CreateDateTime = DateTime.Now;
                flightSegment.UpdateBy = userId;
                flightSegment.UpdateDateTime = DateTime.Now;

                flightSegment.FlightId = segment.flight_id;
                flightSegment.ArrivalTime = segment.arrival_time;

                // add Arrival Date
                flightSegment.ArrivalDate = segment.arrival_date;

                flightSegment.ExchangedSegmentId = segment.exchanged_segment_id;

                flightSegment.FareId = segment.fare_id;
            }

            return flightSegment;
        }

        public static IList<Message.Booking.Mapping> FillOrderToBookingMessage(this  IList<Message.OrderBooking.Mapping> objMessage, Guid booking_id, Guid userId, string agencyCode)
        {
            IList<Avantik.Web.Service.Message.Booking.Mapping> mappingtList = null;
            if (objMessage != null)
            {
                mappingtList = new List<Avantik.Web.Service.Message.Booking.Mapping>();
                for (int i = 0; i < objMessage.Count; i++)
                {
                    mappingtList.Add(objMessage[i].FillOrderToBookingMessage(booking_id, userId, agencyCode));
                }
            }
            return mappingtList;
        }
        public static Message.Booking.Mapping FillOrderToBookingMessage(this  Message.OrderBooking.Mapping m, Guid booking_id, Guid userId, string agencyCode)
        {
            Avantik.Web.Service.Message.Booking.Mapping mapping = null;

            if (m != null)
            {
                mapping = new Avantik.Web.Service.Message.Booking.Mapping();

                mapping.AgencyCode = agencyCode;
                mapping.BookingId = booking_id;
                mapping.BookingSegmentId = m.booking_segment_id;
                mapping.PassengerId = m.passenger_id;

                mapping.NotValidBeforeDate = m.not_valid_before_date;
                mapping.NotValidAfterDate = m.not_valid_after_date;

                mapping.RefundWithChargeHours = m.refund_with_charge_hours;
                mapping.RefundNotPossibleHours = m.refund_not_possible_hours;

                mapping.PieceAllowance = m.piece_allowance;

                mapping.FareCode = m.fare_code;
                if (string.IsNullOrEmpty(m.seat_number) == false && m.seat_number.Length > 1)
                {
                    mapping.SeatColumn = m.seat_number.Substring(m.seat_number.Length - 1, 1);
                    mapping.SeatRow = Convert.ToInt16(m.seat_number.Substring(0, m.seat_number.Length - 1));
                }
                mapping.SeatNumber = m.seat_number;
                mapping.CurrencyRcd = m.currency_rcd;
                mapping.EndorsementText = m.endorsement_text;
                mapping.RestrictionText = m.restriction_text;

                mapping.FareAmount = m.fare_amount;
                mapping.FareAmountIncl = m.fare_amount_incl;
                mapping.FareVat = m.fare_vat;
                mapping.NetTotal = m.net_total;
                //Optinal
                mapping.CommissionAmount = m.commission_amount;
                //Optinal
                mapping.CommissionAmountIncl = m.commission_amount_incl;
                //Optinal
                mapping.CommissionPercentage = m.commission_percentage;
                //Optinal
                mapping.PublicFareAmount = m.public_fare_amount;
                //Optinal
                mapping.PublicFareAmountIncl = m.public_fare_amount_incl;
                mapping.RefundCharge = m.refund_charge;
                mapping.BaggageWeight = m.baggage_weight;
                mapping.RedemptionPoints = m.redemption_points;

                mapping.ETicketFlag = m.e_ticket_flag;
                mapping.RefundableFlag = m.refundable_flag;
                mapping.TransferableFareFlag = m.transferable_fare_flag;
                mapping.ThroughFareFlag = m.through_fare_flag;
                mapping.ItFareFlag = m.it_fare_flag;
                mapping.DutyTravelFlag = m.duty_travel_flag;
                mapping.StandbyFlag = m.standby_flag;
                mapping.ExcludePricingFlag = m.exclude_pricing_flag;

                mapping.CreateBy = userId;
                mapping.CreateDateTime = DateTime.Now;
                mapping.UpdateBy = userId;
                mapping.UpdateDateTime = DateTime.Now;
            }

            return mapping;
        }


        public static IList<Message.Booking.Fee> FillOrderToBookingMessage(this  IList<Message.OrderBooking.Fee> objMessage, Guid bookingId, Guid userId, string agencyCode)
        {
            IList<Avantik.Web.Service.Message.Booking.Fee> feeList = null;
            if (objMessage != null)
            {
                feeList = new List<Message.Booking.Fee>();
                for (int i = 0; i < objMessage.Count; i++)
                {
                    feeList.Add(objMessage[i].FillOrderToBookingMessage(bookingId, userId, agencyCode));
                }
            }
            return feeList;
        }
        public static Message.Booking.Fee FillOrderToBookingMessage(this  Message.OrderBooking.Fee f, Guid bookingId, Guid userId, string agencyCode)
        {
            Avantik.Web.Service.Message.Booking.Fee fee = null;

            if (f != null)
            {
                fee = new Avantik.Web.Service.Message.Booking.Fee();

                fee.BookingId = bookingId;

                fee.PassengerId = f.passenger_id;
                fee.BookingSegmentId = f.booking_segment_id;
                fee.PassengerSegmentServiceId = f.passenger_segment_service_id;

                if (f.fee_id != Guid.Empty)
                    fee.BookingFeeId = f.fee_id;//Guid.NewGuid();
                else
                    fee.BookingFeeId = Guid.NewGuid();

                // Remove assign FeeId

                // for 2020 
                //  fee.FeeId = f.fee_id;

                fee.FeeRcd = f.fee_rcd;
                fee.FeeCategoryRcd = f.fee_category_rcd;
                fee.VendorRcd = f.vendor_rcd;
                fee.OdOriginRcd = f.od_origin_rcd;
                fee.OdDestinationRcd = f.od_destination_rcd;
                fee.CurrencyRcd = f.currency_rcd;
                fee.ChargeCurrencyRcd = f.charge_currency_rcd;
                fee.UnitRcd = f.unit_rcd;
                fee.MpdNumber = f.mpd_number;
                fee.ChangeComment = f.comment;
                fee.ExternalReference = f.external_reference;

                fee.FeeAmount = f.fee_amount;
                fee.FeeAmountIncl = f.fee_amount_incl;
                fee.VatPercentage = f.vat_percentage;
                fee.ChargeAmount = f.charge_amount;
                fee.ChargeAmountIncl = f.charge_amount_incl;
                fee.WeightLbs = f.weight_lbs;
                fee.WeightKgs = f.weight_kgs;
                fee.NumberOfUnits = f.number_of_units;

                fee.CreateBy = userId;
                fee.CreateDateTime = DateTime.Now;
                fee.UpdateBy = userId;
                fee.UpdateDateTime = DateTime.Now;

                fee.AgencyCode = agencyCode;

            }

            return fee;
        }

        public static IList<Message.Booking.PassengerService> FillOrderToBookingMessage(this  IList<Message.OrderBooking.Service> objMessage, Guid userId)
        {
            IList<Avantik.Web.Service.Message.Booking.PassengerService> serviceList = null;
            if (objMessage != null)
            {
                serviceList = new List<Avantik.Web.Service.Message.Booking.PassengerService>();
                for (int i = 0; i < objMessage.Count; i++)
                {
                    serviceList.Add(objMessage[i].FillOrderToBookingMessage(userId));
                }
            }
            return serviceList;
        }
        public static Message.Booking.PassengerService FillOrderToBookingMessage(this Message.OrderBooking.Service ser, Guid userId)
        {
            Avantik.Web.Service.Message.Booking.PassengerService service = null;

            if (ser != null)
            {
                service = new Avantik.Web.Service.Message.Booking.PassengerService();

                if(ser == null)
                    service.PassengerSegmentServiceId = Guid.NewGuid();

                service.PassengerId = ser.passenger_id;
                service.BookingSegmentId = ser.booking_segment_id;

                service.NumberOfUnits = ser.number_of_units;

                service.SpecialServiceRcd = ser.special_service_rcd;
                service.ServiceText = ser.service_text;
                //The value should be SS or NN
                service.SpecialServiceStatusRcd = "NN";
                service.UnitRcd = ser.unit_rcd;

                service.CreateBy = userId;
                service.CreateDateTime = DateTime.Now;
                service.UpdateBy = userId;
                service.UpdateDateTime = DateTime.Now;


            }

            return service;
        }


        public static IList<Message.Booking.Tax> FillOrderToBookingMessage(this  IList<Message.OrderBooking.Tax> objMessage, Guid userId)
        {
            IList<Avantik.Web.Service.Message.Booking.Tax> taxList = null;
            if (objMessage != null)
            {
                taxList = new List<Avantik.Web.Service.Message.Booking.Tax>();
                for (int i = 0; i < objMessage.Count; i++)
                {
                    taxList.Add(objMessage[i].FillOrderToBookingMessage(userId));
                }
            }
            return taxList;
        }

        //// add tax to mapping
        public static IList<Avantik.Web.Service.Message.Booking.Mapping> FillMapping(this IList<Tax> taxList, IList<Avantik.Web.Service.Message.Booking.Mapping> m)
        {

            if (taxList != null && m != null)
            {
                for (int i = 0; i < taxList.Count; i++)
                {
                    for (int j = 0; j < m.Count; j++)
                    {
                        if (taxList[i].booking_segment_id == m[j].BookingSegmentId && taxList[i].passenger_id == m[j].PassengerId)
                            taxList[i].FillMapping(m[j]);
                    }
                }
            }
            return m;
        }
        public static Avantik.Web.Service.Message.Booking.Mapping FillMapping(this Tax tax, Avantik.Web.Service.Message.Booking.Mapping mapping)
        {
            if (tax != null && mapping != null)
            {

                if (tax.tax_rcd != null && tax.tax_rcd.ToUpper() == "YQ")
                {
                    mapping.YqAmount += tax.tax_amount;
                    mapping.YqAmountIncl += tax.tax_amount_incl;

                    mapping.AcctYqAmount += tax.tax_amount;
                    mapping.AcctYqAmountIncl += tax.tax_amount_incl;

                    mapping.YqVat += tax.tax_amount_incl - tax.tax_amount;
                }
                else
                {
                    mapping.TaxAmount += tax.tax_amount;
                    mapping.TaxAmountIncl += tax.tax_amount_incl;
                    mapping.TaxVat += tax.tax_amount_incl - tax.tax_amount;

                    mapping.AcctTaxAmount += tax.tax_amount;
                    mapping.AcctTaxAmountIncl += tax.tax_amount_incl;
                }

            }

            return mapping;
        }


        public static Message.Booking.Tax FillOrderToBookingMessage(this Message.OrderBooking.Tax tx, Guid userId)
        {
            Avantik.Web.Service.Message.Booking.Tax tax = null;

            if (tx != null)
            {
                tax = new Avantik.Web.Service.Message.Booking.Tax();

                tax.PassengerId = tx.passenger_id;
                tax.BookingSegmentId = tx.booking_segment_id;

                tax.TaxRcd = tx.tax_rcd;

                tax.TaxAmount = tx.tax_amount;
                tax.TaxAmountIncl = tx.tax_amount_incl;
                tax.TaxCurrencyRcd = tx.tax_currency_rcd;
                tax.SalesAmount = tx.sales_amount;
                tax.SalesAmountIncl = tx.sales_amount_incl;
                tax.SalesCurrencyRcd = tx.sales_currency_rcd;

                tax.CreateBy = userId;
                tax.CreateDateTime = DateTime.Now;
                tax.UpdateBy = userId;
                tax.UpdateDateTime = DateTime.Now;


            }

            return tax;
        }

    }
}
