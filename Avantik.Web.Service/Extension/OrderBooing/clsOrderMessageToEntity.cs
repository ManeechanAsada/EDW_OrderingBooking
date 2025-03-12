using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Booking;

namespace Avantik.Web.Service.Extension.OrderBooking
{
    public static class OrderMessageToEntity
    {
        public static IList<Entity.Booking.FlightSegment> ToListEntity(this  IList<Message.Booking.FlightSegment> objOrder)
        {
            IList<Entity.Booking.FlightSegment> objEntityList = null;
            if (objOrder != null)
            {
                objEntityList = new List<Entity.Booking.FlightSegment>();

                for (int i = 0; i < objOrder.Count; i++)
                {
                    objEntityList.Add(objOrder[i].ToEntity());
                }
            }

            return objEntityList;
        }
        public static Entity.Booking.FlightSegment ToEntity(this Message.Booking.FlightSegment objMessage)
        {
            Entity.Booking.FlightSegment objEntity = null;
            if (objMessage != null)
            {
                objEntity = new Entity.Booking.FlightSegment();
                //objEntity.BookingSegmentId = objMessage.booking_segment_id;
                //objEntity.SegmentStatusRcd = objMessage.segment_status_rcd;
                //objEntity.BoardingClassRcd = objMessage.boarding_class_rcd;
                //objEntity.BookingClassRcd = objMessage.booking_class_rcd;
                //objEntity.DepartureDate = objMessage.departure_date;
                //objEntity.DestinationRcd = objMessage.destination_rcd;
                //objEntity.FlightConnectionId = objMessage.flight_connection_id;
                //objEntity.FlightId = objMessage.flight_id;
                //objEntity.OdDestinationRcd = objMessage.od_destination_rcd;
                //objEntity.OdOriginRcd = objMessage.od_origin_rcd;
                //objEntity.OriginRcd = objMessage.origin_rcd;
                //objEntity.NumberOfUnits = objMessage.number_of_units;
                //objEntity.AirlineRcd = objMessage.airline_rcd;
                //objEntity.FlightNumber = objMessage.flight_number;

                //objEntity.ArrivalTime = objMessage.arrival_time;
                //objEntity.DepartureTime = objMessage.departure_time;

            }
            return objEntity;
        }

        public static IList<Entity.Booking.Fee> ToListEntity(this  IList<Message.Booking.Fee> objOrder)
        {
            IList<Entity.Booking.Fee> objEntityList = null;
            if (objOrder != null)
            {
                objEntityList = new List<Entity.Booking.Fee>();

                for (int i = 0; i < objOrder.Count; i++)
                {
                    objEntityList.Add(objOrder[i].ToEntity());
                }
            }

            return objEntityList;
        }
        public static Entity.Booking.Fee ToEntity(this Message.Booking.Fee objMessage)
        {
            Entity.Booking.Fee objEntity = null;
            if (objMessage != null)
            {
                objEntity = new Entity.Booking.Fee();
                //objEntity.BookingSegmentId = objMessage.booking_segment_id;
                //objEntity.PassengerId = objMessage.passenger_id;
                //objEntity.PassengerSegmentServiceId = objMessage.passenger_segment_service_id;
                //objEntity.FeeId = objMessage.fee_id;
                //objEntity.FeeRcd = objMessage.fee_rcd;
                //objEntity.VendorRcd = objMessage.vendor_rcd;
                //objEntity.OdOriginRcd = objMessage.od_origin_rcd;
                //objEntity.OdDestinationRcd = objMessage.od_destination_rcd;
                //objEntity.CurrencyRcd = objMessage.currency_rcd;
                //objEntity.ChargeCurrencyRcd = objMessage.charge_currency_rcd;
                //objEntity.UnitRcd = objMessage.unit_rcd;
                //objEntity.MpdNumber = objMessage.mpd_number;
                //objEntity.Comment = objMessage.comment;
                //objEntity.ExternalReference = objMessage.external_reference;
                //objEntity.FeeAmount = objMessage.fee_amount;
                //objEntity.FeeAmountIncl = objMessage.fee_amount_incl;
                //objEntity.VatPercentage = objMessage.vat_percentage;
                //objEntity.ChargeAmount = objMessage.charge_amount;
                //objEntity.ChargeAmountIncl = objMessage.charge_amount_incl;
                //objEntity.WeightKgs = objMessage.weight_kgs;
                //objEntity.WeightLbs = objMessage.weight_lbs;
                //objEntity.NumberOfUnits = objMessage.number_of_units;


            }
            return objEntity;
        }

        public static IList<Entity.Booking.Mapping> ToListEntity(this  IList<Message.Booking.Mapping> objOrder)
        {
            IList<Entity.Booking.Mapping> objEntityList = null;
            if (objOrder != null)
            {
                objEntityList = new List<Entity.Booking.Mapping>();

                for (int i = 0; i < objOrder.Count; i++)
                {
                    objEntityList.Add(objOrder[i].ToEntity());
                }
            }

            return objEntityList;
        }
        public static Entity.Booking.Mapping ToEntity(this Message.Booking.Mapping objMessage)
        {
            Entity.Booking.Mapping objEntity = null;
            if (objMessage != null)
            {
                objEntity = new Entity.Booking.Mapping();
                //objEntity.BookingSegmentId = objMessage.booking_segment_id;
                //objEntity.PassengerId = objMessage.passenger_id;
                //objEntity.FareCode = objMessage.fare_code;
                //objEntity.SeatNumber = objMessage.seat_number;
                //objEntity.CurrencyRcd = objMessage.currency_rcd;
                //objEntity.EndorsementText = objMessage.endorsement_text;
                //objEntity.RestrictionText = objMessage.restriction_text;
                //objEntity.FareAmount = objMessage.fare_amount;
                //objEntity.FareAmountIncl = objMessage.fare_amount_incl;
                //objEntity.FareVat = objMessage.fare_vat;
                //objEntity.NetTotal = objMessage.net_total;
                //objEntity.ETicketFlag = objMessage.e_ticket_flag;
                //objEntity.RefundableFlag = objMessage.refundable_flag;
                //objEntity.TransferableFareFlag = objMessage.transferable_fare_flag;
                //objEntity.ThroughFareFlag = objMessage.through_fare_flag;
                //objEntity.ItFareFlag = objMessage.it_fare_flag;
                //objEntity.DutyTravelFlag = objMessage.duty_travel_flag;
                //objEntity.StandbyFlag = objMessage.standby_flag;
                //objEntity.ExcludePricingFlag = objMessage.exclude_pricing_flag;

            }
            return objEntity;
        }


        public static IList<Entity.Booking.PassengerService> ToListEntity(this  IList<Message.Booking.PassengerService> objOrder)
        {
            IList<Entity.Booking.PassengerService> objEntityList = null;
            if (objOrder != null)
            {
                objEntityList = new List<Entity.Booking.PassengerService>();

                for (int i = 0; i < objOrder.Count; i++)
                {
                    objEntityList.Add(objOrder[i].ToEntity());
                }
            }

            return objEntityList;
        }
        public static Entity.Booking.PassengerService ToEntity(this Message.Booking.PassengerService objMessage)
        {
            Entity.Booking.PassengerService objEntity = null;
            if (objMessage != null)
            {
               // objEntity = new Entity.Booking.PassengerService();
               // objEntity.BookingSegmentId = objMessage.booking_segment_id;
               // objEntity.PassengerId = objMessage.passenger_id;
               // objEntity.PassengerSegmentServiceId = Guid.NewGuid();

               // objEntity.NumberOfUnits = objMessage.number_of_units;
               // objEntity.SpecialServiceRcd = objMessage.special_service_rcd;
               // objEntity.ServiceText = objMessage.service_text;
               //// objEntity.SpecialServiceStatusRcd = objMessage.special_service_status_rcd;
               // objEntity.UnitRcd = objMessage.unit_rcd;

            }
            return objEntity;
        }

    }
}
