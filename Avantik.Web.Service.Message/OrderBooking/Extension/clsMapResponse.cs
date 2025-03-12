using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;

namespace Avantik.Web.Service.Message.OrderBooking.Extension
{
    [MessageContract]
    public static class OrderMessageMapping
    {
        #region MyRegion
        public static FlightSegmentResponse MapToOrderSegmentResponse(this Message.OrderBooking.FlightSegment segment)
        {
            try
            {
                FlightSegmentResponse segRes = new FlightSegmentResponse();

                segRes.booking_segment_id = segment.booking_segment_id;
                segRes.flight_connection_id = segment.flight_connection_id;
                segRes.departure_date = segment.departure_date;
                segRes.airline_rcd = segment.airline_rcd;
                segRes.flight_number = segment.flight_number;
                segRes.origin_rcd = segment.origin_rcd;
                segRes.destination_rcd = segment.destination_rcd;
                segRes.od_origin_rcd = segment.od_origin_rcd;
                segRes.od_destination_rcd = segment.od_destination_rcd;
                segRes.booking_class_rcd = segment.booking_class_rcd;
                segRes.boarding_class_rcd = segment.boarding_class_rcd;
                segRes.segment_status_rcd = segment.segment_status_rcd;
                //Use for passenger count in the flight.
                segRes.number_of_units = segment.number_of_units;

                return segRes;
            }
            catch
            {
                throw;
            }
        }
        public static IList<FlightSegmentResponse> MapToOrderSegmentResponse(this IList<Message.OrderBooking.FlightSegment> segment)
        {
            try
            {
                if (segment != null && segment.Count > 0)
                {
                    IList<FlightSegmentResponse> segments = new List<FlightSegmentResponse>();

                    for (int i = 0; i < segment.Count; i++)
                    {
                        segments.Add(segment[i].MapToOrderSegmentResponse());
                    }

                    return segments;
                }
            }
            catch
            {
                throw;
            }
            return null;
        }

        public static MappingResponse MapToOrderMappingResponse(this Message.OrderBooking.Mapping mapping)
        {
            try
            {
                MappingResponse mpRes = new MappingResponse();

                mpRes.passenger_id = mapping.passenger_id;
                mpRes.booking_segment_id = mapping.booking_segment_id;
                mpRes.not_valid_before_date = mapping.not_valid_before_date;
                mpRes.not_valid_after_date = mapping.not_valid_after_date;
                mpRes.refund_with_charge_hours = mapping.refund_with_charge_hours;
                mpRes.refund_not_possible_hours = mapping.refund_not_possible_hours;
                mpRes.piece_allowance = mapping.piece_allowance;
                mpRes.fare_code = mapping.fare_code;
                mpRes.seat_number = mapping.seat_number;
                mpRes.currency_rcd = mapping.currency_rcd;
                mpRes.endorsement_text = mapping.endorsement_text;
                mpRes.restriction_text = mapping.restriction_text;
                mpRes.fare_amount = mapping.fare_amount;
                mpRes.fare_amount_incl = mapping.fare_amount_incl;
                mpRes.fare_vat = mapping.fare_vat;
                mpRes.net_total = mapping.net_total;
                mpRes.commission_amount = mapping.commission_amount;
                mpRes.commission_amount_incl = mapping.commission_amount_incl;
                mpRes.commission_percentage = mapping.commission_percentage;
                mpRes.public_fare_amount = mapping.public_fare_amount;
                mpRes.public_fare_amount_incl = mapping.public_fare_amount_incl;
                mpRes.refund_charge = mapping.refund_charge;
                mpRes.baggage_weight = mapping.baggage_weight;
                mpRes.redemption_points = mapping.redemption_points;
                mpRes.e_ticket_flag = mapping.e_ticket_flag;
                mpRes.refundable_flag = mapping.refundable_flag;
                mpRes.transferable_fare_flag = mapping.transferable_fare_flag;
                mpRes.through_fare_flag = mapping.through_fare_flag;
                mpRes.it_fare_flag = mapping.it_fare_flag;
                mpRes.duty_travel_flag = mapping.duty_travel_flag;
                mpRes.standby_flag = mapping.standby_flag;
                mpRes.exclude_pricing_flag = mapping.exclude_pricing_flag;

                return mpRes;
            }
            catch
            {
                throw;
            }
            //return null;
        }
        public static IList<MappingResponse> MapToOrderMappingsResponse(this IList<Message.OrderBooking.Mapping> mapping)
        {
            try
            {
                if (mapping != null && mapping.Count > 0)
                {
                    IList<MappingResponse> mappings = new List<MappingResponse>();

                    for (int i = 0; i < mapping.Count; i++)
                    {
                        mappings.Add(mapping[i].MapToOrderMappingResponse());
                    }
                    return mappings;
                }
            }
            catch
            {
                throw;
            }
            return null;
        }

        public static FeeResponse MapToOrderFeeResponse(this Message.OrderBooking.Fee fee)
        {
            try
            {
                FeeResponse feRes = new FeeResponse();

                feRes.passenger_id = fee.passenger_id;
                feRes.booking_segment_id = fee.booking_segment_id;
                feRes.passenger_segment_service_id = fee.passenger_segment_service_id;
                feRes.fee_rcd = fee.fee_rcd;
                feRes.fee_id = fee.fee_id;
                feRes.fee_category_rcd = fee.fee_category_rcd;

                feRes.vendor_rcd = fee.vendor_rcd;
                feRes.od_origin_rcd = fee.od_origin_rcd;
                feRes.od_destination_rcd = fee.od_destination_rcd;
                feRes.currency_rcd = fee.currency_rcd;
                feRes.charge_currency_rcd = fee.charge_currency_rcd;
                feRes.unit_rcd = fee.unit_rcd;
                feRes.mpd_number = fee.mpd_number;
                feRes.comment = fee.comment;
                feRes.external_reference = fee.external_reference;
                feRes.fee_amount = fee.fee_amount;
                feRes.fee_amount_incl = fee.fee_amount_incl;
                feRes.vat_percentage = fee.vat_percentage;
                feRes.charge_amount = fee.charge_amount;
                feRes.charge_amount_incl = fee.charge_amount_incl;
                feRes.weight_lbs = fee.weight_lbs;
                feRes.weight_kgs = fee.weight_kgs;
                feRes.number_of_units = fee.number_of_units;

                return feRes;
            }
            catch
            {
                throw;
            }
        }
        public static IList<FeeResponse> MapToOrderFeesResponse(this IList<Message.OrderBooking.Fee> fee)
        {
            try
            {
                if (fee != null && fee.Count > 0)
                {
                    IList<FeeResponse> fees = new List<FeeResponse>();
                    for (int i = 0; i < fee.Count; i++)
                    {
                        fees.Add(fee[i].MapToOrderFeeResponse());
                    }
                    return fees;
                }
            }
            catch
            {
                throw;
            }
            return null;
        }

        public static ServiceResponse MapToOrderServiceResponse(this Message.OrderBooking.Service service)
        {
            try
            {
                ServiceResponse svRes = new ServiceResponse();

                svRes.passenger_id = service.passenger_id;
                svRes.booking_segment_id = service.booking_segment_id;
                svRes.number_of_units = service.number_of_units;
                svRes.special_service_rcd = service.special_service_rcd;
                svRes.service_text = service.service_text;
                svRes.unit_rcd = service.unit_rcd;

                return svRes;
            }
            catch
            {
                throw;
            }
        }
        public static IList<ServiceResponse> MapToOrderServicesResponse(this IList<Message.OrderBooking.Service> service)
        {
            try
            {
                if (service != null && service.Count > 0)
                {
                    IList<ServiceResponse> services = new List<ServiceResponse>();
                    for (int i = 0; i < service.Count; i++)
                    {
                        services.Add(service[i].MapToOrderServiceResponse());
                    }

                    return services;
                }
            }
            catch
            {
                throw;
            }
            return null;
        }



        public static TaxResponse MapToOrderTaxResponse(this Message.OrderBooking.Tax service)
        {
            try
            {
                TaxResponse svRes = new TaxResponse();

                svRes.passenger_id = service.passenger_id;
                svRes.booking_segment_id = service.booking_segment_id;
                svRes.tax_rcd = service.tax_rcd;
                svRes.sales_amount = service.sales_amount;
                svRes.sales_amount_incl = service.sales_amount_incl;
                svRes.sales_currency_rcd = service.sales_currency_rcd;
                svRes.tax_amount = service.tax_amount;
                svRes.tax_amount_incl = service.tax_amount_incl;
                svRes.tax_currency_rcd = service.tax_currency_rcd;

                return svRes;
            }
            catch
            {
                throw;
            }
        }
        public static IList<TaxResponse> MapToOrderTaxResponse(this IList<Message.OrderBooking.Tax> service)
        {
            try
            {
                if (service != null && service.Count > 0)
                {
                    IList<TaxResponse> services = new List<TaxResponse>();
                    for (int i = 0; i < service.Count; i++)
                    {
                        services.Add(service[i].MapToOrderTaxResponse());
                    }

                    return services;
                }
            }
            catch
            {
                throw;
            }
            return null;
        }

        #endregion        
    }
}
