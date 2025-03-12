using Avantik.Web.Service.Entity.Booking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Exception.Booking;
using Avantik.Web.Service.Entity;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class RecordsetObjectBooking
    {
        public static void ToRecordset(this  BookingHeader header, ref ADODB.Recordset rs)
        {
            if (header != null)
            {
                try
                {
                    rs.AddNew();
                    if (header.BookingId != Guid.Empty)
                    {
                        rs.Fields["booking_id"].Value = header.BookingId.ToRsString();
                    }
                    
                    rs.Fields["booking_source_rcd"].Value = header.BookingSourceRcd;
                    rs.Fields["currency_rcd"].Value = header.CurrencyRcd;
                    if (header.ClientProfileId != Guid.Empty)
                    {
                        rs.Fields["client_profile_id"].Value = header.ClientProfileId.ToRsString();
                    }
                    
                    rs.Fields["booking_number"].Value = header.BookingNumber;
                    rs.Fields["record_locator"].Value = header.RecordLocator;
                    rs.Fields["number_of_adults"].Value = header.NumberOfAdults;
                    rs.Fields["number_of_children"].Value = header.NumberOfChildren;
                    rs.Fields["number_of_infants"].Value = header.NumberOfInfants;
                    rs.Fields["language_rcd"].Value = header.LanguageRcd;
                    rs.Fields["agency_code"].Value = header.AgencyCode;
                    rs.Fields["contact_name"].Value = header.ContactName;
                    rs.Fields["contact_email"].Value = header.ContactEmail;
                    rs.Fields["phone_mobile"].Value = header.PhoneMobile;
                    rs.Fields["phone_home"].Value = header.PhoneHome;
                    rs.Fields["phone_business"].Value = header.PhoneBusiness;
                    rs.Fields["received_from"].Value = header.ReceivedFrom;
                    rs.Fields["phone_fax"].Value = header.PhoneFax;
                    rs.Fields["phone_search"].Value = header.PhoneSearch;
                    rs.Fields["comment"].Value = header.Comment;
                    rs.Fields["notify_by_email_flag"].Value = header.NotifyByEmailFlag;
                    rs.Fields["notify_by_sms_flag"].Value = header.NotifyBySmsFlag;
                    rs.Fields["group_name"].Value = header.GroupName;
                    rs.Fields["group_booking_flag"].Value = header.GroupBookingFlag;
                    rs.Fields["agency_name"].Value = header.AgencyName;
                    rs.Fields["own_agency_flag"].Value = header.OwnAgencyFlag;
                    rs.Fields["web_agency_flag"].Value = header.WebAgencyFlag;
                    rs.Fields["client_number"].Value = header.ClientNumber;
                    rs.Fields["lastname"].Value = header.Lastname;
                    rs.Fields["firstname"].Value = header.Firstname;
                    rs.Fields["city"].Value = header.City;
                    rs.Fields["street"].Value = header.Street;
                    rs.Fields["lock_date_time"].Value = header.LockDateTime;
                    rs.Fields["middlename"].Value = header.Middlename;
                    rs.Fields["address_line1"].Value = header.AddressLine1;
                    rs.Fields["address_line2"].Value = header.AddressLine2;
                    rs.Fields["state"].Value = header.State;
                    rs.Fields["district"].Value = header.District;
                    rs.Fields["province"].Value = header.Province;
                    rs.Fields["zip_code"].Value = header.ZipCode;
                    rs.Fields["po_box"].Value = header.PoBox;
                    rs.Fields["country_rcd"].Value = header.CountryRcd;
                    rs.Fields["title_rcd"].Value = header.TitleRcd;
                    rs.Fields["external_payment_reference"].Value = header.ExternalPaymentReference;
                    rs.Fields["create_by"].Value = header.CreateBy.ToRsString();
                    rs.Fields["update_by"].Value = header.UpdateBy.ToRsString();

                    if (header.CreateDateTime != DateTime.MinValue)
                        rs.Fields["create_date_time"].Value = header.CreateDateTime;
                    if (header.UpdateDateTime != DateTime.MinValue)
                        rs.Fields["update_date_time"].Value = header.UpdateDateTime;

                    rs.Fields["cost_center"].Value = header.CostCenter;
                    rs.Fields["purchase_order"].Value = header.PurchaseOrder;
                    rs.Fields["project_number"].Value = header.ProjectNumber;
                    rs.Fields["ip_address"].Value = header.IpAddress;
                    rs.Fields["approval_flag"].Value = header.ApprovalFlag;
                    rs.Fields["invoice_receiver"].Value = header.InvoiceReceiver;
                    if (string.IsNullOrEmpty(header.TaxId) == false)
                    {
                        rs.Fields["tax_id"].Value = header.TaxId;
                    }
                    
                    rs.Fields["newsletter_flag"].Value = header.NewsletterFlag;
                    rs.Fields["contact_email_cc"].Value = header.ContactEmailCc;
                    rs.Fields["mobile_email"].Value = header.MobileEmail;
                    rs.Fields["vendor_rcd"].Value = header.VendorRcd;
                    rs.Fields["no_vat_flag"].Value = header.NoVatFlag;
                    rs.Fields["company_name"].Value = header.CompanyName;
                    rs.Fields["business_flag"].Value = header.BusinessFlag;
                }
                catch
                {
                    throw;
                }
            }
        }

        public static void ToRecordset(this  IList<FlightSegment> segment, ref ADODB.Recordset rs, bool NOHK)
        {
            if (segment != null && segment.Count > 0)
            {
                try
                {
                    foreach (FlightSegment fs in segment)
                    {
                        if (fs.SegmentStatusRcd != "HK")
                        {
                            rs.AddNew();
                            rs.Fields["origin_rcd"].Value = fs.OriginRcd;
                            rs.Fields["destination_rcd"].Value = fs.DestinationRcd;

                            if (fs.CreateBy != Guid.Empty)
                                rs.Fields["create_by"].Value = fs.CreateBy.ToRsString();

                            if (fs.UpdateBy != Guid.Empty)
                                rs.Fields["update_by"].Value = fs.UpdateBy.ToRsString();

                            if (fs.CreateDateTime != DateTime.MinValue)
                                rs.Fields["create_date_time"].Value = fs.CreateDateTime;
                            if (fs.UpdateDateTime != DateTime.MinValue)
                                rs.Fields["update_date_time"].Value = fs.UpdateDateTime;

                            if (fs.FlightId != Guid.Empty)
                                rs.Fields["flight_id"].Value = fs.FlightId.ToRsString();

                            rs.Fields["airline_rcd"].Value = fs.AirlineRcd;
                            rs.Fields["flight_number"].Value = fs.FlightNumber;

                            if (fs.DepartureDate != DateTime.MinValue)
                                rs.Fields["departure_date"].Value = fs.DepartureDate;
                            if (fs.ArrivalDate != DateTime.MinValue)
                                rs.Fields["arrival_date"].Value = fs.ArrivalDate;

                            rs.Fields["od_origin_rcd"].Value = fs.OdOriginRcd;
                            rs.Fields["od_destination_rcd"].Value = fs.OdDestinationRcd;
                            rs.Fields["overbook_flag"].Value = fs.OverbookFlag;
                            if (fs.FlightConnectionId != Guid.Empty)
                            {
                                rs.Fields["flight_connection_id"].Value = fs.FlightConnectionId.ToRsString();
                            }
                            rs.Fields["journey_time"].Value = fs.JourneyTime;
                            rs.Fields["departure_time"].Value = fs.DepartureTime;
                            rs.Fields["arrival_time"].Value = fs.ArrivalTime;

                            rs.Fields["booking_class_rcd"].Value = fs.BookingClassRcd;
                            rs.Fields["boarding_class_rcd"].Value = fs.BoardingClassRcd;
                            rs.Fields["aircraft_type_rcd"].Value = fs.AircraftTypeRcd;
                            rs.Fields["planned_arrival_time"].Value = fs.PlannedArrivalTime;
                            if (fs.BookingSegmentId != Guid.Empty)
                            {
                                rs.Fields["booking_segment_id"].Value = fs.BookingSegmentId.ToRsString();
                            }
                            rs.Fields["segment_status_rcd"].Value = fs.SegmentStatusRcd;
                            if (fs.BookingId != Guid.Empty)
                            {
                                rs.Fields["booking_id"].Value = fs.BookingId.ToRsString();
                            }
                            rs.Fields["number_of_units"].Value = fs.NumberOfUnits;
                            rs.Fields["info_segment_flag"].Value = fs.InfoSegmentFlag;
                            rs.Fields["temp_seatmap_flag"].Value = fs.TempSeatmapFlag;
                            rs.Fields["allow_web_checkin_flag"].Value = fs.AllowWebCheckinFlag;
                            rs.Fields["segment_change_status_rcd"].Value = fs.SegmentChangeStatusRcd;
                        }
                    }

                }
                catch
                {
                    throw;
                }

            }
        }

        public static void ToRecordset(this  IList<FlightSegment> segment, ref ADODB.Recordset rs)
        {
            if (segment != null && segment.Count > 0)
            {
                try
                {
                    foreach (FlightSegment fs in segment)
                    {
                        rs.AddNew();
                        rs.Fields["origin_rcd"].Value = fs.OriginRcd;
                        rs.Fields["destination_rcd"].Value = fs.DestinationRcd;

                        if (fs.CreateBy != Guid.Empty)
                            rs.Fields["create_by"].Value = fs.CreateBy.ToRsString();

                        if (fs.UpdateBy != Guid.Empty)
                            rs.Fields["update_by"].Value = fs.UpdateBy.ToRsString();

                        if (fs.CreateDateTime != DateTime.MinValue)
                            rs.Fields["create_date_time"].Value = fs.CreateDateTime;
                        if (fs.UpdateDateTime != DateTime.MinValue)
                            rs.Fields["update_date_time"].Value = fs.UpdateDateTime;

                        if (fs.FlightId != Guid.Empty)
                            rs.Fields["flight_id"].Value = fs.FlightId.ToRsString();

                        rs.Fields["airline_rcd"].Value = fs.AirlineRcd;
                        rs.Fields["flight_number"].Value = fs.FlightNumber;

                        if (fs.DepartureDate != DateTime.MinValue)
                            rs.Fields["departure_date"].Value = fs.DepartureDate;
                        if (fs.ArrivalDate != DateTime.MinValue)
                            rs.Fields["arrival_date"].Value = fs.ArrivalDate;

                        rs.Fields["od_origin_rcd"].Value = fs.OdOriginRcd;
                        rs.Fields["od_destination_rcd"].Value = fs.OdDestinationRcd;
                        rs.Fields["overbook_flag"].Value = fs.OverbookFlag;
                        if (fs.FlightConnectionId != Guid.Empty)
                        {
                            rs.Fields["flight_connection_id"].Value = fs.FlightConnectionId.ToRsString();
                        }
                        rs.Fields["journey_time"].Value = fs.JourneyTime;
                        rs.Fields["departure_time"].Value = fs.DepartureTime;
                        rs.Fields["arrival_time"].Value = fs.ArrivalTime;

                        rs.Fields["booking_class_rcd"].Value = fs.BookingClassRcd;
                        rs.Fields["boarding_class_rcd"].Value = fs.BoardingClassRcd;
                        rs.Fields["aircraft_type_rcd"].Value = fs.AircraftTypeRcd;
                        rs.Fields["planned_arrival_time"].Value = fs.PlannedArrivalTime;
                        if (fs.BookingSegmentId != Guid.Empty)
                        {
                            rs.Fields["booking_segment_id"].Value = fs.BookingSegmentId.ToRsString();
                        }
                        rs.Fields["segment_status_rcd"].Value = fs.SegmentStatusRcd;
                        if (fs.BookingId != Guid.Empty)
                        {
                            rs.Fields["booking_id"].Value = fs.BookingId.ToRsString();
                        }
                        rs.Fields["number_of_units"].Value = fs.NumberOfUnits;
                        rs.Fields["info_segment_flag"].Value = fs.InfoSegmentFlag;
                        rs.Fields["temp_seatmap_flag"].Value = fs.TempSeatmapFlag;
                        rs.Fields["allow_web_checkin_flag"].Value = fs.AllowWebCheckinFlag;
                        rs.Fields["segment_change_status_rcd"].Value = fs.SegmentChangeStatusRcd;
                    }

                }
                catch
                {
                    throw;
                }
               
            }
        }

        public static void ToRecordset(this Entity.Booking.Flight flight, ref ADODB.Recordset rs)
        {
            if (flight != null)
            {
                try
                {
                    rs.AddNew();
                    if (flight.FlightId != Guid.Empty)
                    {
                        rs.Fields["flight_id"].Value = flight.FlightId.ToRsString();
                    }
                    if (flight.FairId != Guid.Empty)
                    {
                        rs.Fields["fair_id"].Value = flight.FairId.ToRsString();
                    }
                    if (flight.FlightConnectionId != Guid.Empty)
                    {
                        rs.Fields["flight_connection_id"].Value = flight.FlightConnectionId.ToRsString();
                    }
                    rs.Fields["airline_rcd"].Value = flight.AirlineRcd;
                    rs.Fields["boarding_class_rcd"].Value = flight.BoardingClassRcd;
                    rs.Fields["booking_class_rcd"].Value = flight.BookingClassRcd;
                    rs.Fields["transit_points"].Value = flight.TransitPoints;
                    rs.Fields["transit_points_name"].Value = flight.TransitpointsName;
                    rs.Fields["number_of_units"].Value = flight.NumberOfUnits;
                    rs.Fields["e_ticket_flag"].Value = flight.EticketFlag;

                }
                catch
                {
                    throw;
                }
            }
        }
        public static void ToRecordset(this  IList<Passenger> passenger, ref ADODB.Recordset rs)
        {
            if (passenger != null && passenger.Count > 0)
            {
                try
                {
                    foreach (Passenger p in passenger)
                    {
                        rs.AddNew();
                        rs.Fields["title_rcd"].Value = p.TitleRcd;
                        rs.Fields["lastname"].Value = p.Lastname;
                        rs.Fields["firstname"].Value = p.Firstname;
                        rs.Fields["middlename"].Value = p.Middlename;
                        rs.Fields["nationality_rcd"].Value = p.NationalityRcd;
                        rs.Fields["passenger_weight"].Value = p.PassengerWeight;
                        rs.Fields["gender_type_rcd"].Value = p.GenderTypeRcd;
                        rs.Fields["passenger_type_rcd"].Value = p.PassengerTypeRcd;
                        rs.Fields["client_number"].Value = p.ClientNumber;

                        #region Document Information
                        rs.Fields["document_type_rcd"].Value = p.DocumentTypeRcd;
                        rs.Fields["passport_number"].Value = p.PassportNumber;
                        rs.Fields["passport_issue_date"].Value = p.PassportIssueDate;
                        rs.Fields["passport_expiry_date"].Value = p.PassportExpiryDate;
                        rs.Fields["passport_issue_place"].Value = p.PassportIssuePlace;
                        rs.Fields["passport_birth_place"].Value = p.PassportBirthPlace;
                        rs.Fields["date_of_birth"].Value = p.DateOfBirth;
                        rs.Fields["country_code_long"].Value = p.CountryCodeLong;
                        rs.Fields["passport_issue_country_rcd"].Value = p.PassportIssueCountryRcd;
                        #endregion

                        #region Update Information
                        if (p.CreateBy != Guid.Empty)
                            rs.Fields["create_by"].Value = p.CreateBy.ToRsString();

                        if (p.CreateDateTime != DateTime.MinValue)
                            rs.Fields["create_date_time"].Value = p.CreateDateTime;

                        if (p.UpdateBy != Guid.Empty)
                            rs.Fields["update_by"].Value = p.UpdateBy.ToRsString();

                        if (p.UpdateDateTime != DateTime.MinValue)
                            rs.Fields["update_date_time"].Value = p.UpdateDateTime;

                        #endregion
                        if (p.PassengerId != Guid.Empty)
                        {
                            rs.Fields["passenger_id"].Value = p.PassengerId.ToRsString();
                        }
                        if (p.BookingId != Guid.Empty)
                        {
                            rs.Fields["booking_id"].Value = p.BookingId.ToRsString(); 
                        }
                        if (p.PassengerProfileId != Guid.Empty)
                        {
                            rs.Fields["passenger_profile_id"].Value = p.PassengerProfileId.ToRsString();
                        }
                        
                        rs.Fields["member_number"].Value = p.MemberNumber;
                        rs.Fields["redress_number"].Value = p.RedressNumber;
                        rs.Fields["wheelchair_flag"].Value = p.WheelchairFlag;
                        rs.Fields["vip_flag"].Value = p.VipFlag;
                        if (p.GuardianPassengerId != Guid.Empty)
                        {
                            rs.Fields["guardian_passenger_id"].Value = p.GuardianPassengerId.ToRsString();
                        }
                        rs.Fields["member_airline_rcd"].Value = p.MemberAirlineRcd;
                        

                    }

                }
                catch
                {
                    throw;
                }

            }
        }
        public static void ToRecordset(this  IList<Mapping> mapping, ref ADODB.Recordset rs)
        {
            if (mapping != null && mapping.Count > 0)
            {
                try
                {
                    foreach (Mapping mp in mapping)
                    {
                        rs.AddNew();
                        rs.Fields["origin_rcd"].Value = mp.OriginRcd;
                        rs.Fields["destination_rcd"].Value = mp.DestinationRcd;

                        #region "Property"

                        if (mp.FlightId != Guid.Empty)
                        {
                            rs.Fields["flight_id"].Value = mp.FlightId.ToRsString();
                        }

                        if (mp.ExchangedSegmentId != Guid.Empty)
                        {
                            rs.Fields["exchanged_segment_id"].Value = mp.ExchangedSegmentId.ToRsString();
                        }
                        rs.Fields["airline_rcd"].Value = mp.AirlineRcd;
                        rs.Fields["flight_number"].Value = mp.FlightNumber;
                        rs.Fields["refundable_flag"].Value = mp.RefundableFlag;

                        if (mp.DepartureDate != DateTime.MinValue)
                        {
                            rs.Fields["departure_date"].Value = mp.DepartureDate;
                        }

                        rs.Fields["od_origin_rcd"].Value = mp.OdOriginRcd;
                        rs.Fields["od_destination_rcd"].Value = mp.OdDestinationRcd;
                        if (mp.FlightConnectionId != Guid.Empty)
                        {
                            rs.Fields["flight_connection_id"].Value = mp.FlightConnectionId.ToRsString();
                        }
                        rs.Fields["departure_time"].Value = mp.DepartureTime;


                        if (mp.FareId != Guid.Empty)
                        {
                            rs.Fields["fare_id"].Value = mp.FareId.ToRsString();
                        }

                        rs.Fields["booking_class_rcd"].Value = mp.BookingClassRcd;
                        rs.Fields["boarding_class_rcd"].Value = mp.BoardingClassRcd;
                        #endregion

                        #region "Property"
                        if (mp.BookingSegmentId != Guid.Empty)
                        {
                            rs.Fields["booking_segment_id"].Value = mp.BookingSegmentId.ToRsString();
                        }
                        if (mp.BookingId != Guid.Empty)
                        {
                            rs.Fields["booking_id"].Value = mp.BookingId.ToRsString();
                        }
                        #endregion


                        #region Passenger information
                        if (mp.PassengerId != Guid.Empty)
                        {
                            rs.Fields["passenger_id"].Value = mp.PassengerId.ToRsString();
                        }

                        rs.Fields["gender_type_rcd"].Value = mp.GenderTypeRcd;
                        rs.Fields["title_rcd"].Value = mp.TitleRcd;
                        rs.Fields["firstname"].Value = mp.Firstname;
                        rs.Fields["lastname"].Value = mp.Lastname;
                        rs.Fields["passenger_type_rcd"].Value = mp.PassengerTypeRcd;

                        if (mp.DateOfBirth != DateTime.MinValue)
                        {
                            rs.Fields["date_of_birth"].Value = mp.DateOfBirth;
                        }
                        
                        rs.Fields["contact_email"].Value = mp.ContactEmail;
                        rs.Fields["passenger_status_rcd"].Value = mp.PassengerStatusRcd;

                        #endregion
                        #region Ticket information
                        rs.Fields["inventory_class_rcd"].Value = mp.InventoryClassRcd;
                        rs.Fields["seat_number"].Value = mp.SeatNumber;
                        rs.Fields["seat_row"].Value = mp.SeatRow;
                        rs.Fields["seat_column"].Value = mp.SeatColumn;
                        rs.Fields["net_total"].Value = mp.NetTotal;
                        rs.Fields["tax_amount"].Value = mp.TaxAmount;
                        rs.Fields["fare_amount"].Value = mp.FareAmount;
                        rs.Fields["yq_amount"].Value = mp.YqAmount;
                        rs.Fields["base_ticketing_fee_amount"].Value = mp.BaseTicketingFeeAmount;
                        rs.Fields["ticketing_fee_amount"].Value = mp.TicketingFeeAmount;
                        rs.Fields["reservation_fee_amount"].Value = mp.ReservationFeeAmount;
                        rs.Fields["commission_amount"].Value = mp.CommissionAmount;
                        rs.Fields["fare_vat"].Value = mp.FareVat;
                        rs.Fields["tax_vat"].Value = mp.TaxVat;
                        rs.Fields["yq_vat"].Value = mp.YqVat;
                        rs.Fields["ticketing_fee_vat"].Value = mp.TicketingFeeVat;
                        rs.Fields["reservation_fee_vat"].Value = mp.ReservationFeeVat;
                        rs.Fields["fare_amount_incl"].Value = mp.FareAmountIncl;
                        rs.Fields["tax_amount_incl"].Value = mp.TaxAmountIncl;
                        rs.Fields["yq_amount_incl"].Value = mp.YqAmountIncl;
                        rs.Fields["public_fare_amount_incl"].Value = mp.PublicFareAmountIncl;
                        rs.Fields["public_fare_amount"].Value = mp.PublicFareAmount;
                        rs.Fields["commission_amount_incl"].Value = mp.CommissionAmountIncl;
                        rs.Fields["commission_percentage"].Value = mp.CommissionPercentage;
                        rs.Fields["ticketing_fee_amount_incl"].Value = mp.TicketingFeeAmountIncl;
                        rs.Fields["reservation_fee_amount_incl"].Value = mp.ReservationFeeAmountIncl;
                        rs.Fields["acct_net_total"].Value = mp.AcctNetTotal;
                        rs.Fields["acct_tax_amount"].Value = mp.AcctTaxAmount;
                        rs.Fields["acct_fare_amount"].Value = mp.AcctFareAmount;
                        rs.Fields["acct_yq_amount"].Value = mp.AcctYqAmount;
                        rs.Fields["acct_ticketing_fee_amount"].Value = mp.AcctTicketingFeeAmount;
                        rs.Fields["acct_reservation_fee_amount"].Value = mp.AcctReservationFeeAmount;
                        rs.Fields["acct_commission_amount"].Value = mp.AcctCommissionAmount;
                        rs.Fields["acct_fare_amount_incl"].Value = mp.AcctFareAmountIncl;
                        rs.Fields["acct_tax_amount_incl"].Value = mp.AcctTaxAmountIncl;
                        rs.Fields["acct_yq_amount_incl"].Value = mp.AcctYqAmountIncl;
                        rs.Fields["acct_commission_amount_incl"].Value = mp.AcctCommissionAmountIncl;
                        rs.Fields["acct_ticketing_fee_amount_incl"].Value = mp.AcctTicketingFeeAmountIncl;
                        rs.Fields["acct_reservation_fee_amount_incl"].Value = mp.AcctReservationFeeAmountIncl;
                        rs.Fields["fare_code"].Value = mp.FareCode;

                        if (mp.FareDateTime != DateTime.MinValue)
                        {
                            rs.Fields["fare_date_time"].Value = mp.FareDateTime;
                        }
                        
                        rs.Fields["e_ticket_flag"].Value = mp.ETicketFlag;

                        rs.Fields["e_ticket_status"].Value = mp.ETicketStatus;

                        rs.Fields["standby_flag"].Value = mp.StandbyFlag;
                        rs.Fields["transferable_fare_flag"].Value = mp.TransferableFareFlag;
                        rs.Fields["agency_code"].Value = mp.AgencyCode;
                        rs.Fields["ticket_number"].Value = mp.TicketNumber;

                        if (mp.TicketingDateTime != DateTime.MinValue)
                        {
                            rs.Fields["ticketing_date_time"].Value = mp.TicketingDateTime;
                        }

                        if (mp.TicketingBy != Guid.Empty)
                        {
                            rs.Fields["ticketing_by"].Value = mp.TicketingBy.ToRsString();
                        }
                        
                        rs.Fields["check_in_sequence"].Value = mp.CheckInSequence;
                        rs.Fields["group_sequence"].Value = mp.GroupSequence;
                        rs.Fields["unload_by"].Value = mp.UnloadBy.ToRsString();

                        if (mp.UnloadDateTime != DateTime.MinValue)
                        {
                            rs.Fields["unload_date_time"].Value = mp.UnloadDateTime;
                        }
                        
                        rs.Fields["number_of_pieces"].Value = mp.NumberOfPieces;
                        rs.Fields["baggage_weight"].Value = mp.BaggageWeight;
                        rs.Fields["excess_baggage_weight"].Value = mp.ExcessBaggageWeight;
                        rs.Fields["check_in_baggage_weight"].Value = mp.CheckInBaggageWeight;

                        if (mp.CheckInBy != Guid.Empty)
                        {
                            rs.Fields["check_in_by"].Value = mp.CheckInBy.ToRsString();
                        }
                        
                        if (mp.CancelBy != Guid.Empty)
                        {
                            rs.Fields["cancel_by"].Value = mp.CancelBy.ToRsString();
                        }
                        

                        if (mp.ExchangedDateTime != DateTime.MinValue)
                        {
                            rs.Fields["exchanged_date_time"].Value = mp.ExchangedDateTime;
                        }
                        

                        if (mp.CancelDateTime != DateTime.MinValue)
                        {
                            rs.Fields["cancel_date_time"].Value = mp.CancelDateTime;
                        }
                        
                        rs.Fields["restriction_text"].Value = mp.RestrictionText;
                        rs.Fields["endorsement_text"].Value = mp.EndorsementText;
                        rs.Fields["exclude_pricing_flag"].Value = mp.ExcludePricingFlag;

                        if (mp.VoidDateTime != DateTime.MinValue)
                        {
                            rs.Fields["void_date_time"].Value = mp.VoidDateTime;
                        }

                        rs.Fields["refund_charge"].Value = mp.RefundCharge;
                        rs.Fields["refundable_amount"].Value = mp.RefundableAmount;

                        if (mp.RefundDateTime != DateTime.MinValue)
                        {
                            rs.Fields["refund_date_time"].Value = mp.RefundDateTime;
                        }

                        if (mp.CheckInDateTime != DateTime.MinValue) 
                        {
                            rs.Fields["check_in_date_time"].Value = mp.CheckInDateTime;
                        }

                        if (mp.OnwardDepartureDate != DateTime.MinValue)
                        {
                            rs.Fields["onward_departure_date"].Value = mp.OnwardDepartureDate;
                        }
                            
                        rs.Fields["fare_type_rcd"].Value = mp.FareTypeRcd;
                        rs.Fields["redemption_points"].Value = mp.RedemptionPoints;
                        rs.Fields["seat_fee_rcd"].Value = mp.SeatFeeRcd;
                        rs.Fields["refund_with_charge_hours"].Value = mp.RefundWithChargeHours;
                        rs.Fields["refund_not_possible_hours"].Value = mp.RefundNotPossibleHours;
                        rs.Fields["duty_travel_flag"].Value = mp.DutyTravelFlag;

                        if (mp.NotValidAfterDate != DateTime.MinValue)
                            rs.Fields["not_valid_after_date"].Value = mp.NotValidAfterDate;
                        if (mp.NotValidBeforeDate != DateTime.MinValue)
                            rs.Fields["not_valid_before_date"].Value = mp.NotValidBeforeDate;

                        rs.Fields["advanced_seating_flag"].Value = mp.AdvancedSeatingFlag;
                        rs.Fields["fare_column"].Value = mp.FareColumn;
                        rs.Fields["currency_rcd"].Value = mp.CurrencyRcd;
                        #endregion

                        rs.Fields["piece_allowance"].Value = mp.PieceAllowance;
                        rs.Fields["it_fare_flag"].Value = mp.ItFareFlag;
                        rs.Fields["passenger_check_in_status_rcd"].Value = mp.PassengerCheckinStatusRcd;

                        //if (mp.CreateBy != Guid.Empty)
                        //{
                        //    rs.Fields["create_by"].Value = mp.CreateBy;
                        //}

                        //if (mp.UpdateBy != Guid.Empty)
                        //{
                        //    rs.Fields["update_by"].Value = mp.UpdateBy;
                        //}
                    }

                }
                catch
                {
                    throw;
                }

            }
        }
        public static void ToRecordset(this  IList<Payment> payment, ref ADODB.Recordset rs)
        {
            if (payment != null && payment.Count > 0)
            {
                try
                {
                    foreach (Payment p in payment)
                    {
                        rs.AddNew();
                        if (p.BookingPaymentId != Guid.Empty)
                        {
                            rs.Fields["booking_payment_id"].Value = p.BookingPaymentId.ToRsString();
                        }
                        if (p.BookingSegmentId != Guid.Empty)
                        {
                            rs.Fields["booking_segment_id"].Value = p.BookingSegmentId.ToRsString();
                        }
                        if (p.BookingId != Guid.Empty)
                        {
                            rs.Fields["booking_id"].Value = p.BookingId.ToRsString();
                        }
                        if (p.VoucherPaymentId != Guid.Empty)
                        {
                            rs.Fields["voucher_payment_id"].Value = p.VoucherPaymentId.ToRsString();
                        }
                        rs.Fields["form_of_payment_rcd"].Value = p.FormOfPaymentRcd;
                        rs.Fields["currency_rcd"].Value = p.CurrencyRcd;
                        rs.Fields["receive_currency_rcd"].Value = p.ReceiveCurrencyRcd;
                        rs.Fields["agency_code"].Value = p.AgencyCode;
                        rs.Fields["debit_agency_code"].Value = p.DebitAgencyCode;
                        rs.Fields["payment_amount"].Value = p.PaymentAmount;
                        rs.Fields["receive_payment_amount"].Value = p.ReceivePaymentAmount;
                        rs.Fields["acct_payment_amount"].Value = p.AcctPaymentAmount;

                        if (p.PaymentBy != Guid.Empty)
                            rs.Fields["payment_by"].Value = p.PaymentBy.ToRsString();

                        if (p.PaymentDateTime != DateTime.MinValue)
                            rs.Fields["payment_date_time"].Value = p.PaymentDateTime;
                        if (p.PaymentDueDateTime != DateTime.MinValue)
                            rs.Fields["payment_due_date_time"].Value = p.PaymentDueDateTime;

                        rs.Fields["document_amount"].Value = p.DocumentAmount;

                        if (p.VoidBy != Guid.Empty)
                            rs.Fields["void_by"].Value = p.VoidBy.ToRsString();

                        rs.Fields["expiry_month"].Value = p.ExpiryMonth;
                        rs.Fields["expiry_year"].Value = p.ExpiryYear;

                        if (p.VoidDateTime != DateTime.MinValue)
                            rs.Fields["void_date_time"].Value = p.VoidDateTime;

                        rs.Fields["record_locator"].Value = p.RecordLocator;
                        rs.Fields["cvv_code"].Value = p.CvvCode;
                        rs.Fields["name_on_card"].Value = p.NameOnCard;
                        rs.Fields["document_number"].Value = p.DocumentNumber;
                        rs.Fields["city"].Value = p.City;
                        rs.Fields["state"].Value = p.State;
                        rs.Fields["street"].Value = p.Street;

                        rs.Fields["address_line1"].Value = p.AddressLine1;
                        rs.Fields["address_line2"].Value = p.AddressLine2;
                        rs.Fields["zip_code"].Value = p.ZipCode;
                        rs.Fields["country_rcd"].Value = p.CountryRcd;

                        if (p.CreateBy != Guid.Empty)
                            rs.Fields["create_by"].Value = p.CreateBy.ToRsString();

                        rs.Fields["create_date_time"].Value = p.CreateDateTime;

                        if (p.UpdateBy != Guid.Empty)
                            rs.Fields["update_by"].Value = p.UpdateBy.ToRsString();

                        if (p.UpdateDateTime != DateTime.MinValue)
                            rs.Fields["update_date_time"].Value = p.UpdateDateTime;

                        rs.Fields["pos_indicator"].Value = p.PosIndicator;

                        rs.Fields["payment_remark"].Value = p.PaymentRemark;
                        rs.Fields["payment_number"].Value = p.PaymentNumber;

                        rs.Fields["issue_month"].Value = p.IssueMonth;
                        rs.Fields["issue_year"].Value = p.IssueYear;
                        rs.Fields["issue_number"].Value = p.IssueNumber;

                        if (p.ClientProfileId != Guid.Empty)
                        {
                            rs.Fields["client_profile_id"].Value = p.ClientProfileId.ToRsString();
                        }
                        rs.Fields["transaction_reference"].Value = p.TransactionReference;
                        rs.Fields["approval_code"].Value = p.ApprovalCode;
                        rs.Fields["bank_name"].Value = p.BankName;
                        rs.Fields["bank_code"].Value = p.BankCode;
                        rs.Fields["bank_iban"].Value = p.BankIban;
                        rs.Fields["ip_address"].Value = p.IpAddress;
                        rs.Fields["payment_reference"].Value = p.PaymentReference;
                        rs.Fields["payment_type_rcd"].Value = p.PaymentTypeRcd;
                        rs.Fields["form_of_payment_subtype_rcd"].Value = p.FormOfPaymentSubtypeRcd;
                        
                    }

                }
                catch
                {
                    throw;
                }

            }
        }
        public static void ToRecordset(this  IList<Entity.Booking.Fee> fee, ref ADODB.Recordset rs)
        {
            if (fee != null && fee.Count > 0)
            {
                try
                {
                    foreach (Entity.Booking.Fee f in fee)
                    {
                        if (f != null)
                        {
                            rs.AddNew();
                            if (f.BookingFeeId != Guid.Empty)
                            {
                                rs.Fields["booking_fee_id"].Value = f.BookingFeeId.ToRsString();
                            }
                            rs.Fields["fee_amount"].Value = f.FeeAmount;
                            if (f.BookingId != Guid.Empty)
                            {
                                rs.Fields["booking_id"].Value = f.BookingId.ToRsString();
                            }
                            if (f.PassengerId != Guid.Empty)
                            {
                                rs.Fields["passenger_id"].Value = f.PassengerId.ToRsString();
                            }
                            rs.Fields["currency_rcd"].Value = f.CurrencyRcd;
                            rs.Fields["acct_fee_amount"].Value = f.AcctFeeAmount;
                            if (f.FeeId != Guid.Empty)
                            {
                                rs.Fields["fee_id"].Value = f.FeeId.ToRsString();
                            }

                            rs.Fields["vat_percentage"].Value = f.VatPercentage;
                            rs.Fields["fee_amount_incl"].Value = f.FeeAmountIncl;
                            rs.Fields["acct_fee_amount_incl"].Value = f.AcctFeeAmountIncl;
                            rs.Fields["fee_rcd"].Value = f.FeeRcd;
                            rs.Fields["display_name"].Value = f.DisplayName;

                             if (f.AccountFeeBy != Guid.Empty)
                                rs.Fields["account_fee_by"].Value = f.AccountFeeBy.ToRsString();

                             if (f.AccountFeeDateTime != DateTime.MinValue)
                                 rs.Fields["account_fee_date_time"].Value = f.AccountFeeDateTime;
                             if (f.VoidDateTime != DateTime.MinValue)
                                 rs.Fields["void_date_time"].Value = f.VoidDateTime;

                            if (f.VoidBy != Guid.Empty)
                                rs.Fields["void_by"].Value = f.VoidBy.ToRsString();

                             if (f.CreateBy != Guid.Empty)
                                rs.Fields["create_by"].Value = f.CreateBy.ToRsString();

                             if (f.CreateDateTime != DateTime.MinValue)
                                 rs.Fields["create_date_time"].Value = f.CreateDateTime;

                            if (f.UpdateBy != Guid.Empty)
                                rs.Fields["update_by"].Value = f.UpdateBy.ToRsString();

                            if (f.UpdateDateTime != DateTime.MinValue)
                                rs.Fields["update_date_time"].Value = f.UpdateDateTime;

                            if (f.BookingSegmentId != Guid.Empty)
                            {
                                rs.Fields["booking_segment_id"].Value = f.BookingSegmentId.ToRsString();
                            }

                            rs.Fields["agency_code"].Value = f.AgencyCode;
                            if (f.PassengerSegmentServiceId != Guid.Empty)
                            {
                                rs.Fields["passenger_segment_service_id"].Value = f.PassengerSegmentServiceId.ToRsString();
                            }

                            rs.Fields["fee_category_rcd"].Value = f.FeeCategoryRcd;
                            rs.Fields["origin_rcd"].Value = f.OriginRcd;
                            rs.Fields["destination_rcd"].Value = f.DestinationRcd;
                            rs.Fields["od_origin_rcd"].Value = f.OdOriginRcd;
                            rs.Fields["od_destination_rcd"].Value = f.OdDestinationRcd;
                            rs.Fields["number_of_units"].Value = f.NumberOfUnits;
                            rs.Fields["change_comment"].Value = f.ChangeComment;
                            rs.Fields["comment"].Value = f.Comment;

                            if (!string.IsNullOrEmpty(f.Units))
                                rs.Fields["units"].Value = f.Units;

                            rs.Fields["charge_currency_rcd"].Value = f.ChargeCurrencyRcd;
                            rs.Fields["charge_amount"].Value = f.ChargeAmount;
                            rs.Fields["charge_amount_incl"].Value = f.ChargeAmountIncl;
                            rs.Fields["document_number"].Value = f.DocumentNumber;

                            if (f.DocumentDateTime != DateTime.MinValue)
                                rs.Fields["document_date_time"].Value = f.DocumentDateTime;

                            rs.Fields["external_reference"].Value = f.ExternalReference;
                            rs.Fields["vendor_rcd"].Value = f.VendorRcd;

                            rs.Fields["weight_lbs"].Value = f.WeightLbs;
                            rs.Fields["weight_kgs"].Value = f.WeightKgs;
                            rs.Fields["mpd_number"].Value = f.MpdNumber;

                            //rs.Fields["unit_rcd"].Value = f.UnitRcd;
                        }

                    }

                }
                catch
                {
                    throw;
                }


            }
        }

        public static void ToRecordset(this  IList<Entity.Booking.Fee> fee, ref ADODB.Recordset rs, Guid user)
        {
            if (fee != null && fee.Count > 0)
            {
                try
                {
                    foreach (Entity.Booking.Fee f in fee)
                    {
                        if (f != null)
                        {
                            rs.AddNew();
                            if (f.BookingFeeId != Guid.Empty)
                            {
                                rs.Fields["booking_fee_id"].Value = f.BookingFeeId.ToRsString();
                            }
                            rs.Fields["fee_amount"].Value = f.FeeAmount;
                            if (f.BookingId != Guid.Empty)
                            {
                                rs.Fields["booking_id"].Value = f.BookingId.ToRsString();
                            }
                            if (f.PassengerId != Guid.Empty)
                            {
                                rs.Fields["passenger_id"].Value = f.PassengerId.ToRsString();
                            }
                            rs.Fields["currency_rcd"].Value = f.CurrencyRcd;
                            rs.Fields["acct_fee_amount"].Value = f.AcctFeeAmount;
                            if (f.FeeId != Guid.Empty)
                            {
                                rs.Fields["fee_id"].Value = f.FeeId.ToRsString();
                            }

                            rs.Fields["vat_percentage"].Value = f.VatPercentage;
                            rs.Fields["fee_amount_incl"].Value = f.FeeAmountIncl;
                            rs.Fields["acct_fee_amount_incl"].Value = f.AcctFeeAmountIncl;
                            rs.Fields["fee_rcd"].Value = f.FeeRcd;
                            rs.Fields["display_name"].Value = f.DisplayName;

                            if (f.AccountFeeBy != Guid.Empty)
                                rs.Fields["account_fee_by"].Value = f.AccountFeeBy.ToRsString();
                            else
                                rs.Fields["account_fee_by"].Value = user.ToRsString();

                            if (f.AccountFeeDateTime != DateTime.MinValue)
                                rs.Fields["account_fee_date_time"].Value = f.AccountFeeDateTime;
                            if (f.VoidDateTime != DateTime.MinValue)
                                rs.Fields["void_date_time"].Value = f.VoidDateTime;

                            if (f.VoidBy != Guid.Empty)
                                rs.Fields["void_by"].Value = f.VoidBy.ToRsString();

                            if (f.CreateBy != Guid.Empty)
                                rs.Fields["create_by"].Value = f.CreateBy.ToRsString();
                            else
                                rs.Fields["create_by"].Value = user.ToRsString();

                            if (f.CreateDateTime != DateTime.MinValue)
                                rs.Fields["create_date_time"].Value = f.CreateDateTime;

                            if (f.UpdateBy != Guid.Empty)
                                rs.Fields["update_by"].Value = f.UpdateBy.ToRsString();

                            if (f.UpdateDateTime != DateTime.MinValue)
                                rs.Fields["update_date_time"].Value = f.UpdateDateTime;

                            if (f.BookingSegmentId != Guid.Empty)
                            {
                                rs.Fields["booking_segment_id"].Value = f.BookingSegmentId.ToRsString();
                            }

                            rs.Fields["agency_code"].Value = f.AgencyCode;
                            if (f.PassengerSegmentServiceId != Guid.Empty)
                            {
                                rs.Fields["passenger_segment_service_id"].Value = f.PassengerSegmentServiceId.ToRsString();
                            }

                            rs.Fields["fee_category_rcd"].Value = f.FeeCategoryRcd;
                            rs.Fields["origin_rcd"].Value = f.OriginRcd;
                            rs.Fields["destination_rcd"].Value = f.DestinationRcd;
                            rs.Fields["od_origin_rcd"].Value = f.OdOriginRcd;
                            rs.Fields["od_destination_rcd"].Value = f.OdDestinationRcd;
                            rs.Fields["number_of_units"].Value = f.NumberOfUnits;
                            rs.Fields["change_comment"].Value = f.ChangeComment;
                            rs.Fields["comment"].Value = f.Comment;

                            if (!string.IsNullOrEmpty(f.Units))
                                rs.Fields["units"].Value = f.Units;

                            rs.Fields["charge_currency_rcd"].Value = f.ChargeCurrencyRcd;
                            rs.Fields["charge_amount"].Value = f.ChargeAmount;
                            rs.Fields["charge_amount_incl"].Value = f.ChargeAmountIncl;
                            rs.Fields["document_number"].Value = f.DocumentNumber;

                            if (f.DocumentDateTime != DateTime.MinValue)
                                rs.Fields["document_date_time"].Value = f.DocumentDateTime;

                            rs.Fields["external_reference"].Value = f.ExternalReference;
                            rs.Fields["vendor_rcd"].Value = f.VendorRcd;

                            rs.Fields["weight_lbs"].Value = f.WeightLbs;
                            rs.Fields["weight_kgs"].Value = f.WeightKgs;
                            rs.Fields["mpd_number"].Value = f.MpdNumber;

                            //rs.Fields["unit_rcd"].Value = f.UnitRcd;
                        }

                    }

                }
                catch
                {
                    throw;
                }


            }
        }

        public static void ToRecordset(this  IList<Remark> remark, ref ADODB.Recordset rs)
        {
            if (remark != null && remark.Count > 0)
            {
                try
                {
                    foreach (Remark r in remark)
                    {
                        rs.AddNew();
                        if (r.BookingRemarkId != Guid.Empty)
                        {
                            rs.Fields["booking_remark_id"].Value = r.BookingRemarkId.ToRsString();
                        }
                        
                        rs.Fields["remark_type_rcd"].Value = r.RemarkTypeRcd;
                        rs.Fields["timelimit_date_time"].Value = r.TimelimitDateTime;
                        rs.Fields["complete_flag"].Value = r.CompleteFlag;
                        rs.Fields["remark_text"].Value = r.RemarkText;
                        if (r.BookingId != Guid.Empty)
                        {
                            rs.Fields["booking_id"].Value = r.BookingId.ToRsString();
                        }
                        rs.Fields["added_by"].Value = r.AddedBy; // string
                        if (r.ClientProfileId != Guid.Empty)
                        {
                            rs.Fields["client_profile_id"].Value = r.ClientProfileId.ToRsString();
                        }
                        rs.Fields["agency_code"].Value = r.AgencyCode;

                        if (r.CreateBy != Guid.Empty)
                            rs.Fields["create_by"].Value = r.CreateBy.ToRsString();

                        if (r.CreateDateTime != DateTime.MinValue)
                            rs.Fields["create_date_time"].Value = r.CreateDateTime;

                        if (r.UpdateBy != Guid.Empty)
                            rs.Fields["update_by"].Value = r.UpdateBy.ToRsString();

                        if (r.UpdateDateTime != DateTime.MinValue)
                            rs.Fields["update_date_time"].Value = r.UpdateDateTime;

                        rs.Fields["system_flag"].Value = r.SystemFlag;
                        rs.Fields["utc_timelimit_date_time"].Value = r.UtcTimelimitDateTime;
                        rs.Fields["nickname"].Value = r.Nickname;
                        rs.Fields["protected_flag"].Value = r.ProtectedFlag;
                        rs.Fields["warning_flag"].Value = r.WarningFlag;
                        rs.Fields["process_message_flag"].Value = r.ProcessMessageFlag;
                    }

                }
                catch
                {
                    throw;
                }


            }
        }
        public static void ToRecordset(this  IList<PassengerService> service, ref ADODB.Recordset rs)
        {
            if (service != null && service.Count > 0)
            {
                try
                {
                    foreach (PassengerService s in service)
                    {
                        rs.AddNew();
                        if (s.PassengerSegmentServiceId != Guid.Empty)
                        {
                            rs.Fields["passenger_segment_service_id"].Value = s.PassengerSegmentServiceId.ToRsString();
                        }
                        if (s.PassengerId != Guid.Empty)
                        {
                            rs.Fields["passenger_id"].Value = s.PassengerId.ToRsString();
                        }
                        if (s.BookingSegmentId != Guid.Empty)
                        {
                            rs.Fields["booking_segment_id"].Value = s.BookingSegmentId.ToRsString();
                        }

                        //EDWDEV-771
                        rs.Fields["special_service_rcd"].Value = s.SpecialServiceRcd;

                        if (s.ServiceOnRequestFlag == 1)
                            rs.Fields["special_service_status_rcd"].Value = "RQ";
                        else
                            rs.Fields["special_service_status_rcd"].Value = "HK";

                        rs.Fields["special_service_change_status_rcd"].Value = s.SpecialServiceChangeStatusRcd;

                        rs.Fields["service_text"].Value = s.ServiceText;

                        if (s.CreateBy != Guid.Empty)
                            rs.Fields["create_by"].Value = s.CreateBy.ToRsString();

                        if (s.CreateDateTime != DateTime.MinValue)
                            rs.Fields["create_date_time"].Value = s.CreateDateTime;

                        if (s.UpdateBy != Guid.Empty)
                            rs.Fields["update_by"].Value = s.UpdateBy.ToRsString();

                        if (s.UpdateDateTime != DateTime.MinValue)
                            rs.Fields["update_date_time"].Value = s.UpdateDateTime;

                        rs.Fields["number_of_units"].Value = s.NumberOfUnits;
                        rs.Fields["origin_rcd"].Value = s.OriginRcd;
                        rs.Fields["destination_rcd"].Value = s.DestinationRcd;

                        rs.Fields["unit_rcd"].Value = s.UnitRcd;


                    }

                }
                catch
                {
                    throw;
                }               
            }
        }

        //EDWDEV-771
        public static void ToRecordsetPEC(this IList<PassengerService> service, ref ADODB.Recordset rs)
        {
            if (service != null && service.Count > 0)
            {
                try
                {
                    foreach (PassengerService s in service)
                    {
                        rs.AddNew();
                        if (s.PassengerSegmentServiceId != Guid.Empty)
                        {
                            rs.Fields["passenger_segment_service_id"].Value = s.PassengerSegmentServiceId.ToRsString();
                        }
                        if (s.PassengerId != Guid.Empty)
                        {
                            rs.Fields["passenger_id"].Value = s.PassengerId.ToRsString();
                        }
                        if (s.BookingSegmentId != Guid.Empty)
                        {
                            rs.Fields["booking_segment_id"].Value = s.BookingSegmentId.ToRsString();
                        }

                        //EDWDEV-771
                        rs.Fields["special_service_rcd"].Value = s.SpecialServiceRcd;

                        rs.Fields["special_service_status_rcd"].Value = "HK";
                        rs.Fields["special_service_change_status_rcd"].Value = "KK";


                        rs.Fields["service_text"].Value = s.ServiceText;

                        if (s.CreateBy != Guid.Empty)
                            rs.Fields["create_by"].Value = s.CreateBy.ToRsString();

                        if (s.CreateDateTime != DateTime.MinValue)
                            rs.Fields["create_date_time"].Value = s.CreateDateTime;

                        if (s.UpdateBy != Guid.Empty)
                            rs.Fields["update_by"].Value = s.UpdateBy.ToRsString();

                        if (s.UpdateDateTime != DateTime.MinValue)
                            rs.Fields["update_date_time"].Value = s.UpdateDateTime;

                        rs.Fields["number_of_units"].Value = s.NumberOfUnits;
                        rs.Fields["origin_rcd"].Value = s.OriginRcd;
                        rs.Fields["destination_rcd"].Value = s.DestinationRcd;

                        rs.Fields["unit_rcd"].Value = s.UnitRcd;


                    }

                }
                catch
                {
                    throw;
                }
            }
        }


        public static void ToRecordset(this  IList<Tax> tax, ref ADODB.Recordset rs)
        {
            if (tax != null && tax.Count > 0)
            {
                try
                {
                    foreach (Tax t in tax)
                    {
                        rs.AddNew();
                        rs.Fields["passenger_segment_tax_id"].Value = Guid.NewGuid().ToRsString(); //t.PassengerSegmentTaxId.ToRsGuid();
                        rs.Fields["tax_amount"].Value = t.TaxAmount;
                        rs.Fields["tax_amount_incl"].Value = t.TaxAmountIncl;
                        rs.Fields["acct_amount"].Value = t.AcctAmount;
                        rs.Fields["acct_amount_incl"].Value = t.AcctAmountIncl;
                        rs.Fields["sales_amount"].Value = t.SalesAmount;
                        rs.Fields["sales_amount_incl"].Value = t.SalesAmountIncl;
                        rs.Fields["tax_rcd"].Value = t.TaxRcd;
                        if (t.PassengerId != Guid.Empty)
                        {
                            rs.Fields["passenger_id"].Value = t.PassengerId.ToRsString();
                        }
                        if (t.TaxId != Guid.Empty)
                        {
                            rs.Fields["tax_id"].Value = t.TaxId.ToRsString();
                        }
                        if (t.BookingSegmentId != Guid.Empty)
                        {
                            rs.Fields["booking_segment_id"].Value = t.BookingSegmentId.ToRsString();
                        }
                        rs.Fields["tax_currency_rcd"].Value = t.TaxCurrencyRcd;
                        rs.Fields["sales_currency_rcd"].Value = t.SalesCurrencyRcd;
                        rs.Fields["display_name"].Value = t.DisplayName;
                        rs.Fields["summarize_up"].Value = t.SummarizeUp;
                        rs.Fields["coverage_type"].Value = t.CoverageType;

                        if (t.CreateBy != Guid.Empty)
                            rs.Fields["create_by"].Value = t.CreateBy.ToRsString();

                        if (t.UpdateBy != Guid.Empty)
                            rs.Fields["update_by"].Value = t.UpdateBy.ToRsString();

                        if (t.CreateDateTime != DateTime.MinValue)
                            rs.Fields["create_date_time"].Value = t.CreateDateTime;
                        if (t.UpdateDateTime != DateTime.MinValue)
                            rs.Fields["update_date_time"].Value = t.UpdateDateTime;
                        rs.Fields["vat_percentage"].Value = t.VatPercentage;
                    }

                }
                catch
                {
                    throw;
                }
            }
        }


        public static void ToRecordset(this  Voucher v, ref ADODB.Recordset rs)
        {
            if (v != null )
            {
                try
                {
                    rs.AddNew();
                    if (v.VoucherId != null && v.VoucherId != "00000000-0000-0000-0000-000000000000")
                        rs.Fields["voucher_id"].Value = new Guid(v.VoucherId).ToRsString();

                    rs.Fields["payment_total"].Value = v.PaymentTotal;

                    if (v.CreateBy != Guid.Empty)
                        rs.Fields["create_by"].Value = v.CreateBy.ToRsString();
                    if (v.CreateDateTime != DateTime.MinValue)
                        rs.Fields["create_date_time"].Value = v.CreateDateTime;
                    if (v.UpdateBy != Guid.Empty)
                        rs.Fields["update_by"].Value = v.UpdateBy.ToRsString();
                    if (v.UpdateDateTime != DateTime.MinValue)
                        rs.Fields["update_date_time"].Value = v.UpdateDateTime;

                    rs.Fields["agency_code"].Value = v.AgencyCode;
                    rs.Fields["expiry_date_time"].Value = v.ExpiryDateTime;
                    rs.Fields["percentage_flag"].Value = v.PercentageFlag;
                    rs.Fields["refundable_flag"].Value = v.RefundableFlag;
                    rs.Fields["single_flight_flag"].Value = v.SingleFlightFlag;
                    rs.Fields["single_passenger_flag"].Value = v.SinglePassengerFlag;
                    rs.Fields["recipient_name"].Value = v.RecipientName;
                    rs.Fields["voucher_number"].Value = v.VoucherNumber;
                    rs.Fields["voucher_password"].Value = v.VoucherPassword;
                    rs.Fields["voucher_status_rcd"].Value = v.VoucherStatusRcd;
                    rs.Fields["voucher_value"].Value = v.VoucherValue;
                }
                catch
                {
                    throw;
                }
            }
        }

        // for save payment
        public static void ToRecordset(this  IList<PaymentAllocation> paymentAllocation, ref ADODB.Recordset rs)
        {
            if (paymentAllocation != null && paymentAllocation.Count > 0)
            {
                try
                {
                    foreach (PaymentAllocation p in paymentAllocation)
                    {
                        rs.AddNew();
                        if (p.booking_payment_id != Guid.Empty)
                            rs.Fields["booking_payment_id"].Value = p.booking_payment_id.ToRsString();

                        if (p.passenger_id != Guid.Empty)
                            rs.Fields["passenger_id"].Value = p.passenger_id.ToRsString();

                        if (p.booking_segment_id != Guid.Empty)
                            rs.Fields["booking_segment_id"].Value = p.booking_segment_id.ToRsString();

                        if (p.booking_fee_id != Guid.Empty)
                            rs.Fields["booking_fee_id"].Value = p.booking_fee_id.ToRsString();

                        if (p.voucher_id != Guid.Empty)
                            rs.Fields["voucher_id"].Value = p.voucher_id.ToRsString();

                        if (p.passenger_segment_service_id != Guid.Empty)
                            rs.Fields["passenger_segment_service_id"].Value = p.passenger_segment_service_id.ToRsString();

                        if (p.fee_id != Guid.Empty)
                            rs.Fields["fee_id"].Value = p.fee_id.ToRsString();

                        if (p.user_id != Guid.Empty)
                            rs.Fields["user_id"].Value = p.user_id.ToRsString();

                        rs.Fields["currency_rcd"].Value = p.currency_rcd;
                        rs.Fields["charge_currency_rcd"].Value = p.charge_currency_rcd;

                        // BAGSALES
                        if (string.IsNullOrEmpty(p.fee_category_rcd) == false)
                        {
                            if (p.fee_category_rcd.Contains("BAG"))
                                rs.Fields["fee_category_rcd"].Value = "BAGF";
                            else if (p.fee_category_rcd.Length > 4)
                                rs.Fields["fee_category_rcd"].Value = p.fee_category_rcd.Substring(0, 4);
                            else
                                rs.Fields["fee_category_rcd"].Value = p.fee_category_rcd;
                        }

                        rs.Fields["vendor_rcd"].Value = p.vendor_rcd;

                        if (!string.IsNullOrEmpty(p.units))
                            rs.Fields["units"].Value = p.units;

                        rs.Fields["external_reference"].Value = p.external_reference;

                        rs.Fields["sales_amount"].Value = p.sales_amount;
                        rs.Fields["payment_amount"].Value = p.payment_amount;
                        rs.Fields["account_amount"].Value = p.account_amount;
                        rs.Fields["charge_amount"].Value = p.charge_amount;
                        rs.Fields["charge_amount_incl"].Value = p.charge_amount_incl;
                        rs.Fields["weight_lbs"].Value = p.weight_lbs;
                        rs.Fields["weight_kgs"].Value = p.weight_kgs;

                    }

                }
                catch
                {
                    throw;
                } 

            }
        }

        public static void SetRSnewSegmentId(this  List<KeyValuePair<string, string>> objSegmentIdMapping, ref ADODB.Recordset rs)
        {
            if (objSegmentIdMapping != null && objSegmentIdMapping.Count > 0)
            {
                foreach (var element in objSegmentIdMapping)
                {
                    if (rs != null && rs.RecordCount > 0)
                    {
                        rs.MoveFirst();
                        while (!rs.EOF)
                        {
                            if (RecordsetHelper.ToGuid(rs, "booking_segment_id") == new Guid(element.Key))
                            {
                                rs.Fields["booking_segment_id"].Value = "{" + element.Value + "}";
                            }
                            rs.MoveNext();
                        }
                    }
                }
            }

        }

        public static void SetSeat(this IList<Mapping> mappings, ref ADODB.Recordset rsMapping)
        {
            foreach (Entity.Booking.Mapping m in mappings)
            {
                rsMapping.MoveFirst();
                while (!rsMapping.EOF)
                {
                    if (RecordsetHelper.ToGuid(rsMapping, "booking_segment_id") == m.BookingSegmentId &&
                        RecordsetHelper.ToGuid(rsMapping, "passenger_id") == m.PassengerId && 
                        m.SeatNumber.Length > 0 && m.SeatFeeRcd.Length > 0)
                    {
                        rsMapping.Fields["seat_column"].Value = m.SeatColumn;
                        rsMapping.Fields["seat_row"].Value = m.SeatRow;
                        rsMapping.Fields["seat_number"].Value = m.SeatNumber;

                        if (m.SeatFeeRcd != null)
                            rsMapping.Fields["seat_fee_rcd"].Value = m.SeatFeeRcd;

                        rsMapping.Fields["update_date_time"].Value = DateTime.Now;
                        break;
                    }
                    rsMapping.MoveNext();
                }
            }

        }

        public static void SetNewSegmentId(this IList<PassengerService> services, List<KeyValuePair<string, string>> objSegmentIdMapping)
        {
            if (objSegmentIdMapping != null && objSegmentIdMapping.Count > 0)
            {
                foreach (Entity.Booking.PassengerService ps in services)
                {
                    foreach (var element in objSegmentIdMapping)
                    {
                        if (ps.BookingSegmentId == new Guid(element.Key))
                        {
                            ps.BookingSegmentId = new Guid(element.Value);
                        }
                    }
                }
            }

        }
    }
}
