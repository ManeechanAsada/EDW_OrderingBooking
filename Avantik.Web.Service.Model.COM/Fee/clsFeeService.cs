using Avantik.Web.Service.Model.Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.Entity.Agency;
using Avantik.Web.Service.COMHelper;
using System.Runtime.InteropServices;
using Avantik.Web.Service.Model.COM.Extension;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Entity.Booking;
using ADODB;
using Avantik.Web.Service.Helpers;

namespace Avantik.Web.Service.Model.COM
{
    public class FeeService : RunComplus, IFeeService
    {
        string _server = string.Empty;
        public FeeService(string server, string user, string pass, string domain)
            :base(user,pass,domain)
        {
            _server = server;
        }

        public IList<Entity.Fee> GetFee(string strFeeRcd,
                          string strCurrencyCode,
                          string strAgencyCode,
                          string strClass,
                          string strFareBasis,
                          string strOrigin,
                          string strDestination,
                          string strFlightNumber,
                          DateTime dtDate,
                          string strLanguage,
                          bool bNovat)
        {
            tikAeroProcess.Settings objSettings = null;
            ADODB.Recordset rs = null;
            IList<Entity.Fee> fees = new List<Entity.Fee>();
            try
            {
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Settings", _server);
                    objSettings = (tikAeroProcess.Settings)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objSettings = new tikAeroProcess.Settings(); }

                rs = objSettings.GetFee(strFeeRcd,
                                        strCurrencyCode,
                                        strAgencyCode,
                                        strClass,
                                        strFareBasis,
                                        strOrigin,
                                        strDestination,
                                        strFlightNumber,
                                        dtDate,
                                        strLanguage, bNovat);

                //convert Recordset to Object
                if (rs != null && rs.RecordCount > 0)
                {
                    Entity.Fee f;
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        f = new Entity.Fee();

                        f.FeeId = RecordsetHelper.ToGuid(rs, "fee_id");
                        f.FeeRcd = RecordsetHelper.ToString(rs, "fee_rcd");
                        f.CurrencyRcd = RecordsetHelper.ToString(rs, "currency_rcd");
                        f.FeeAmount = RecordsetHelper.ToDecimal(rs, "fee_amount");
                        f.FeeAmountIncl = RecordsetHelper.ToDecimal(rs, "fee_amount_incl");
                        f.VatPercentage = RecordsetHelper.ToDecimal(rs, "vat_percentage");
                        f.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        f.FeeCategoryRcd = RecordsetHelper.ToString(rs, "fee_category_rcd");
                        f.MinimumFeeAmountFlag = RecordsetHelper.ToByte(rs, "minimum_fee_amount_flag");
                        f.FeePercentage = RecordsetHelper.ToDecimal(rs, "fee_percentage");
                        f.OdFlag = RecordsetHelper.ToByte(rs, "od_flag");

                        fees.Add(f);
                        f = null;
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
                if (objSettings != null)
                {
                    Marshal.FinalReleaseComObject(objSettings);
                    objSettings = null;
                }

                RecordsetHelper.ClearRecordset(ref rs);
            }

            return fees;
        }

        public IList<Entity.Booking.Fee> CalculateFeesBookingCreate(string AgencyCode,
                                                        string currency,
                                                        Booking booking,
                                                        string strLanguage) 
        {
            tikAeroProcess.Fees objFees = null;
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;

            ADODB.Recordset rsFlight = null;
            ADODB.Recordset rsQuote = null;

            string strBookingId = string.Empty;
            string strOther = string.Empty;
            string strUserId = string.Empty;
            string strAgencyCode = string.Empty;
            string strCurrency = string.Empty;

            short iAdult = 0;
            short iChild = 0;
            short iInfant = 0;
            short iOther = 0;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objFees = new tikAeroProcess.Fees(); }

                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                //Get Empty Recordset
                objBooking.GetEmpty(ref rsHeader,
                                    ref rsSegment,
                                    ref rsPassenger,
                                    ref rsRemark,
                                    ref rsPayment,
                                    ref rsMapping,
                                    ref rsService,
                                    ref rsTax,
                                    ref rsFees,
                                    ref rsFlight,
                                    ref rsQuote,
                                    ref strBookingId,
                                    ref strAgencyCode,
                                    ref strCurrency,
                                    ref iAdult,
                                    ref iChild,
                                    ref iInfant,
                                    ref iOther,
                                    ref strOther,
                                    ref strUserId);
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }

                if (booking.Header != null)
                {
                    booking.Header.ToRecordset(ref rsHeader);
                }
                if (booking.Segments != null && booking.Segments.Count > 0)
                {
                    booking.Segments.ToRecordset(ref rsSegment);
                }
                if (booking.Passengers != null && booking.Passengers.Count > 0)
                {
                    booking.Passengers.ToRecordset(ref rsPassenger);
                }
                if (booking.Fees != null && booking.Fees.Count > 0)
                {
                  //  booking.Fees.ToRecordset(ref rsFees);
                }
                if (booking.Remarks != null && booking.Remarks.Count > 0)
                {
                    booking.Remarks.ToRecordset(ref rsRemark);
                }
                if (booking.Mappings != null && booking.Mappings.Count > 0)
                {
                    booking.Mappings.ToRecordset(ref rsMapping);
                }
                if (booking.Services != null && booking.Services.Count > 0)
                {
                    booking.Services.ToRecordset(ref rsService);
                }

                objFees.BookingCreate(AgencyCode,
                                        currency,
                                        booking.Header.BookingId.ToString(),
                                        ref rsHeader,
                                        ref rsSegment,
                                        ref rsPassenger,
                                        ref rsFees,
                                        ref rsRemark,
                                        ref rsMapping,
                                        ref rsService,
                                        ref strLanguage);

                if (rsFees != null && rsFees.RecordCount > 0)
                {
                    booking.Fees = booking.Fees.FillBooking(rsFees);
                }

            }
            catch
            { }
            finally
            {
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }

                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
            }
            return booking.Fees;
        }


        public IList<Entity.Booking.Fee> CalculateFeesBookingChange(string AgencyCode,
                                                    string currency,
                                                    Booking booking,
                                                    string strLanguage)
        {
            tikAeroProcess.Fees objFees = null;
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;

            ADODB.Recordset rsFlight = null;
            ADODB.Recordset rsQuote = null;

            string strBookingId = string.Empty;
            string strOther = string.Empty;
            string strUserId = string.Empty;
            string strAgencyCode = string.Empty;
            string strCurrency = string.Empty;

            short iAdult = 0;
            short iChild = 0;
            short iInfant = 0;
            short iOther = 0;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objFees = new tikAeroProcess.Fees(); }

                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                //Get Empty Recordset
                objBooking.GetEmpty(ref rsHeader,
                                    ref rsSegment,
                                    ref rsPassenger,
                                    ref rsRemark,
                                    ref rsPayment,
                                    ref rsMapping,
                                    ref rsService,
                                    ref rsTax,
                                    ref rsFees,
                                    ref rsFlight,
                                    ref rsQuote,
                                    ref strBookingId,
                                    ref strAgencyCode,
                                    ref strCurrency,
                                    ref iAdult,
                                    ref iChild,
                                    ref iInfant,
                                    ref iOther,
                                    ref strOther,
                                    ref strUserId);
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }

                if (booking.Header != null)
                {
                    booking.Header.ToRecordset(ref rsHeader);
                }
                if (booking.Segments != null && booking.Segments.Count > 0)
                {
                    booking.Segments.ToRecordset(ref rsSegment);
                }
                if (booking.Passengers != null && booking.Passengers.Count > 0)
                {
                    booking.Passengers.ToRecordset(ref rsPassenger);
                }
                if (booking.Fees != null && booking.Fees.Count > 0)
                {
                  //  booking.Fees.ToRecordset(ref rsFees);
                }
                if (booking.Remarks != null && booking.Remarks.Count > 0)
                {
                    booking.Remarks.ToRecordset(ref rsRemark);
                }
                if (booking.Mappings != null && booking.Mappings.Count > 0)
                {
                    booking.Mappings.ToRecordset(ref rsMapping);
                }
                if (booking.Services != null && booking.Services.Count > 0)
                {
                    booking.Services.ToRecordset(ref rsService);
                }

                objFees.BookingChange(AgencyCode,
                                        currency,
                                        booking.Header.BookingId.ToString(),
                                        ref rsHeader,
                                        ref rsSegment,
                                        ref rsPassenger,
                                        ref rsFees,
                                        ref rsRemark,
                                        ref rsMapping,
                                        ref rsService,
                                        ref strLanguage);

                if (rsFees != null && rsFees.RecordCount > 0)
                {
                    booking.Fees = booking.Fees.FillBooking(rsFees);
                }

            }
            catch
            { }
            finally
            {
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }

                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
            }

            return booking.Fees;
        }

        public IList<Entity.Booking.Fee> CalculateFeesSeatAssignment(string AgencyCode,
                                                        string currency,
                                                        Booking booking,
                                                        string strLanguage,
                                                        bool bNovat)
        {
            tikAeroProcess.Fees objFees = null;
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;

            ADODB.Recordset rsFlight = null;
            ADODB.Recordset rsQuote = null;

            string strBookingId = string.Empty;
            string strOther = string.Empty;
            string strUserId = string.Empty;
            string strAgencyCode = string.Empty;
            string strCurrency = string.Empty;
            short iAdult = 0;
            short iChild = 0;
            short iInfant = 0;
            short iOther = 0;

            IList<Entity.Booking.Fee> Fees = new List<Entity.Booking.Fee>();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objFees = new tikAeroProcess.Fees(); }

                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                //Get Empty Recordset
                objBooking.GetEmpty(ref rsHeader,
                                    ref rsSegment,
                                    ref rsPassenger,
                                    ref rsRemark,
                                    ref rsPayment,
                                    ref rsMapping,
                                    ref rsService,
                                    ref rsTax,
                                    ref rsFees,
                                    ref rsFlight,
                                    ref rsQuote,
                                    ref strBookingId,
                                    ref strAgencyCode,
                                    ref strCurrency,
                                    ref iAdult,
                                    ref iChild,
                                    ref iInfant,
                                    ref iOther,
                                    ref strOther,
                                    ref strUserId);
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }

                if (booking.Header != null)
                {
                    booking.Header.ToRecordset(ref rsHeader);
                }
                if (booking.Mappings != null && booking.Mappings.Count > 0)
                {
                    booking.Mappings.ToRecordset(ref rsMapping);
                }

                foreach (Entity.Booking.Mapping m in booking.Mappings)
                {
                    rsMapping.MoveFirst();
                    while (!rsMapping.EOF)
                    {
                        if (RecordsetHelper.ToGuid(rsMapping, "booking_segment_id").ToString().ToUpper().Equals(m.BookingSegmentId.ToString().ToUpper()) &&
                            
                            RecordsetHelper.ToGuid(rsMapping, "passenger_id").ToString().ToUpper().Equals(m.PassengerId.ToString().ToUpper())
                            )
                        {
                            rsMapping.Fields["seat_column"].Value = m.SeatColumn;
                            rsMapping.Fields["seat_row"].Value = m.SeatRow;
                            rsMapping.Fields["seat_number"].Value = m.SeatNumber;

                            // revise for fix seat fee is null return wrong fee
                            // sent seat_fee_rcd = null or empty return fee always
                            // should be set wrong feercd
                            if (string.IsNullOrEmpty(m.SeatFeeRcd))
                                 rsMapping.Fields["seat_fee_rcd"].Value = "cba.yxz";
                            else
                                rsMapping.Fields["seat_fee_rcd"].Value = m.SeatFeeRcd;

                            rsMapping.Fields["update_date_time"].Value = DateTime.Now;
                            break;
                        }
                        rsMapping.MoveNext();
                    }
                }
                objFees.SeatAssignment(AgencyCode,
                                        currency, 
                                        booking.Header.BookingId.ToString(),
                                        ref rsHeader,
                                        ref rsSegment,
                                        ref rsPassenger,
                                        ref rsFees,
                                        ref rsRemark,
                                        ref rsMapping,
                                        ref rsService,
                                        ref strLanguage,
                                        ref bNovat);

                if (rsFees != null && rsFees.RecordCount > 0)
                {
                   Fees = Fees.FillBooking(rsFees);
                }

            }
            catch
            { }
            finally
            {
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }

                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
            }

            IList<Entity.Booking.Fee> fees = new List<Entity.Booking.Fee>();
            return Fees;
        }

        public IList<Entity.Booking.Fee> CalculateFeesNameChange(string AgencyCode,
                                                    string currency,
                                                    Booking booking,
                                                    string strLanguage,
                                                    string strUserId)
        {
            tikAeroProcess.Fees objFees = null;
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;

            ADODB.Recordset rsQuote = null;

            string strBookingId = string.Empty;
            string strOther = string.Empty;
            string strAgencyCode = string.Empty;
            string strCurrency = string.Empty;

            strBookingId = booking.Header.BookingId.ToString();

            IList<Entity.Booking.Fee> Fees = new List<Entity.Booking.Fee>();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objFees = new tikAeroProcess.Fees(); }

                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                if (objBooking.Read("{" + strBookingId.ToUpper() + "}",
                                    string.Empty,
                                    0,
                                    ref rsHeader,
                                    ref rsSegment,
                                    ref rsPassenger,
                                    ref rsRemark,
                                    ref rsPayment,
                                    ref rsMapping,
                                    ref rsService,
                                    ref rsTax,
                                    ref rsFees,
                                    ref rsQuote,
                                    false,
                                    false,
                                    "{" + strUserId.ToUpper() + "}",
                                    true,
                                    true,
                                    true,
                                    true,
                                    true,
                                    true,
                                    true,
                                    true,
                                    true,
                                    true,
                                    string.Empty,
                                    string.Empty) == true)
                {

                    // need to clear fee from old booking that not paid
                     RecordsetHelper.DeleteFees(ref rsFees);


                    foreach (Mapping m in booking.Mappings)
                    {
                        rsMapping.MoveFirst();
                        while (!rsMapping.EOF)
                        {
                            if (m.PassengerId == RecordsetHelper.ToGuid(rsMapping, "passenger_id"))
                            {
                                rsMapping.Fields["title_rcd"].Value = m.TitleRcd;
                                rsMapping.Fields["firstname"].Value = m.Firstname;
                                rsMapping.Fields["lastname"].Value = m.Lastname;
                                rsMapping.Fields["date_of_birth"].Value = m.DateOfBirth;
                            }
                            rsMapping.MoveNext();
                        }
                    }

                    foreach (Passenger p in booking.Passengers)
                    {
                        rsPassenger.MoveFirst();
                        while (!rsPassenger.EOF)
                        {
                            if (p.PassengerId == RecordsetHelper.ToGuid(rsPassenger, "passenger_id"))
                            {
                                rsPassenger.Fields["title_rcd"].Value = p.TitleRcd;
                                rsPassenger.Fields["firstname"].Value = p.Firstname;
                                rsPassenger.Fields["middlename"].Value = p.Middlename;
                                rsPassenger.Fields["lastname"].Value = p.Lastname;
                                rsPassenger.Fields["date_of_birth"].Value = p.DateOfBirth;

                                break;
                            }
                            rsPassenger.MoveNext();
                        }
                    }


                }

                objFees.NameChange(AgencyCode,
                                        currency,
                                        booking.Header.BookingId.ToString(),
                                        ref rsHeader,
                                        ref rsSegment,
                                        ref rsPassenger,
                                        ref rsFees,
                                        ref rsRemark,
                                        ref rsMapping,
                                        ref rsService,
                                        ref strLanguage,
                                        false);

                if (rsFees != null && rsFees.RecordCount > 0)
                {
                    Fees = Fees.FillNameChangeFee(rsFees);
                }

            }
            catch
            { }
            finally
            {
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }

                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
            }

            return Fees;
        }

        //ssr
        public Entity.Booking.Booking CalculateFeesSpecialService(string AgencyCode,
                                                        string currency,
                                                        Booking booking,
                                                        string strLanguage,
                                                        bool bNovat)
        {
            tikAeroProcess.Fees objFees = null;
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;

            ADODB.Recordset rsFlight = null;
            ADODB.Recordset rsQuote = null;
            Booking bookingResponse = new Booking();

            string strBookingId = string.Empty;
            string strOther = string.Empty;
            string strUserId = string.Empty;
            string strAgencyCode = string.Empty;
            string strCurrency = string.Empty;

            short iAdult = 0;
            short iChild = 0;
            short iInfant = 0;
            short iOther = 0;


            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objFees = new tikAeroProcess.Fees(); }

                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }
                
                //Get Empty Recordset
                objBooking.GetEmpty(ref rsHeader,
                                    ref rsSegment,
                                    ref rsPassenger,
                                    ref rsRemark,
                                    ref rsPayment,
                                    ref rsMapping,
                                    ref rsService,
                                    ref rsTax,
                                    ref rsFees,
                                    ref rsFlight,
                                    ref rsQuote,
                                    ref strBookingId,
                                    ref strAgencyCode,
                                    ref strCurrency,
                                    ref iAdult,
                                    ref iChild,
                                    ref iInfant,
                                    ref iOther,
                                    ref strOther,
                                    ref strUserId);
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }


                if (booking.Header != null)
                {
                    booking.Header.ToRecordset(ref rsHeader);
                }
                if (booking.Mappings != null && booking.Mappings.Count > 0)
                {
                    booking.Mappings.ToRecordset(ref rsMapping);
                }
                if (booking.Fees != null && booking.Fees.Count > 0)
                {
                  //  booking.Fees.ToRecordset(ref rsFees);
                }

                if (booking.Services != null && booking.Services.Count > 0)
                {
                    booking.Services.ToRecordset(ref rsService);
                }
                // clear fee
                // RecordsetHelper.DeleteFees(ref rsFees);

                objFees.SpecialService(AgencyCode,
                                        currency,
                                        booking.Header.BookingId.ToString(),
                                        ref rsHeader,
                                        ref rsService,
                                        ref rsFees,
                                        ref rsRemark,
                                        ref rsMapping,
                                        ref strLanguage,
                                        ref bNovat);

                if (rsFees != null && rsFees.RecordCount > 0)
                {
                    bookingResponse.Fees = bookingResponse.Fees.FillBooking(rsFees);
                }
                if (rsService != null && rsService.RecordCount > 0)
                {
                    bookingResponse.Services = bookingResponse.Services.FillBooking(rsService);
                }

             //   Logger.SaveLog("3", DateTime.Now, DateTime.Now, "xxxx", XMLHelper.Serialize(bookingResponse.Fees, false));
            }
            catch
            { }
            finally
            {
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }

                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
            }

            return bookingResponse;
        }

        public List<ServiceFee> GetSegmentFee(string agencyCode,
                                       string currencyCode,
                                       string languageCode,
                                       int numberOfPassenger,
                                       int numberOfInfant,
                                       IList<PassengerService> services,
                                       IList<SegmentService> segmentService,
                                       bool SpecialService, 
                                       bool bNovat)
        {
            List<ServiceFee> fees = new List<ServiceFee>();
            tikAeroProcess.Fees objFees = null;
            ADODB.Recordset rs = null;
            try
            {
                if (segmentService != null && segmentService.Count > 0 && services != null && services.Count > 0)
                {
                    bool b = false;
                    //Declaration
                    if (string.IsNullOrEmpty(_server) == false)
                    {
                        Type remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                        objFees = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                        remote = null;
                    }
                    else
                    { objFees = new tikAeroProcess.Fees(); }

                    //Implement Segment Fee function
                    ADODB.Recordset rsSegment;
                    //Fill Segment information
                    for (int f = 0; f < services.Count; f++)
                    {
                        //Create new segmentFee instance.
                        rsSegment = objFees.GetSegmentFeeRecordset();
                        for (int i = 0; i < segmentService.Count; i++)
                        {
                            rsSegment.AddNew();
                            if (segmentService[i].FlightConnectionId.Equals(Guid.Empty) == false)
                            {
                                rsSegment.Fields["flight_connection_id"].Value = "{" + segmentService[i].FlightConnectionId.ToString() + "}";
                            }
                            rsSegment.Fields["origin_rcd"].Value = segmentService[i].OriginRcd;
                            rsSegment.Fields["destination_rcd"].Value = segmentService[i].DestinationRcd;
                            rsSegment.Fields["od_origin_rcd"].Value = segmentService[i].OdOriginRcd;
                            rsSegment.Fields["od_destination_rcd"].Value = segmentService[i].OdDestinationRcd;
                            rsSegment.Fields["booking_class_rcd"].Value = segmentService[i].BookingClassRcd;
                            rsSegment.Fields["fare_code"].Value = segmentService[i].FareCode;
                            rsSegment.Fields["airline_rcd"].Value = segmentService[i].AirlineRcd;
                            rsSegment.Fields["flight_number"].Value = segmentService[i].FlightNumber;
                            rsSegment.Fields["departure_date"].Value = segmentService[i].DepartureDate;
                            rsSegment.Fields["pax_count"].Value = numberOfPassenger;
                            rsSegment.Fields["inf_count"].Value = numberOfInfant;
                        }
                        //Call COM plus function
                        if (SpecialService == true)
                        {
                            b = objFees.SpecialServiceFee(agencyCode, currencyCode, services[f].SpecialServiceRcd, ref rsSegment, languageCode, bNovat);
                        }
                        else
                        {
                            b = objFees.SegmentFee(agencyCode, currencyCode, ref rsSegment, services[f].SpecialServiceRcd, languageCode);
                        }

                        ServiceFee sf;
                        if (b == true)
                        {
                            if (rsSegment != null && rsSegment.RecordCount > 0)
                            {
                                rsSegment.MoveFirst();
                                while (!rsSegment.EOF)
                                {
                                    sf = new ServiceFee();
                                    sf.FeeRcd = services[f].SpecialServiceRcd;
                                    sf.SpecialServiceRcd = services[f].SpecialServiceRcd;
                                    sf.DisplayName = RecordsetHelper.ToString(rsSegment, "display_name");
                                    sf.ServiceOnRequestFlag = Convert.ToBoolean(services[f].ServiceOnRequestFlag);
                                    sf.CutOffTime = Convert.ToBoolean(services[f].CutOffTime);
                                    sf.OriginRcd = RecordsetHelper.ToString(rsSegment, "origin_rcd");
                                    sf.DestinationRcd = RecordsetHelper.ToString(rsSegment, "destination_rcd");
                                    sf.OdOriginRcd = RecordsetHelper.ToString(rsSegment, "od_origin_rcd");
                                    sf.OdDestinationRcd = RecordsetHelper.ToString(rsSegment, "od_destination_rcd");
                                    sf.BookingClassRcd = RecordsetHelper.ToString(rsSegment, "booking_class_rcd");
                                    sf.FareCode = RecordsetHelper.ToString(rsSegment, "fare_code");
                                    sf.AirlineRcd = RecordsetHelper.ToString(rsSegment, "airline_rcd");
                                    sf.FlightNumber = RecordsetHelper.ToString(rsSegment, "flight_number");
                                    sf.DepartureDate = RecordsetHelper.ToDateTime(rsSegment, "departure_date");
                                    sf.FeeAmount = RecordsetHelper.ToDecimal(rsSegment, "fee_amount");
                                    sf.FeeAmountIncl = RecordsetHelper.ToDecimal(rsSegment, "fee_amount_incl");
                                    sf.TotalFeeAmount = RecordsetHelper.ToDecimal(rsSegment, "total_fee_amount");
                                    sf.TotalFeeAmountIncl = RecordsetHelper.ToDecimal(rsSegment, "total_fee_amount_incl");
                                    sf.CurrencyRcd = RecordsetHelper.ToString(rsSegment, "currency_rcd");
                                    rsSegment.MoveNext();

                                    if (!string.IsNullOrEmpty(sf.CurrencyRcd))
                                    {
                                        fees.Add(sf);
                                    }
                                }

                            }
                        }
                        RecordsetHelper.ClearRecordset(ref rsSegment, true);
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (objFees != null)
                {
                    Marshal.FinalReleaseComObject(objFees);
                    objFees = null;
                }
                RecordsetHelper.ClearRecordset(ref rs, true);
            }

            return fees;
        }


        public IList<Entity.Booking.Fee> GetBaggageFee(
            IList<Entity.Booking.Mapping> mapping, 
            Guid bookingSegmentId, 
            Guid passengerId, 
            string agencyCode, 
            string languageCode, 
            int maxUnits, 
            IList<Entity.Booking.Fee> fees, 
            bool bNovat)
        {
            tikAeroProcess.Fees objFee = null;
            tikAeroProcess.Booking objBooking = null;

            ADODB.Recordset rsBagFee = null;



            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;

            ADODB.Recordset rsFlight = null;
            ADODB.Recordset rsQuote = null;

            string strBookingId = string.Empty;
            string strOther = string.Empty;
            string strUserId = string.Empty;
            string strAgencyCode = string.Empty;
            string strCurrency = string.Empty;

            short iAdult = 0;
            short iChild = 0;
            short iInfant = 0;
            short iOther = 0;

            IList<Entity.Booking.Fee> baggageFees = new List<Entity.Booking.Fee>();

            try
            {
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Fees", _server);
                    objFee = (tikAeroProcess.Fees)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objFee = new tikAeroProcess.Fees(); }

                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objBooking = new tikAeroProcess.Booking(); }

                //Get Empty Recordset
                objBooking.GetEmpty(ref rsHeader,
                                    ref rsSegment,
                                    ref rsPassenger,
                                    ref rsRemark,
                                    ref rsPayment,
                                    ref rsMapping,
                                    ref rsService,
                                    ref rsTax,
                                    ref rsFees,
                                    ref rsFlight,
                                    ref rsQuote,
                                    ref strBookingId,
                                    ref strAgencyCode,
                                    ref strCurrency,
                                    ref iAdult,
                                    ref iChild,
                                    ref iInfant,
                                    ref iOther,
                                    ref strOther,
                                    ref strUserId);
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }

                if (mapping != null && mapping.Count > 0)
                {
                    string strPassengerId = string.Empty;
                    if (passengerId.Equals(Guid.Empty) == false)
                    {
                        strPassengerId = passengerId.ToString();
                    }

                    mapping.ToRecordset(ref rsMapping);

                    rsBagFee = objFee.GetBaggageFeeOptions(rsMapping, bookingSegmentId.ToString(), strPassengerId, agencyCode, languageCode, maxUnits, rsFees, bNovat);

                    if (rsBagFee != null && rsBagFee.RecordCount > 0)
                    {
                        baggageFees.FillBaggage(ref rsBagFee);
                    }
                }

            }
            catch (SystemException ex)
            {
                throw ex;
            }
            finally
            {
                if (objFee != null)
                {
                    Marshal.FinalReleaseComObject(objFee);
                    objFee = null;
                }
                RecordsetHelper.ClearRecordset(ref rsBagFee);
                RecordsetHelper.ClearRecordset(ref rsMapping, false);
                RecordsetHelper.ClearRecordset(ref rsFees);
            }
            return baggageFees;
        }
    }
}
