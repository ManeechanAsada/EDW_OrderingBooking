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
using Avantik.Web.Service.Exception.Booking;
using Avantik.Web.Service.Entity.Booking;
using Avantik.Web.Service.Entity;
using ADODB;
using System.Data;

namespace Avantik.Web.Service.Model.COM
{
    public class PaymentService : RunComplus, IPaymentService
    {
        string _server = string.Empty;
        string _user = string.Empty;
        string _pass = string.Empty;
        string _domain = string.Empty;
        public PaymentService(string server, string user, string pass, string domain)
            : base(user, pass, domain)
        {
            _server = server;
            _user = user;
            _pass = pass;
            _domain = domain;
        }

        public bool SavePayment(IList<Payment> payments, IList<PaymentAllocation> paymentAllocations)
        {
            tikAeroProcess.Payment objPayment = null;

            Recordset rsPayment = null;
            Recordset rsAllocation = null;
            Recordset rsRefundVoucher = null;
            bool result = false;

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


                // get empty rs object
                rsPayment = objPayment.GetEmpty();

                rsAllocation = RecordsetHelper.FabricatePaymentAllocationRecordset();

                //fill rs
                if (payments != null && payments.Count > 0)
                {
                    payments.ToRecordset(ref rsPayment);
                }
                if (paymentAllocations != null && paymentAllocations.Count > 0)
                {
                    paymentAllocations.ToRecordset(ref rsAllocation);
                }

                result = objPayment.Save(ref rsPayment, ref rsAllocation, ref rsRefundVoucher);

            }
            catch (System.Runtime.InteropServices.COMException exc)
            {
                throw exc;
            }
            catch (BookingException ex)
            {
                throw ex;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (objPayment != null)
                { Marshal.FinalReleaseComObject(objPayment); }
                objPayment = null;

                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsAllocation);
                RecordsetHelper.ClearRecordset(ref rsRefundVoucher);
            }

            return result;

        }

        public bool PaymentCreditCard(Entity.Booking.BookingHeader header,
                                      IList<Entity.Booking.Payment> payments,
                                      IList<Entity.PaymentAllocation> paymentAllocations)
        {
            tikAeroProcess.Booking objBooking = null;
            tikAeroProcess.clsCreditCard objCreditCard = null;
            tikAeroProcess.Payment objPayment = null;

            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsPayment2 = null;

            ADODB.Recordset rsTempPayment = null;
            ADODB.Recordset rsMapping = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsQuote = null;
            ADODB.Recordset rsAllocation = null;
            ADODB.Recordset rsVoucher = null;

            string strBookingId = header.BookingId.ToString();
            string strOther = string.Empty;
            string strUserId = header.UpdateBy.ToString();
            string strAgencyCode = string.Empty;
            string strCurrency = string.Empty;


            string securityToken = string.Empty;
            string authenticationToken = string.Empty;
            string commerceIndicator = string.Empty;
            string bookingReference = string.Empty;
            string requestSource = string.Empty;
            bool bWeb = false;
            System.Data.DataSet ds = null;

            bool bResult = false;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.clsCreditCard", _server);
                    objCreditCard = (tikAeroProcess.clsCreditCard)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                {
                    objBooking = new tikAeroProcess.Booking();
                    objCreditCard = new tikAeroProcess.clsCreditCard();
                    objPayment = new tikAeroProcess.Payment();
                }

                // get empty rs object
                rsPayment = objPayment.GetEmpty();

                payments.ToRecordset(ref rsPayment);
                bookingReference = header.RecordLocator;
                rsAllocation = RecordsetHelper.FabricatePaymentAllocationRecordset();

                paymentAllocations.ToRecordset(ref rsAllocation);

                 if (objBooking.Read("{" + strBookingId.ToUpper() + "}",
                                    string.Empty,
                                    0,
                                    ref rsHeader,
                                    ref rsSegment,
                                    ref rsPassenger,
                                    ref rsRemark,
                                    ref rsPayment2,
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
                     // *********** if need   should be sent fee and mapping fill to RS before Authorize ************ 

                    //Call to validate credit card payment.
                    ADODB.Recordset rs = objCreditCard.Authorize(rsPayment,
                                                                rsSegment,
                                                                ref rsAllocation,
                                                                ref rsVoucher,
                                                                ref rsMapping,
                                                                ref rsFees,
                                                                ref rsHeader,
                                                                ref rsTax,
                                                                ref securityToken,
                                                                ref authenticationToken,
                                                                ref commerceIndicator,
                                                                ref bookingReference,
                                                                ref requestSource,
                                                                ref bWeb);


                    ds = RecordsetHelper.RecordsetToDataset(rs, "Payments");
                    if (ds != null && ds.Tables.Count > 0)
                    {
                        if (ds.Tables[0].Rows[0]["ResponseCode"].ToString() == "APPROVED")
                        {
                            bResult = true;
                        }
                        else
                        {
                            throw new BookingException("PaymentCreditCard fail : ");
                        }
                    }
                }
                else
                {
                    throw new BookingException("PaymentCreditCard fail : ");
                }
            }
            catch
            {
                throw;// new BookingException("PaymentCreditCard fail :  " + ex.Message);
            }
            finally
            {
                if (objBooking != null)
                {
                    Marshal.FinalReleaseComObject(objBooking);
                    objBooking = null;
                }
                if (objCreditCard != null)
                {
                    Marshal.FinalReleaseComObject(objCreditCard);
                    objCreditCard = null;
                }

                RecordsetHelper.ClearRecordset(ref rsHeader);
                RecordsetHelper.ClearRecordset(ref rsSegment);
                RecordsetHelper.ClearRecordset(ref rsPassenger);
                RecordsetHelper.ClearRecordset(ref rsRemark);
                RecordsetHelper.ClearRecordset(ref rsPayment);
                RecordsetHelper.ClearRecordset(ref rsTempPayment);
                RecordsetHelper.ClearRecordset(ref rsMapping);
                RecordsetHelper.ClearRecordset(ref rsService);
                RecordsetHelper.ClearRecordset(ref rsTax);
                RecordsetHelper.ClearRecordset(ref rsFees);
                RecordsetHelper.ClearRecordset(ref rsAllocation);
                RecordsetHelper.ClearRecordset(ref rsVoucher);
                ds = null;
            }

            return bResult;
        }

        private Entity.Voucher ValidateVoucher(Entity.Voucher voucher, IList<Entity.Booking.Mapping> Mappings, IList<Entity.Booking.Fee> Fees)
        {
            tikAeroProcess.Booking objBooking = null;
            tikAeroProcess.Payment objPayments = null;
            ADODB.Recordset rsVoucher = null;
            ADODB.Recordset rsMappings = null;
            ADODB.Recordset rsFees = null;
            ADODB.Recordset rsHeader = null;
            ADODB.Recordset rsSegment = null;
            ADODB.Recordset rsPassenger = null;
            ADODB.Recordset rsRemark = null;
            ADODB.Recordset rsPayment = null;
            ADODB.Recordset rsService = null;
            ADODB.Recordset rsTax = null;
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
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Booking", _server);
                    objBooking = (tikAeroProcess.Booking)Activator.CreateInstance(remote);
                    remote = null;
                    remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayments = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                {
                    objPayments = new tikAeroProcess.Payment();
                    objBooking = new tikAeroProcess.Booking(); 
                }

                // get empty
                string result = objBooking.GetEmpty(ref rsHeader,
                                   ref rsSegment,
                                   ref rsPassenger,
                                   ref rsRemark,
                                   ref rsPayment,
                                   ref rsMappings,
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

                if (result != string.Empty)
                {
                    if (Fees != null && Fees.Count > 0)
                    {
                        Fees.ToRecordset(ref rsFees);
                    }
                    if (Mappings != null && Mappings.Count > 0)
                    {
                        Mappings.ToRecordset(ref rsMappings);
                    }
                }

                // set rs voucher
                rsVoucher = RecordsetHelper.FabricateVoucherRecordset();
                voucher.ToRecordset(ref rsVoucher);

                // validate 
                 ADODB.Recordset rsVoucherAllocation = null;
                 rsVoucherAllocation = objPayments.ValidateVoucher(ref rsVoucher, ref rsMappings,ref rsFees, String.Empty);

                 //rsVoucherAllocation.MoveFirst();
                 //while (!rsVoucherAllocation.EOF)
                 //{
                 //    string currency_rcd = RecordsetHelper.ToString(rsVoucherAllocation, "currency_rcd");
                 //    double max_amount = RecordsetHelper.ToDouble(rsVoucherAllocation, "max_amount");
                 //    double dif_amount = RecordsetHelper.ToDouble(rsVoucherAllocation, "dif_amount");
                 //    double outstanding_amount = RecordsetHelper.ToDouble(rsVoucherAllocation, "outstanding_amount");
                 //}

                 if (rsVoucherAllocation == null)
                 {
                     voucher.VoucherStatusRcd = "NP";
                 }
                 else if (rsVoucherAllocation.RecordCount == 0)
                 {
                     voucher.VoucherStatusRcd = "NP";
                 }
                 else
                 {
                     rsVoucherAllocation.Filter = "dif_amount > 0";
                     if (rsVoucherAllocation.RecordCount > 0)
                     {
                         voucher.VoucherStatusRcd = "NP";
                     }
                 }

            }
            catch (BookingException ex)
            {
                throw ex;
            }
            finally
            {
                if (objPayments != null)
                {
                    Marshal.FinalReleaseComObject(objPayments);
                    objPayments = null;
                }

                RecordsetHelper.ClearRecordset(ref rsVoucher);
            }

            return voucher;
        }

        // for cob 
        public bool PaymentVoucher(decimal totalOutStanding,IList<Entity.Booking.Mapping> Mappings, IList<Entity.Booking.Fee> Fees,
                                   IList<Entity.Booking.Payment> payments,
                                   IList<Entity.PaymentAllocation> paymentAllocations, string agencyCode,string userId)
        {
            tikAeroProcess.Payment objPayment = null;
            tikAeroProcess.Miscellaneous objMiscellaneous = null;
            ADODB.Recordset rsPayment = null;
            bool result = false;
            ADODB.Recordset rsAllocation = null;
            ADODB.Recordset rsRefundVoucher = null;
            Entity.Voucher voucher = new Entity.Voucher();
            string error = string.Empty;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;

                    remote = Type.GetTypeFromProgID("tikAeroProcess.Miscellaneous", _server);
                    objMiscellaneous = (tikAeroProcess.Miscellaneous)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                {
                    objPayment = new tikAeroProcess.Payment();
                    objMiscellaneous = new tikAeroProcess.Miscellaneous();
                }

                // get empty rs object
                rsPayment = objPayment.GetEmpty();

                payments.ToRecordset(ref rsPayment);

                rsAllocation = RecordsetHelper.FabricatePaymentAllocationRecordset();

                VoucherService vcService = new VoucherService(_server, _user, _pass, _domain);
                bool IsValidVoucher = true;

                // filter only VC already
                if (payments != null && payments.Count > 0)
                {
                    foreach (Entity.Booking.Payment p in payments)
                    {
                        voucher.VoucherNumber = p.DocumentNumber;
                        voucher.VoucherPassword = p.DocumentPassword;
                        voucher.CurrencyRcd = p.CurrencyRcd;
                        voucher = vcService.GetVoucher(voucher, Mappings, Fees, agencyCode,userId);

                        // valid and check NP
                        if (voucher == null || voucher.VoucherStatusRcd == "NP")
                        {
                            IsValidVoucher = false;
                            error = "Invalid voucher or not match condition";
                        }
                        else
                        {
                            rsPayment.MoveFirst();
                            while (!rsPayment.EOF)
                            {
                                rsPayment.Fields["voucher_payment_id"].Value = Guid.Parse(voucher.VoucherId).ToRsString();
                                rsPayment.MoveNext();
                            }
                        }
                    }
                }

               
                if (IsValidVoucher)
                {
                    paymentAllocations.ToRecordset(ref rsAllocation);

                    if (voucher.ValidateVoucherEnough(totalOutStanding) == true)
                    {
                        result = objPayment.Save(ref rsPayment, ref rsAllocation, ref rsRefundVoucher);

                        if (result)
                        {
                            // createTickets = true;
                        }
                        else
                        {
                            throw new ModifyBookingException("PAYMENT FAIL", "P001");
                        }
                    }
                    else
                    {
                        throw new ModifyBookingException("VALUE NOT ENOUGH", "P002");
                    }
                }
                else
                {
                    throw new ModifyBookingException("PAYMENT DOCUMENT NOT FOUND", "P003");
                }
            }
            catch
            {
               // throw new ModifyBookingException(error, "P001");
                throw;
            }
            finally
            {
                // clear voucher session
                string strVoucherSessionId = Guid.NewGuid().ToString();
                if (voucher != null )
                {
                    objMiscellaneous.ReleaseVoucherSession(strVoucherSessionId, voucher.VoucherId);
                }

                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }

                RecordsetHelper.ClearRecordset(ref rsPayment);
            }

            return result;
        }
        //

        public bool PaymentAgencyAccount(decimal totalOutStanding, 
                                         string paymentType, 
                                         string agencyCode, 
                                         string userId,
                                         IList<Entity.Booking.Payment> payments,
                                         IList<Entity.PaymentAllocation> paymentAllocations)
        {

            tikAeroProcess.Payment objPayment = null;

            ADODB.Recordset rsPayment = null;
            bool result = false;
            ADODB.Recordset rsAllocation = null;
            ADODB.Recordset rsRefundVoucher = null;

            Agent agent = new Agent();
            AgencyService agencyService = new AgencyService(_server, _user, _pass, _domain);
            agent = agencyService.GetAgencySessionProfile(agencyCode, userId);

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                {
                    objPayment = new tikAeroProcess.Payment();
                }
                // get empty rs object
                rsPayment = objPayment.GetEmpty();

                payments.ToRecordset(ref rsPayment);

                rsAllocation = RecordsetHelper.FabricatePaymentAllocationRecordset();

                paymentAllocations.ToRecordset(ref rsAllocation);

                string agencyPaymentType = string.Empty;
                bool checkAgency = true;

                agencyPaymentType = agent.AgencyPaymentTypeRcd;

                if (agencyPaymentType.Equals(paymentType))
                {
                    if (agent.AgencyAccount - agent.BookingPayment > totalOutStanding)
                    {
                        foreach (Entity.Booking.Payment p in payments)
                        {
                            if (!p.ValidateCreditAgency(agent))
                            {
                                checkAgency = false;
                            }
                        }
                        if (checkAgency)
                        {
                            result = objPayment.Save(ref rsPayment, ref rsAllocation, ref rsRefundVoucher);

                            if (!result)
                            {
                                throw new ModifyBookingException("PAYMENT FAIL", "P001");
                            }
                        }
                    }
                    else
                    {
                        throw new ModifyBookingException("VALUE NOT ENOUGH", "P002");
                    }
                }
                else
                {
                    throw new ModifyBookingException("PAYMENT TYPE MISMATCH", "P004");
                }

            }
            catch
            {
                throw;// new ModifyBookingException("PAYMENT FAIL", "P001");
            }
            finally
            {
                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }

                RecordsetHelper.ClearRecordset(ref rsPayment);
            }

            return result;
        }

        public IList<Entity.Booking.Fee> GetPaymentTypeFee(
                                        string originRcd,
                                        string destinationRcd,
                                        string formOfPayment,
                                        string formOfPaymentSubtype,
                                        string currencyRcd,
                                        string agency,
                                        DateTime feeDate)
        {
            tikAeroProcess.Payment objPayment = null;
            ADODB.Recordset rsFees = null;
            Booking booking = new Booking();
            List<Entity.Booking.Fee> fees = new List<Entity.Booking.Fee>();

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Payment", _server);
                    objPayment = (tikAeroProcess.Payment)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                {
                    objPayment = new tikAeroProcess.Payment();
                }

                // get fee
                // CC  formOfPayment = CC, formOfPaymentSubtype = string.empty, currency,agency, DateTime.MinValue
                rsFees = objPayment.GetFormOfPaymentSubtypeFees(ref formOfPayment, ref formOfPaymentSubtype, ref currencyRcd, ref agency, ref feeDate);

                if (rsFees != null && rsFees.RecordCount > 0)
                {
                    booking.Fees = booking.Fees.FillBooking(rsFees);
                }
                //wail for revise code
                if (booking.Fees != null && booking.Fees.Count > 0)
                {
                    foreach (Entity.Booking.Fee f in booking.Fees)
                    {
                        if (f.FeeRcd == formOfPaymentSubtype && f.OriginRcd.ToUpper() == originRcd.ToUpper() && f.DestinationRcd.ToUpper() == destinationRcd.ToUpper())
                        {
                            fees.Add(f);
                            break;
                        }
                    }

                    // origin match
                    if (fees.Count == 0)
                    {
                        foreach (Entity.Booking.Fee f in booking.Fees)
                        {
                            if (f.FeeRcd == formOfPaymentSubtype && f.OriginRcd == originRcd && f.DestinationRcd == "")
                            {
                                fees.Add(f);
                                break;
                            }
                        }
                    }

                    // destination match
                    if (fees.Count == 0)
                    {
                        foreach (Entity.Booking.Fee f in booking.Fees)
                        {
                            if (f.FeeRcd == formOfPaymentSubtype && f.DestinationRcd == destinationRcd && f.OriginRcd == "")
                            {
                                fees.Add(f);
                                break;
                            }
                        }
                    }

                    // swap root
                    if (fees.Count == 0)
                    {
                        foreach (Entity.Booking.Fee f in booking.Fees)
                        {
                            if (f.FeeRcd == formOfPaymentSubtype && f.OriginRcd.ToUpper() == destinationRcd.ToUpper() && f.DestinationRcd.ToUpper() == originRcd.ToUpper())
                            {
                                fees.Add(f);
                                break;
                            }
                        }
                    }


                    // root == empty
                    if (fees.Count == 0)
                    {
                        foreach (Entity.Booking.Fee f in booking.Fees)
                        {
                            if (f.FeeRcd == formOfPaymentSubtype && string.IsNullOrEmpty(f.OriginRcd) && string.IsNullOrEmpty(f.DestinationRcd))
                            {
                                fees.Add(f);
                                break;
                            }
                        }
                    }


                    //wail for revise code
                    //if (fees.Count == 0)
                    //{
                    //    foreach (Entity.Booking.Fee f in booking.Fees)
                    //    {
                    //        if (f.FeeRcd.Trim().ToUpper() == formOfPaymentSubtype.Trim().ToUpper())
                    //        {
                    //            fees.Add(f);
                    //            break;
                    //        }
                    //    }
                    //}

                    // assign fee
                    booking.Fees = fees;
                }
            }
            catch
            {
                throw new ModifyBookingException("Get Payment Fee Fail", "P001");
            }
            finally
            {
                if (objPayment != null)
                {
                    Marshal.FinalReleaseComObject(objPayment);
                    objPayment = null;
                }

                RecordsetHelper.ClearRecordset(ref rsFees);
            }

            return booking.Fees;

        }
    }
}
