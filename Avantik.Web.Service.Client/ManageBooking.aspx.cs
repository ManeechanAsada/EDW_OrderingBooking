using Avantik.Web.Service.Message;
using Avantik.Web.Service.Message.Booking;
using Avantik.Web.Service.Proxy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Runtime.Serialization.Json;
using Avantik.Web.Service.Entity.Route;
using Avantik.Web.Service.Message.System;
using Avantik.Web.Service.Message.ManageBooking;

namespace Avantik.Web.Service.Client
{
    public partial class ManageBooking : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //AuthenticationServiceProxy objAuthen = new AuthenticationServiceProxy();
            //ManageBookingServiceProxy objBooking = new ManageBookingServiceProxy();
            //InitializeRequest req = new InitializeRequest();
            //InitializeResponse res = new InitializeResponse();

            //req.AgencyCode = "DATALEX";
            //req.AgentLogon = "DEMO";
            //req.AgentPassword = "1Password";

            //res = objAuthen.ServiceInitialize(req);

            //if (res.Success == true)
            //{

            //    SegmentCancelRequest reqCancel = new SegmentCancelRequest();
            //    SegmentCancelResponse responseCancel = new SegmentCancelResponse();

            //    reqCancel.Token = res.Token;
            //    reqCancel.BookingId = "7fe49269-e06f-43a7-bd2e-e1b1f57291f9";
            //    reqCancel.SegmentId = "d1a72acf-562d-4361-8b53-3ec84ef0ec37";
            //    reqCancel.UserId = "";
            //    responseCancel = objBooking.CancelSegment(reqCancel);

            //    bool test = responseCancel.Success;


            //}

            //ModifyBooking();
        }

        private void ModifyBooking()
        {
            AuthenticationServiceProxy objAuthen = new AuthenticationServiceProxy();
            ManageBookingServiceProxy objBooking = new ManageBookingServiceProxy();
            InitializeRequest req = new InitializeRequest();
            InitializeResponse res = new InitializeResponse();

            req.AgencyCode = "DATALEX";
            req.AgentLogon = "DEMO";
            req.AgentPassword = "1Password";

            res = objAuthen.ServiceInitialize(req);

            if (res.Success == true)
            {

                //ModifyBookingRequest reqModify = new ModifyBookingRequest();
                //ModifyBookingResponse responseModify = new ModifyBookingResponse();

                //string BookingId = "4fbb4219-2ce3-4372-9fa0-297dc5d36805";

                //Message.ManageBooking.Flight flt = new Message.ManageBooking.Flight();
                //flt.BoardingClassRcd = "Y";
                //flt.BookingClassRcd = "Z";
                //flt.DepartureDate = new DateTime(2015, 01, 14);
                //flt.DestinationRcd = "KMS";
                //flt.FairId = new Guid("D1C2296A-B1DE-4B72-AF37-C024BFC19A5A");
                //flt.FlightConnectionId = new Guid("00000000-0000-0000-0000-000000000000");
                //flt.FlightId = new Guid("9DB4DD6F-D7B0-44D4-B70E-795A29FDD914");
                //flt.OdDestinationRcd = "KMS";
                //flt.OdOriginRcd = "ACC";
                //flt.OriginRcd = "ACC";

                //IList<Message.ManageBooking.Flight> flts = new List<Message.ManageBooking.Flight>();
                //flts.Add(flt);

                //Message.ManageBooking.Payment pay = new Message.ManageBooking.Payment();

                ////pay.BookingPaymentId = new Guid("f38e4248-5380-4c50-bc0e-0c5621044419");
                ////pay.BookingSegmentId = System.Guid.NewGuid();
                ////pay.BookingId = new Guid(BookingId);
                ////pay.VoucherPaymentId = System.Guid.NewGuid();
                //pay.FormOfPaymentRcd = "VOUCHER";
                //pay.CurrencyRcd = "EUR";
                //pay.ReceiveCurrencyRcd = "EUR";
                //pay.AgencyCode = "B2CTPC";
                //pay.DebitAgencyCode = "B2CTPC";
                ////pay.PaymentAmount = decimal.Parse("32.4");
                ////pay.ReceivePaymentAmount = decimal.Parse("32.4");
                ////pay.AcctPaymentAmount = 162;
                ////pay.PaymentBy = System.Guid.NewGuid();
                ////pay.PaymentDateTime = new DateTime(2015, 01, 07);
                //pay.PaymentDueDateTime = new DateTime(1900, 12, 12);
                ////pay.DocumentAmount = decimal.Parse("32.4");
                ////pay.VoidBy = System.Guid.NewGuid();
                //pay.ExpiryMonth = 0;
                //pay.ExpiryYear = 0;
                ////pay.VoidDateTime = new DateTime(1900, 12, 12);
                //pay.RecordLocator = "record_locator1";
                //pay.NameOnCard = "name_on_card1";
                //pay.DocumentNumber = "3700002465";
                //pay.FormOfPaymentSubtypeRcd = "VOUCHER";
                //pay.City = "city1";
                //pay.State = "state1";
                //pay.Street = "street1";
                //pay.AddressLine1 = "address_line11";
                //pay.ZipCode = "zip_code1";
                //pay.CountryRcd = "EUR";
                ////pay.CreateBy = System.Guid.NewGuid();
                ////pay.CreateDateTime = new DateTime(1900, 12, 12);
                ////pay.UpdateBy = System.Guid.NewGuid();
                ////pay.UpdateDateTime = new DateTime(1900, 12, 12);
                //pay.PosIndicator = "pos_indicator1";
                //pay.IssueMonth = 0;
                //pay.IssueYear = 0;
                //pay.IssueNumber = "issue_number1";
                ////pay.ClientProfileId = System.Guid.NewGuid();
                //pay.TransactionReference = "transaction_reference1";
                //pay.ApprovalCode = "0";
                //pay.BankName = "bank_name1";
                //pay.BankCode = "bank_code1";
                //pay.BankIban = "bank_iban1";
                //pay.IpAddress = "ip_address1";
                //pay.PaymentReference = "0";
                ////pay.AllocatedAmount = decimal.Parse("32.4");
                ////pay.PaymentTypeRcd = "PAYMENT";
                //pay.RefundedAmount = 0;

                //IList<Message.ManageBooking.Payment> pp = new List<Message.ManageBooking.Payment>();
                //pp.Add(pay);

                //reqModify.BookingId = BookingId;
                //reqModify.ActionCode = "PRI";
                //reqModify.Payments = pp;
                //reqModify.Flight = flts;
                //reqModify.Token = res.Token;
                ////reqModify.UserId = "";
                //responseModify = objBooking.ModifyBooking(reqModify);

            }
        }
    }
}