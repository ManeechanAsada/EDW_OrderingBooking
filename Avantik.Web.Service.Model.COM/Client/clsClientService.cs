using Avantik.Web.Service.Model.Contract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avantik.Web.Service.COMHelper;
using System.Runtime.InteropServices;
using Avantik.Web.Service.Model.COM.Extension;
using Avantik.Web.Service.Entity;
using Avantik.Web.Service.Entity.Client;
using Avantik.Web.Service.Entity.Booking;

namespace Avantik.Web.Service.Model.COM
{
    public class ClientService : RunComplus, IClientService
    {
        string _server = string.Empty;
        public ClientService(string server, string user, string pass, string domain)
            :base(user,pass,domain)
        {
            _server = server;
        }
        public bool ClientSave(ClientProfile clientProfile)
        {
            tikAeroProcess.Client objClient = null;
            ADODB.Recordset rsClient = new ADODB.Recordset();
            ADODB.Recordset rsPassengerProfile = new ADODB.Recordset();
            ADODB.Recordset rsRemark = new ADODB.Recordset();
            ADODB.Recordset rsClientSegment = new ADODB.Recordset();
            bool result = false;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Client", _server);
                    objClient = (tikAeroProcess.Client)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objClient = new tikAeroProcess.Client(); }

                //get empty
                if (objClient.GetEmpty(ref rsClient, ref rsPassengerProfile, ref rsRemark))
                {
                    // map client to rs
                    clientProfile.Client.ToRecordsetClient(ref rsClient);
                    clientProfile.PassengerProfiles.ToRecordsetClient(ref rsPassengerProfile);
                    clientProfile.BookingRemarks.ToRecordset(ref rsRemark);
                }

                result = objClient.Save(ref rsClient, ref rsPassengerProfile, ref rsRemark);
            }
            catch
            {
            }
            finally
            {
                if (objClient != null)
                {
                    Marshal.FinalReleaseComObject(objClient);
                    objClient = null;
                }

                RecordsetHelper.ClearRecordset(ref rsClient);
                RecordsetHelper.ClearRecordset(ref rsPassengerProfile);
                RecordsetHelper.ClearRecordset(ref rsRemark);
            }

            return result;

        }

        public ClientProfile ClientRead(string clientProfileId)
        {
            tikAeroProcess.Client objClient = null;
            ADODB.Recordset rsClient = new ADODB.Recordset();
            ADODB.Recordset rsPassengerProfile = new ADODB.Recordset();
            ADODB.Recordset rsRemark = new ADODB.Recordset();
            ADODB.Recordset rsClientSegment = new ADODB.Recordset();
            ClientProfile ClientProfile = new Entity.Client.ClientProfile();
            clientProfileId = "{" + clientProfileId + "}";

            bool result = false;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Client", _server);
                    objClient = (tikAeroProcess.Client)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objClient = new tikAeroProcess.Client(); }

                result = objClient.Read(clientProfileId,ref rsClient, ref rsPassengerProfile, ref rsRemark);

                if (result == true)
                {
                    ClientProfile.Client = ClientProfile.Client.FillClientObject(rsClient);
                    ClientProfile.PassengerProfiles = ClientProfile.PassengerProfiles.FillClientObject(rsPassengerProfile);
                    ClientProfile.BookingRemarks = ClientProfile.BookingRemarks.FillBooking(rsRemark);
                }
            }
            catch
            {
            }
            finally
            {
                if (objClient != null)
                {
                    Marshal.FinalReleaseComObject(objClient);
                    objClient = null;
                }

                RecordsetHelper.ClearRecordset(ref rsClient);
                RecordsetHelper.ClearRecordset(ref rsPassengerProfile);
                RecordsetHelper.ClearRecordset(ref rsRemark);
            }

            return ClientProfile;

        }

        public bool EditClientProfile(ClientProfile clientProfile)
        {
            tikAeroProcess.Client objClient = null;
            ADODB.Recordset rsClient = new ADODB.Recordset();
            ADODB.Recordset rsPassengerProfile = new ADODB.Recordset();
            ADODB.Recordset rsRemark = new ADODB.Recordset();
            ADODB.Recordset rsClientSegment = new ADODB.Recordset();

            string clientProfileId = "{"+clientProfile.Client.ClientProfileId.ToString()+"}";

            bool result = false;
            bool resultRead = false;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Client", _server);
                    objClient = (tikAeroProcess.Client)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objClient = new tikAeroProcess.Client(); }

                // read
                resultRead = objClient.Read(clientProfileId, ref rsClient, ref rsPassengerProfile, ref rsRemark, ref rsClientSegment);
               
                // save
                if (resultRead)
                {
                    // map client to rs
                        clientProfile.Client.ToRecordsetClient(ref rsClient);

                    // find passenger role myself for edit
                        clientProfile.PassengerProfiles.ToRecordsetEditPassenger(ref rsPassengerProfile);

                    // remark
                        clientProfile.BookingRemarks.ToRecordset(ref rsRemark);
                }

                result = objClient.Save(ref rsClient, ref rsPassengerProfile, ref rsRemark);
            }
            catch
            {
            }
            finally
            {
                if (objClient != null)
                {
                    Marshal.FinalReleaseComObject(objClient);
                    objClient = null;
                }

                RecordsetHelper.ClearRecordset(ref rsClient);
                RecordsetHelper.ClearRecordset(ref rsPassengerProfile);
                RecordsetHelper.ClearRecordset(ref rsRemark);
            }

            return result;

        }

    }
}
