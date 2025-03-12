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
using System.Web;
using System.Data;
using System.Data.SqlClient;

namespace Avantik.Web.Service.Model.COM
{
    public class AgencyService : RunComplus, IAgencyService
    {
        string _server = string.Empty;
        public AgencyService(string server, string user, string pass, string domain)
            :base(user,pass,domain)
        {
            _server = server;
        }

        public Agent GetAgencySessionProfile(string strAgencyCode, string strUserId)
        {
            Agent agent = null; 
            ADODB.Recordset rsAgent = null;
            tikAeroProcess.System objSystem = null;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.System", _server);
                    objSystem = (tikAeroProcess.System)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objSystem = new tikAeroProcess.System(); }

                // implement agent cache
               // agent = (Agent)HttpRuntime.Cache[strAgencyCode];

                if (agent == null)
                {
                    rsAgent = objSystem.GetAgencySessionProfile(strAgencyCode, strUserId);
                    agent = agent.FillAgency(rsAgent);
                   // HttpRuntime.Cache.Insert(strAgencyCode, agent, null, System.Web.Caching.Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(20), System.Web.Caching.CacheItemPriority.Normal, null);
                }

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
                if (objSystem != null)
                { Marshal.FinalReleaseComObject(objSystem); }
                objSystem = null;

                RecordsetHelper.ClearRecordset(ref rsAgent);
            }

            return agent;
        }

        public Agent TravelAgentLogon(string agencyCode, string agentLogon, string agentPassword)
        {

            Agent agent = null;
            tikAeroProcess.System objSystem = null;
            ADODB.Recordset rsAgent = null;
            ADODB.Recordset rsUserLogon = null;
            ADODB.Recordset rsUser = null;
            string userId = string.Empty;

            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.System", _server);
                    objSystem = (tikAeroProcess.System)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objSystem = new tikAeroProcess.System(); }

                IList<User> userLogons = new List<User>();

                rsUserLogon = objSystem.TravelAgentLogon(agencyCode, agentLogon, agentPassword);
                userLogons = userLogons.FillUser(rsUserLogon);

                if (userLogons != null && userLogons.Count > 0)
                {
                    userId = userLogons[0].UserAccountId.ToRsString();
                    agent = GetAgencySessionProfile(agencyCode, userId);

                    // implement user cache
                   // agent.Users = (IList<User>)HttpRuntime.Cache[userId.ToUpper()];

                    if (agent.Users == null || agent.Users.Count == 0)
                    {
                        rsUser = UserRead(userId);
                        agent.Users = agent.Users.FillUser(rsUser);

                     //   HttpRuntime.Cache.Insert(userId.ToUpper(), agent.Users, null, System.Web.Caching.Cache.NoAbsoluteExpiration, TimeSpan.FromMinutes(20), System.Web.Caching.CacheItemPriority.Normal, null);
                    }
                }
                else
                {
                    throw new AgentLogonException("Invalid userlogon or password");
                }
            }
            catch (System.Runtime.InteropServices.COMException exc)
            {
                throw exc;
            }
            catch
            {
                throw ;
            }
            finally
            {
                if (objSystem != null)
                { Marshal.FinalReleaseComObject(objSystem); }
                objSystem = null;

                RecordsetHelper.ClearRecordset(ref rsUser);
                RecordsetHelper.ClearRecordset(ref rsUserLogon);
                RecordsetHelper.ClearRecordset(ref rsAgent);
            }

            return agent;
        }
       
        public ADODB.Recordset UserRead(string userId)
        {
            tikAeroProcess.Settings objSetting = null;

            ADODB.Recordset rs = null;
            try
            {
                if (_server.Length > 0)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Settings", _server);
                    objSetting = (tikAeroProcess.Settings)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objSetting = new tikAeroProcess.Settings(); }

                rs = objSetting.UserRead(userId);
            }
            catch
            { }
            finally
            {
                if (objSetting != null)
                {
                    Marshal.FinalReleaseComObject(objSetting);
                    objSetting = null;
                }
            }


            return rs;
        }

    }
}
