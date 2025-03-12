using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using Avantik.Web.Service.Model.Contract;
using Avantik.Web.Service.Entity.Route;
using Avantik.Web.Service.COMHelper;

namespace Avantik.Web.Service.Model.COM
{
    public class RouteService : RunComplus, IRouteService
    {
        string _server = string.Empty;
        public RouteService(string server, string user, string pass, string domain)
            : base(user, pass, domain)
        {
            _server = server;
        }

        public List<Route> GetOrigins(string strLanguage, bool b2cFlag, bool b2bFlag, bool b2EFlag, bool b2SFlag, bool APIFlag)
        {
            tikAeroProcess.Inventory objInventory = null;
            ADODB.Recordset rs = null;
            try
            {
                List<Route> routes = new List<Route>();
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Inventory", _server);
                    objInventory = (tikAeroProcess.Inventory)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objInventory = new tikAeroProcess.Inventory(); }

                rs = objInventory.GetOrigins(ref strLanguage, ref b2cFlag, ref b2bFlag, ref b2EFlag, ref b2SFlag, ref APIFlag);

                if (rs != null && rs.RecordCount > 0)
                {
                    routes.FillOrigin(ref rs);
                }

                return routes;
            }
            catch (SystemException ex)
            {
                throw ex;
            }
            finally
            {
                if (objInventory != null)
                {
                    Marshal.FinalReleaseComObject(objInventory);
                    objInventory = null;
                }
                RecordsetHelper.ClearRecordset(ref rs);
            }
        }

        public List<Route> GetDestinations(string strLanguage, bool b2cFlag, bool b2bFlag, bool b2EFlag, bool b2SFlag, bool APIFlag)
        {
            tikAeroProcess.Inventory objInventory = null;
            ADODB.Recordset rs = null;
            try
            {
                List<Route> routes = new List<Route>();
                if (string.IsNullOrEmpty(_server) == false)
                {
                    Type remote = Type.GetTypeFromProgID("tikAeroProcess.Inventory", _server);
                    objInventory = (tikAeroProcess.Inventory)Activator.CreateInstance(remote);
                    remote = null;
                }
                else
                { objInventory = new tikAeroProcess.Inventory(); }

                rs = objInventory.GetDestinations(ref strLanguage, ref b2cFlag, ref b2bFlag, ref b2EFlag, ref b2SFlag, ref APIFlag);

                if (rs != null && rs.RecordCount > 0)
                {
                    routes.FillDestination(ref rs);
                }

                return routes;
            }
            catch (SystemException ex)
            {
                throw ex;
            }
            finally
            {
                if (objInventory != null)
                {
                    Marshal.FinalReleaseComObject(objInventory);
                    objInventory = null;
                }

                RecordsetHelper.ClearRecordset(ref rs);
            }
        }
    }
}
