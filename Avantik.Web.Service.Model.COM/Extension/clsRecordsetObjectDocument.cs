using Avantik.Web.Service.COMHelper;
using Avantik.Web.Service.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Avantik.Web.Service.Model.COM.Extension
{
    public static class RecordsetObjectDocument
    {
        public static void FillDocumentType(this IList<Document> docs, ref ADODB.Recordset rs)
        {
            if (rs != null && rs.RecordCount > 0)
            {

                Document doc = null;

                try
                {
                    rs.MoveFirst();
                    while (!rs.EOF)
                    {
                        doc = new Document();
                        doc.DocumentTypeRcd = RecordsetHelper.ToString(rs, "document_type_rcd");
                        doc.DisplayName = RecordsetHelper.ToString(rs, "display_name");
                        docs.Add(doc);
                        rs.MoveNext();
                    }

                }
                catch
                {
                    throw;
                }
            }

        }
    }
}
