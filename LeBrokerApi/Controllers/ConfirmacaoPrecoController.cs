using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace WebApplication2.Controllers
{
    public class ConfirmacaoPrecoController : ApiController
    {
      public HttpResponseMessage Get()
      {
         var dataSet = new Models.DataSetFuncionario();
         dataSet.Funcionario.AddFuncionarioRow("André", "Lima", DateTime.Now);

         var report = new Microsoft.Reporting.WebForms.LocalReport();
         report.ReportPath = System.Web.Hosting.HostingEnvironment.MapPath("~/Reports/Contratos/relConfirmacaoPreco.rdlc");
         report.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource("DataSetFuncionario", (System.Data.DataTable)dataSet.Funcionario));
         report.Refresh();

         string mimeType = "";
         string encoding = "";
         string filenameExtension = "";
         string[] streams = null;
         Microsoft.Reporting.WebForms.Warning[] warnings = null;
         byte[] bytes = report.Render("PDF", null, out mimeType, out encoding, out filenameExtension, out streams, out warnings);

         HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
         result.Content = new ByteArrayContent(bytes);
         result.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue(mimeType);
         return result;
      }
   }
}
