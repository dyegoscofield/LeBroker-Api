using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Data;
using System.Net.Http;
using System.Web.Http;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using Classes.Uteis;

namespace WebApplication2.Controllers
{
   public class TesteReplaceController : ApiController
   {
      public HttpResponseMessage Get()
      {
         HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);

         DataTable dt = new DataTable();
         dt.Columns.AddRange(new DataColumn[29] {
            new DataColumn("PRODUTO"),
            new DataColumn("NUMERO_CONTRATO"),
            new DataColumn("CIDADE_DOC"),
            new DataColumn("DATA"),
            new DataColumn("COMPRADOR"),
            new DataColumn("ENDERECO_COMPRADOR"),
            new DataColumn("CNPJ_COMPRADOR"),
            new DataColumn("INSCRICAO_COMPRADOR"),
            new DataColumn("VENDEDOR"),
            new DataColumn("ENDERECO_VENDEDOR"),
            new DataColumn("CNPJ_VENDEDOR"),
            new DataColumn("INSCRICAO_VENDEDOR"),
            new DataColumn("SAFRA"),
            new DataColumn("QUANTIDADE"),
            new DataColumn("QUANTIDADE_EXTENSO"),
            new DataColumn("PRAZO_ENTREGA"),
            new DataColumn("LOCAL_ENTREGA"),
            new DataColumn("TAMANHO_SACA"),
            new DataColumn("VLR_SACA"),
            new DataColumn("VLR_EXTENSO"),
            new DataColumn("VLR_BRUTO"),
            new DataColumn("BRUTO_EXTENSO"),
            new DataColumn("DT_CONTRATO"),
            new DataColumn("BANCO"),
            new DataColumn("AGENCIA"),
            new DataColumn("CONTA"),
            new DataColumn("TITULARIDADE"),
            new DataColumn("CPF"),
            new DataColumn("DT_EXTENSO"),
         });
         dt.Rows.Add("SOJA",
            "304AG",
            "UBERLANDIA/MG",
            DateTime.Now.ToShortDateString(),
            "AMAGGI EXPORTAÇÃO E IMPORTAÇÃO LTDA",
            "AVENIDA PRESIDENTE VARGAS, QUADRA R, SALA 604. 6º ANDAR, COND. EMP. LE MONDE, BAIRRO MARCONAL, RIO VERDE – GO, CEP 75.901-551",
            "77.294.254/0074-40",
            "10.593.088-1",
            "ISIS IPOLITO DE PAULA RIBEIRO",
            "ROD. GO 118 KM 153 A DIREITA 20 KM – FAZENDA SÃO JORGE, SÃO JOÃO D ALIANÇA - GO, CEP 73.760-000",
            "001.966.561-09",
            "11.402.866-4",
            "2019/2019",
            "1.200.000",
            NumeroExtenso.NumeroPorExtenso(Decimal.Parse("1.200.000")).ToUpper(),
            "01/09/2019 a 15/10/2019",
            "FAZENDA SÃO JORGE, no município de SÃO JOÃO D´ALIANÇA – GO",
            "60KG",
            "22,80",
            ValorExtenso.NumeroPorExtenso(Decimal.Parse("22,80")).ToUpper(),
            "456.000,00",
            ValorExtenso.NumeroPorExtenso(Decimal.Parse("456.000,00")).ToUpper(),
            "04/11/2019",
            "BANCO DO BRASIL (001)",
            "0377-8",
            "47.233-6",
            "ISIS IPOLIO DE PAULA RIBEIRO",
            "001.966.561-09",
            DataExtenso.RetornaDataExtenso(DateTime.Now)
            );
         if (dt.Rows.Count > 0)
         {
            object fileName = System.Web.Hosting.HostingEnvironment.MapPath("~/Files/CONFIRMACAO_FIXACAO_PRECO.docx");
            Word.Application word = new Word.Application();
            Word.Document doc = new Word.Document();
            object missing = System.Type.Missing;
            
            try
            {
               doc = word.Documents.Add(fileName, missing, missing, missing);
               doc.Activate();

               string path = System.Web.Hosting.HostingEnvironment.MapPath("~/Files");
               Object oSaveAsFile = (Object)(path + @"\CONFIRMACAO_FIXACAO_PRECO - "+ dt.Rows[0]["NUMERO_CONTRATO"].ToString() + " - " + dt.Rows[0]["COMPRADOR"].ToString() + ".docx");
               Object oMissing = System.Reflection.Missing.Value;

               FindAndReplace(word, "#PRODUTO", dt.Rows[0]["PRODUTO"].ToString());
               FindAndReplace(word, "#NUMERO_CONTRATO", dt.Rows[0]["NUMERO_CONTRATO"].ToString());
               FindAndReplace(word, "#CIDADE_DOC", dt.Rows[0]["CIDADE_DOC"].ToString());
               FindAndReplace(word, "#DATA", dt.Rows[0]["DATA"].ToString());
               FindAndReplace(word, "#COMPRADOR", dt.Rows[0]["COMPRADOR"].ToString());
               FindAndReplace(word, "#ENDERECO_COMPRADOR", dt.Rows[0]["ENDERECO_COMPRADOR"].ToString());
               FindAndReplace(word, "#CNPJ_COMPRADOR", dt.Rows[0]["CNPJ_COMPRADOR"].ToString());
               FindAndReplace(word, "#INSCRICAO_COMPRADOR", dt.Rows[0]["INSCRICAO_COMPRADOR"].ToString());
               FindAndReplace(word, "#VENDEDOR", dt.Rows[0]["VENDEDOR"].ToString());
               FindAndReplace(word, "#ENDERECO_VENDEDOR", dt.Rows[0]["ENDERECO_VENDEDOR"].ToString());
               FindAndReplace(word, "#CNPJ_VENDEDOR", dt.Rows[0]["CNPJ_VENDEDOR"].ToString());
               FindAndReplace(word, "#INSCRICAO_VENDEDOR", dt.Rows[0]["INSCRICAO_VENDEDOR"].ToString());
               FindAndReplace(word, "#SAFRA", dt.Rows[0]["SAFRA"].ToString());
               FindAndReplace(word, "#QUANTIDADE", dt.Rows[0]["QUANTIDADE"].ToString());
               FindAndReplace(word, "#QTD_EXTENSO", dt.Rows[0]["QUANTIDADE_EXTENSO"].ToString());

               if(dt.Rows[0]["PRODUTO"].ToString() == "MILHO")
                  FindAndReplace(word, "#PADRAO_ANEC", "com até 14% de umidade, 5% de impurezas totais, sendo 1% de ardido (PADRÃO EXPORTAÇÃO ANEC 44)");
               else
                  FindAndReplace(word, "#PADRAO_ANEC", "com até 14% de umidade, 1% de impurezas, 8% de avariados, estes últimos com até 6% de mofados, 4% de ardidos e 1% queimados, 8% de grãos esverdeados, 30% de grãos quebrados. A Classificação será efetuada a cada caminhão e/ou vagão recebido");

               FindAndReplace(word, "#PRAZO_ENTREGA", dt.Rows[0]["PRAZO_ENTREGA"].ToString());
               FindAndReplace(word, "#LOCAL_ENTREGA", dt.Rows[0]["LOCAL_ENTREGA"].ToString());
               FindAndReplace(word, "#TAMANHO_SACA", dt.Rows[0]["TAMANHO_SACA"].ToString());
               FindAndReplace(word, "#VLR_SACA", dt.Rows[0]["VLR_SACA"].ToString());
               FindAndReplace(word, "#VLR_EXTENSO", dt.Rows[0]["VLR_EXTENSO"].ToString());
               FindAndReplace(word, "#VLR_BRUTO", dt.Rows[0]["VLR_BRUTO"].ToString());
               FindAndReplace(word, "#BRUTO_EXTENSO", dt.Rows[0]["BRUTO_EXTENSO"].ToString());
               FindAndReplace(word, "#DT_CONTRATO", dt.Rows[0]["DT_CONTRATO"].ToString());
               FindAndReplace(word, "#BANCO", dt.Rows[0]["BANCO"].ToString());
               FindAndReplace(word, "#AGENCIA", dt.Rows[0]["AGENCIA"].ToString());
               FindAndReplace(word, "#CONTA", dt.Rows[0]["CONTA"].ToString());
               FindAndReplace(word, "#TITULARIDADE", dt.Rows[0]["TITULARIDADE"].ToString());
               FindAndReplace(word, "#CPF", dt.Rows[0]["CPF"].ToString());
               FindAndReplace(word, "#DT_EXTENSO", dt.Rows[0]["DT_EXTENSO"].ToString().ToLower());

               doc.SaveAs2(
                ref oSaveAsFile, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);
               doc.Close(ref missing, ref missing, ref missing);
               ((Word._Application)word).Quit();

               GerarArquivoPdf(path + @"\CONFIRMACAO_FIXACAO_PRECO - " + dt.Rows[0]["NUMERO_CONTRATO"].ToString() + " - " + dt.Rows[0]["COMPRADOR"].ToString() + ".docx", path + @"\CONFIRMACAO_FIXACAO_PRECO - " + dt.Rows[0]["NUMERO_CONTRATO"].ToString() + " - " + dt.Rows[0]["COMPRADOR"].ToString() + ".pdf");
            }
            catch (Exception ex)
            {
               result.StatusCode = HttpStatusCode.InternalServerError;
            }
         }

         return result;
      }

      private static void GerarArquivoPdf(string caminhoDoc, string caminhoPDF)
      {
         try
         {
            // Abrir Aplicacao Word
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            // Arquivo de Origem
            object filename = caminhoDoc;

            object newFileName = caminhoPDF;

            object missing = System.Type.Missing;

            // Abrir documento
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref filename, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);

            // Formato para Salvar o Arquivo – Destino  - No caso, PDF
            object formatoArquivo = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;

            // Salvar Arquivo
            doc.SaveAs(ref newFileName, ref formatoArquivo, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            // Não salvar alterações no arquivo original
            object salvarAlteracoesArqOriginal = false;
            wordApp.Quit(ref salvarAlteracoesArqOriginal, ref missing, ref missing);

         }
         catch (Exception ex)
         {
            
         }
      }

      private void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
      {
         //options
         object matchCase = false;
         object matchWholeWord = true;
         object matchWildCards = false;
         object matchSoundsLike = false;
         object matchAllWordForms = false;
         object forward = true;
         object format = false;
         object matchKashida = false;
         object matchDiacritics = false;
         object matchAlefHamza = false;
         object matchControl = false;
         object read_only = false;
         object visible = true;
         object replace = 2;
         object wrap = 1;
         //execute find and replace
         doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
             ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
             ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
      }

   }
}
