using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace Classes.Uteis
{
   public static class DataExtenso
   {
      public static string RetornaDataExtenso(DateTime data)
      {
         CultureInfo culture = new CultureInfo("pt-BR");
         DateTimeFormatInfo dtfi = culture.DateTimeFormat;

         int dia = data.Day;
         int ano = data.Year;
         string mes = culture.TextInfo.ToTitleCase(dtfi.GetMonthName(data.Month));
         string dataExtenso = dia + " de " + mes + " de " + ano;

         return dataExtenso;
      }
   }
}
