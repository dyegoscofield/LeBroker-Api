using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace Classes.Uteis
{
   public static class ValorExtenso
   {
      private static ArrayList numeroLista;

      private static Int32 num;

      //array de 2 linhas e 14 colunas[2][14]
      private static readonly String[,] qualificadores = new String[,] {
//			{"milésimo de real","milésimos de real"},//[0][0] e [0][1]
			{"centavo", "centavos"},//[1][0] e [1][1]
			{"", ""},//[2][0],[2][1]
			{"Mil", "Mil"},
         {"Milhão", "Milhões"},
         {"Bilhão", "Bilhões"},
         {"Trilhão", "Trilhões"},
         {"Quatrilhão", "Quatrilhões"},
         {"Quintilhão", "Quintilhões"},
         {"Sextilhão", "Sextilhões"},
         {"Setilhão", "Setilhões"},
         {"Octilhão","Octilhões"},
         {"Nonilhão","Nonilhões"},
         {"Decilhão","Decilhões"}
         };

      private static readonly String[,] numeros = new String[,] {
         {"Zero", "Um", "Dois", "Três", "Quatro",
          "Cinco", "Seis", "Sete", "Oito", "Nove",
          "Dez","Onze", "Doze", "Treze", "Quatorze",
          "Quinze", "Dezesseis", "Dezessete", "Dezoito", "Dezenove"},
         {"Vinte", "Trinta", "Quarenta", "Cinquenta", "Sessenta",
          "Setenta", "Oitenta", "Noventa",
          null,null,null,null,null,null,null,
          null,null,null,null,null},
         {"Cem", "Cento",
          "Duzentos", "Trezentos", "Quatrocentos",
          "Quinhentos", "Seiscentos",
          "Setecentos", "Oitocentos", "Novecentos",
          null,null,null,null,null,null,null,null,
          null,null}
         };

      public static void NumeroPorExtenso()
      {
         numeroLista = new ArrayList();
      }

      public static string NumeroPorExtenso(Decimal dec)
      {
         numeroLista = new ArrayList();
         SetNumero(dec);
         return ConverteNrExtenso();

      }

      public static void SetNumero(Decimal dec)
      {
         dec = Decimal.Round(dec, 2);
         dec = dec * 100;
         num = Convert.ToInt32(dec);


         //numeroLista.Clear();
         if (num == 0)
         {
            numeroLista.Add(0);
            numeroLista.Add(0);
         }
         else
         {
            AddRemainder(100);

            while (num != 0)
            {
               AddRemainder(1000);
            }
         }
      }

      private static void AddRemainder(Int32 divisor)
      {
         Int32 div = num / divisor;
         Int32 mod = num % divisor;

         Int32[] newNum = new Int32[] { div, mod };

         numeroLista.Add(mod);

         num = div;
      }

      private static bool TemMaisGrupos(Int32 ps)
      {
         while (ps > 0)
         {
            if ((Int32)numeroLista[ps] != 00 && !TemMaisGrupos(ps - 1))
               return true;
            ps--;
         }
         return true;
      }

      private static bool EhPrimeiroGrupoUm()
      {
         if ((Int32)numeroLista[numeroLista.Count - 1] == 1)
            return true;
         else
            return false;
      }

      private static bool EhUltimoGrupo(Int32 ps)
      {
         return ((ps > 0) && ((Int32)numeroLista[ps] != 0) || !TemMaisGrupos(ps - 1));
      }

      private static bool EhGrupoZero(Int32 ps)
      {
         if (ps <= 0 || ps >= numeroLista.Count)
            return true;
         return ((Int32)numeroLista[ps] == 0);
      }

      private static bool EhUnicoGrupo()
      {
         if (numeroLista.Count <= 3)
            return false;

         if (!EhGrupoZero(1) && !EhGrupoZero(2))
            return false;

         bool hasOne = false;

         for (Int32 i = 3; i < numeroLista.Count; i++)
         {
            if ((Int32)numeroLista[i] != 0)
            {
               if (hasOne)
                  return false;
               hasOne = true;
            }
         }
         return true;
      }

      private static String NumToString(Int32 numero, Int32 escala)
      {
         Int32 unidade = (numero % 10);
         Int32 dezena = (numero % 100);
         Int32 centena = (numero / 100);

         StringBuilder buf = new StringBuilder();

         if (numero != 0)
         {
            if (centena != 0)
            {
               if (dezena == 0 && centena == 1)
               {
                  buf.Append(numeros[2, 0]);
               }
               else
               {
                  buf.Append(numeros[2, centena]);
               }
            }

            if (buf.Length > 0 && dezena != 0)
            {
               buf.Append(" e ");
            }

            if (dezena > 19)
            {
               dezena = dezena / 10;
               buf.Append(numeros[1, dezena - 2]);
               if (unidade != 0)
               {
                  buf.Append(" e ");
                  buf.Append(numeros[0, unidade]);
               }
            }
            else if (centena == 0 || dezena != 0)
            {
               buf.Append(numeros[0, dezena]);
            }

            buf.Append(" ");

            if (numero == 1)
            {
               buf.Append(qualificadores[escala, 0]);
            }
            else
            {
               buf.Append(qualificadores[escala, 1]);
            }

         }
         return buf.ToString();
      }

      public static String ConverteNrExtenso()
      {
         StringBuilder buf = new StringBuilder();

         Int32 numero = (Int32)numeroLista[0];
         Int32 count;
         for (count = numeroLista.Count - 1; count > 0; count--)
         {
            if (buf.Length > 0 && !EhGrupoZero(count))
            {
               buf.Append(" e ");
            }
            buf.Append(NumToString((Int32)numeroLista[count], count));
         }

         if (buf.Length > 0)
         {

            while (buf.ToString().EndsWith(" "))
               buf.Length = buf.Length - 1;

            if (EhUnicoGrupo())
            {
               buf.Append(" de");
            }
            else
            {
               buf.Append(" ");
            }

            if (EhPrimeiroGrupoUm())
            {
               buf.Insert(0, "");
            }

            if (numeroLista.Count == 2 && ((Int32)numeroLista[1] == 1))
            {
               buf.Append(" Real");
            }
            else
            {
               buf.Append(" Reais");
            }

            if ((Int32)numeroLista[0] != 0)
            {
               buf.Append(" e ");
            }
         }

         if ((Int32)numeroLista[0] != 0)
         {
            buf.Append(NumToString((Int32)numeroLista[0], 0));
         }

         return buf.ToString();
      }
   }
}

