using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelXmlWriter
{

    struct Overpunch
    {
        public bool IsOverpunchable
        { get; set; }
        public double val
        { get; set; }
    }

    static class StaticFunctions
    {

        static readonly Dictionary<char, int> d = new Dictionary<char, int>() { 
        { '{', 0 }

        , { 'A', 1 } 
        , { 'a', 1 } 
        , { 'B', 2 } 
        , { 'b', 2 } 
        , { 'C', 3 } 
        , { 'c', 3 } 
        , { 'D', 4 } 
        , { 'd', 4 } 
        , { 'E', 5 } 
        , { 'e', 5 } 
        , { 'F', 6 } 
        , { 'f', 6 } 
        , { 'G', 7 } 
        , { 'g', 7 } 
        , { 'H', 8 } 
        , { 'h', 8 } 
        , { 'I', 9 } 
        , { 'i', 9 } 

        , { '}', 0 } 
        
        , { 'J', -1 }
        , { 'j', -1 }
        , { 'K', -2 }
        , { 'k', -2 }
        , { 'L', -3 }
        , { 'l', -3 }
        , { 'M', -4 }
        , { 'm', -4 }
        , { 'N', -5 }
        , { 'n', -5 }
        , { 'O', -6 }
        , { 'o', -6 }
        , { 'P', -7 }
        , { 'p', -7 }
        , { 'Q', -8 }
        , { 'q', -8 }
        , { 'R', -9 }
        , { 'r', -9 }
        };

        static readonly int excelMaxNumberLength = Convert.ToInt32(Resource1.ExcelMaximumNumberLength, CultureInfo.InvariantCulture);

        internal static Overpunch applyOverPunch(string stringToOverpunch, double significantDigits)
        {
            if (!string.IsNullOrEmpty(stringToOverpunch))
            {
                // remove anything invalid
                stringToOverpunch = Regex.Replace(stringToOverpunch, @"[^0-9a-rA-R]", string.Empty, RegexOptions.CultureInvariant);

                StringBuilder sb = new StringBuilder(stringToOverpunch);

                double test;
                // if conversion succeeds we're done
                if (double.TryParse(sb.ToString(), NumberStyles.Currency, CultureInfo.CurrentCulture, out test))
                    return new Overpunch() { val = test / significantDigits, IsOverpunchable = true };

                if (sb.Length > 0 && d.Any(x => x.Key == sb[sb.Length - 1]))
                {

                    int v = d.First(x => char.Equals(x.Key, sb[sb.Length - 1])).Value;
                    sb.Remove(sb.Length - 1, 1);
                    sb.Append(Math.Abs(v));
                    if (double.TryParse(sb.ToString(), NumberStyles.Currency, CultureInfo.CurrentCulture, out test))
                    {
                        if (v < 0)
                        {
                            test *= -1;
                        }

                        return new Overpunch() { val = test / significantDigits, IsOverpunchable = true };
                    }

                }
            }

            // can't do anyting
            
                return new Overpunch() { val = 0.0, IsOverpunchable = false };
        }

        internal static ExcelDataType ResolveDataType(string dataValue, double significantDigits)
        {
            decimal throwaway = new decimal();
            DateTime d = new DateTime();

            // Credit Card problem
            if (Decimal.TryParse(dataValue, out throwaway)
                && dataValue.Trim().Length <= excelMaxNumberLength
                && !dataValue.Trim().EndsWith("-")
                )
                return ExcelDataType.Number;
            else if (DateTime.TryParse(dataValue, out d)
                // excel doesn't like dates before 1900
                && d >= Convert.ToDateTime("1900-01-01")
                && d <= Convert.ToDateTime("9999-12-31"))
                return ExcelDataType.Date;
            else
            {
                //var o = applyOverPunch(dataValue, significantDigits);
                //if (o.IsOverpunchable)
                //    return ExcelDataType.OverpunchNumber;
                return ExcelDataType.String;
            }
        }
    }
}
