using System;

namespace MMCSirUtilities
{
    public static class MMStringExtensions
    {
        /// <summary>
        /// Use this function when you need to fix an string to n digits
        /// </summary>
        /// <param name="data">The value to fix</param>
        /// <param name="symbol">The symbol that you want to use to fix the string</param>
        /// <param name="numberOfDigits">The quantity of digits that you need to add in the string</param>
        /// <returns></returns>
        public static string FixStrToNDigits(this string data, string symbol, int numberOfDigits = 1)
        {
            string retValue = string.Empty;
            if (symbol == null || symbol == string.Empty)
            {
                throw new FormatException("The Symbol can not be null or empty");
            }
            if (numberOfDigits <= 0)
            {
                throw new ArgumentException("The number of digits can not be equal to Zero or less.");
            }

            if (numberOfDigits < data.Length)
            {
                throw new FormatException("The data to fix is less than the number of digits.");
            }
            for (int i = 0; i < numberOfDigits; i++)
            {
                retValue += symbol;
            }
            retValue += data;
            retValue = retValue.Substring(retValue.Length - numberOfDigits, numberOfDigits);
            return retValue;
        }
    }
}