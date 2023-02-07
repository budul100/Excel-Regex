using ExcelDna.Integration;
using System;
using System.Text.RegularExpressions;

namespace ExcelRegex
{
    public static class Functions
    {
        #region Public Methods

        [ExcelFunction(Description = "Get regular expression match.\n" +
            "Refer to MS Regular Expression Language Reference for more details")]
        public static object RegexExtract([ExcelArgument("Input string")] string input,
            [ExcelArgument("Regex Pattern")] string pattern,
            [ExcelArgument("(Optional) group number to return. Default = 1.")]
            Object matchNo)
        {
            int groupNumber;

            if (matchNo is ExcelMissing)
            {
                groupNumber = 1; // default match
            }
            else
            {
                try
                {
                    groupNumber = Convert.ToInt32(matchNo);
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: '{0}'", e);
                    return ExcelError.ExcelErrorNum;
                }
            }

            var rgx = new Regex(pattern);
            var matches = rgx.Matches(input);

            // Return error if requested group number doesn't exist

            if (groupNumber > 1
                && matches[0].Groups.Count <= groupNumber)
            {
                return ExcelError.ExcelErrorNum;
            }

            // Return the first group if groups were found, otherwise return the match

            return matches[0].Groups.Count > 1
                ? matches[0].Groups[groupNumber].Value
                : (object)matches[0].Value;
        }

        [ExcelFunction(Description = "Test regular expression matches.\n" +
            "Refer to MS Regular Expression Language Reference for more details")]
        public static bool RegexMatches([ExcelArgument("Input string")] string input,
            [ExcelArgument("Regex Pattern")] string pattern)
        {
            var rgx = new Regex(pattern);
            var matches = rgx.Matches(input);

            return matches.Count > 0;
        }

        #endregion Public Methods
    }
}