using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Math;

namespace ExcelFunctions
{
    class Math
    {
        /// <summary>
        /// Returns number raised to the power
        /// </summary>
        /// <param name="number"></param>
        /// <param name="power"></param>
        /// <returns></returns>
        public double ePower(double number, double power)
        {
            return Pow(number, power);
        }

        /// <summary>
        /// Returns maximum value
        /// </summary>
        /// <param name="numbers"></param>
        /// <returns></returns>
        public double eMax(double[] numbers)
        {
            return numbers.Max();
        }

        /// <summary>
        /// Returns minimal value
        /// </summary>
        /// <param name="numbers"></param>
        /// <returns></returns>
        public double eMin(double[] numbers)
        {
            return numbers.Min();
        }

        /// <summary>
        /// Returns minimal value that meets a certain criterion
        /// </summary>
        /// <param name="numbers"></param>
        /// <param name="conditionSign">"use one of: &lt;,&lt;=,&gt;,&gt;=,&lt;&gt;,="</param>
        /// <param name="conditionValue"></param>
        /// <returns></returns>
        public double eMinIf(double[] numbers, string conditionSign, double conditionValue)
        {
            if (conditionSign == "=")
            {
                return numbers.Where(number => number == conditionValue).Min();
            }
            else if (conditionSign == ">")
            {
                return numbers.Where(number => number > conditionValue).Min();
            }
            else if (conditionSign == ">=")
            {
                return numbers.Where(number => number >= conditionValue).Min();
            }
            else if (conditionSign == "<")
            {
                return numbers.Where(number => number < conditionValue).Min();
            }
            else if (conditionSign == "<=")
            {
                return numbers.Where(number => number <= conditionValue).Min();
            }
            else if (conditionSign == "<>")
            {
                return numbers.Where(number => number != conditionValue).Min();
            }
            else
            {
                throw new ArgumentException("wrong conditionSign, use one of: >, >=, <, <=, =, <> ");
            }
        }
        /// <summary>
        /// Returns maximal value that meets a certain criterion
        /// </summary>
        /// <param name="numbers"></param>
        /// <param name="conditionSign">"use one of: &lt;,&lt;=,&gt;,&gt;=,&lt;&gt;,="</param>
        /// <param name="conditionValue"></param>
        /// <returns></returns>
        public double eMaxIf(double[] numbers, string conditionSign, double conditionValue)
        {
            if (conditionSign == "=")
            {
                return numbers.Where(number => number == conditionValue).Max();
            }
            else if (conditionSign == ">")
            {
                return numbers.Where(number => number > conditionValue).Max();
            }
            else if (conditionSign == ">=")
            {
                return numbers.Where(number => number >= conditionValue).Max();
            }
            else if (conditionSign == "<")
            {
                return numbers.Where(number => number < conditionValue).Max();
            }
            else if (conditionSign == "<=")
            {
                return numbers.Where(number => number <= conditionValue).Max();
            }
            else if (conditionSign == "<>")
            {
                return numbers.Where(number => number != conditionValue).Max();
            }
            else
            {
                throw new ArgumentException("wrong conditionSign, use one of: >, >=, <, <=, =, <> ");
            }
        }

        /// <summary>
        /// Returns absolute value
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public double eAbs(double number)
        {
            return (number > 0 ? number : -number);
        }

        /// <summary>
        /// Multiplies the corresponding items in the arrays and returns the sum of the results. If one of arrays is shorter, function repeats its values.
        /// </summary>
        /// <param name="tab1"></param>
        /// <param name="tab2"></param>
        /// <returns></returns>
        public double eSumProduct(double[] tab1, double[] tab2)
        {
            int len = Max(tab1.Length, tab2.Length);
            double[] result = new double[len];
            int j = -1;

            if (tab1.Length <= tab2.Length)
            {
                for (int i = 0; i < len; i++)
                {
                    j = j + 1;
                    result[i] = tab1[j] * tab2[i];

                    if (j == tab1.Length - 1)
                        j = -1;
                }
            }
            else
            {
                for (int i = 0; i < len; i++)
                {
                    j = j + 1;
                    result[i] = tab2[j] * tab1[i];

                    if (j == tab2.Length - 1)
                        j = -1;
                }
            }

            return result.Sum();
        }


        /// <summary>
        /// Returns sum of values that meet a certain criterion
        /// </summary>
        /// <param name="numbers"></param>
        /// <param name="conditionValue"></param>
        /// <param name="conditionSign">"use one of: &lt;,&lt;=,&gt;,&gt;=,&lt;&gt;,="</param>
        /// <returns></returns>
        public double eSumIf(double[] numbers, string conditionSign, double conditionValue)
        {
            if (conditionSign == "=")
            {
                return numbers.Sum(element => (element == conditionValue ? element : 0));
            }
            else if (conditionSign == ">")
            {
                return numbers.Sum(element => (element > conditionValue ? element : 0));
            }
            else if (conditionSign == ">=")
            {
                return numbers.Sum(element => (element >= conditionValue ? element : 0));
            }
            else if (conditionSign == "<")
            {
                return numbers.Sum(element => (element < conditionValue ? element : 0));
            }
            else if (conditionSign == "<=")
            {
                return numbers.Sum(element => (element <= conditionValue ? element : 0));
            }
            else if (conditionSign == "<>")
            {
                return numbers.Sum(element => (element != conditionValue ? element : 0));
            }
            else
            {
                throw new ArgumentException("wrong conditionSign, use one of: >, >=, <, <=, =, <> ");
            }
        }

        /// <summary>
        /// Returns sum of values that meet a certain criterion (beetween value1 and value2)
        /// </summary>
        /// <param name="numbers"></param>
        /// <param name="conditionValue1"></param>
        /// <param name="conditionValue2"></param>
        /// <param name="between">"closed or open interval"</param>
        /// <returns></returns>
        public double eSumIf(double[] numbers, double conditionValue1, double conditionValue2, string between = "closed")
        {
            if (between == "open")
            {
                return numbers.Sum(element => ((element > conditionValue1 && element < conditionValue2) ? element : 0));
            }
            else if (between == "closed")
            {
                return numbers.Sum(element => ((element >= conditionValue1 && element <= conditionValue2) ? element : 0));
            }
            else
            {
                throw new ArgumentException("wrong conditionSign, use one of: >, >=, <, <=, =, <> ");
            }
        }

        /// <summary>
        /// Returns average of values that meet a certain criterion
        /// </summary>
        /// <param name="numbers"></param>
        /// <param name="conditionValue"></param>
        /// <param name="conditionSign">"use one of: &lt;,&lt;=,&gt;,&gt;=,&lt;&gt;,="</param>
        /// <returns></returns>
        public double eAvgIf(double[] numbers, string conditionSign, double conditionValue)
        {
            if (conditionSign == "=")
            {
                return numbers.Sum(element => (element == conditionValue ? element : 0)) / numbers.Sum(element => (element == conditionValue ? 1 : 0));
            }
            else if (conditionSign == ">")
            {
                return numbers.Sum(element => (element > conditionValue ? element : 0)) / numbers.Sum(element => (element > conditionValue ? 1 : 0));
            }
            else if (conditionSign == ">=")
            {
                return numbers.Sum(element => (element >= conditionValue ? element : 0)) / numbers.Sum(element => (element >= conditionValue ? 1 : 0));
            }
            else if (conditionSign == "<")
            {
                return numbers.Sum(element => (element < conditionValue ? element : 0)) / numbers.Sum(element => (element < conditionValue ? 1 : 0));
            }
            else if (conditionSign == "<=")
            {
                return numbers.Sum(element => (element <= conditionValue ? element : 0)) / numbers.Sum(element => (element <= conditionValue ? 1 : 0));
            }
            else if (conditionSign == "<>")
            {
                return numbers.Sum(element => (element != conditionValue ? element : 0)) / numbers.Sum(element => (element != conditionValue ? 1 : 0));
            }
            else
            {
                throw new ArgumentException("wrong conditionSign, use one of: >, >=, <, <=, =, <> ");
            }
        }

        /// <summary>
        /// Returns average of values that meet a certain criterion (beetween value1 and value2)
        /// </summary>
        /// <param name="numbers"></param>
        /// <param name="conditionValue1"></param>
        /// <param name="conditionValue2"></param>
        /// <param name="between">"closed or open interval"</param>
        /// <returns></returns>
        public double eAvgIf(double[] numbers, double conditionValue1, double conditionValue2, string between = "closed")
        {
            if (between == "open")
            {
                return numbers.Sum(element => ((element > conditionValue1 && element < conditionValue2) ? element : 0)) / numbers.Sum(element => ((element > conditionValue1 && element < conditionValue2) ? 1 : 0));
            }
            else if (between == "closed")
            {
                return numbers.Sum(element => ((element >= conditionValue1 && element <= conditionValue2) ? element : 0)) / numbers.Sum(element => ((element >= conditionValue1 && element <= conditionValue2) ? 1 : 0));
            }
            else
            {
                throw new ArgumentException("wrong conditionSign, use one of: >, >=, <, <=, =, <> ");
            }
        }

        /// <summary>
        /// Counts values that meet a certain criterion
        /// </summary>
        /// <param name="numbers"></param>
        /// <param name="conditionValue"></param>
        /// <param name="conditionSign">"use one of: &lt;,&lt;=,&gt;,&gt;=,&lt;&gt;,="</param>
        /// <returns></returns>
        public double eCountIf(double[] numbers, string conditionSign, double conditionValue)
        {
            if (conditionSign == "=")
            {
                return numbers.Sum(element => (element == conditionValue ? 1 : 0));
            }
            else if (conditionSign == ">")
            {
                return numbers.Sum(element => (element > conditionValue ? 1 : 0));
            }
            else if (conditionSign == ">=")
            {
                return numbers.Sum(element => (element >= conditionValue ? 1 : 0));
            }
            else if (conditionSign == "<")
            {
                return numbers.Sum(element => (element < conditionValue ? 1 : 0));
            }
            else if (conditionSign == "<=")
            {
                return numbers.Sum(element => (element <= conditionValue ? 1 : 0));
            }
            else if (conditionSign == "<>")
            {
                return numbers.Sum(element => (element != conditionValue ? 1 : 0));
            }
            else
            {
                throw new ArgumentException("wrong conditionSign, use one of: >, >=, <, <=, =, <> ");
            }
        }

        /// <summary>
        /// Counts values that meet a certain criterion
        /// </summary>
        /// <param name="numbers"></param>
        /// <param name="conditionValue1"></param>
        /// <param name="conditionValue2"></param>
        /// <param name="between">"closed or open interval"</param>
        /// <returns></returns>
        public double eCountIf(double[] numbers, double conditionValue1, double conditionValue2, string between = "closed")
        {
            if (between == "open")
            {
                return numbers.Sum(element => ((element > conditionValue1 && element < conditionValue2) ? 1 : 0));
            }
            else if (between == "closed")
            {
                return numbers.Sum(element => ((element >= conditionValue1 && element <= conditionValue2) ? 1 : 0));
            }
            else
            {
                throw new ArgumentException("wrong conditionSign, use one of: >, >=, <, <=, =, <> ");
            }
        }
        /// <summary>
        /// Counts strings that equals a certain criterion
        /// </summary>
        /// <param name="strings"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public double eCountIf(string[] strings, string value)
        {
            return strings.Sum(element => (element == value ? 1 : 0));
        }
    }
}
