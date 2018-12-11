using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFunctions
{
    class Logic
    {
        /// <summary>
        /// Returns TRUE if the expression is a null value
        /// </summary>
        /// <param name="check"></param>
        /// <returns></returns>
        public bool eIsNull(object check)
        {
            return (check == null);
        }

        /// <summary>
        /// Returns TRUE if the expression is a numeric value
        /// </summary>
        /// <param name="check"></param>
        /// <returns></returns>
        public bool eIsNumber(object check)
        {
            try
            {
                Convert.ToInt32(check);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Returns TRUE if the expression is a even number
        /// </summary>
        /// <param name="check"></param>
        /// <returns></returns>
        public bool eEven(int check)
        {
            return (check % 2 == 0);
        }
    }
}
