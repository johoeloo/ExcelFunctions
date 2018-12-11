using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelFunctions
{
    class Find
    {        /// <summary>
             /// Returns the position of item in an array or in string
             /// </summary>
             /// <param name="find"></param>
             /// <param name="tab"></param>
             /// <returns></returns>
        public string eMatch(string find, string[] tab)
        {
            if (Array.IndexOf(tab, find) < 0)
                return "Error, can't find text";

            return (Array.IndexOf(tab, find) + 1).ToString();
        }

        /// <summary>
        /// Returns the position of item in an array or in string
        /// </summary>
        /// <param name="find"></param>
        /// <param name="tab"></param>
        /// <returns></returns>
        public string eMatch(char find, string text)
        {
            if (text.IndexOf(find) < 0)
                return "Error, can't find text";

            return (text.IndexOf(find) + 1).ToString();
        }

        /// <summary>
        /// Returns the position of item in an array or in string
        /// </summary>
        /// <param name="find"></param>
        /// <param name="tab"></param>
        /// <returns></returns>
        public string eMatch(string find, string text)
        {
            if (text.IndexOf(find) < 0)
                return "Error, can't find text";

            return (text.IndexOf(find) + 1).ToString();
        }

        /// <summary>
        /// Returns the position of item in an array
        /// </summary>
        /// <param name="find"></param>
        /// <param name="tab"></param>
        /// <returns></returns>
        public string eMatch(int find, int[] tab)
        {
            if (Array.IndexOf(tab, find) < 0)
                return "Error, can't find text";

            return (Array.IndexOf(tab, find) + 1).ToString();
        }

        /// <summary>
        /// Returns the position of item in an array
        /// </summary>
        /// <param name="find"></param>
        /// <param name="tab"></param>
        /// <returns></returns>
        public string eMatch(int find, int text)
        {
            if (text.ToString().IndexOf(find.ToString()) < 0)
                return "Error, can't find text";

            return (text.ToString().IndexOf(find.ToString()) + 1).ToString();
        }
        /// <summary>
        /// Returns a string that is a specified number of strings from a string. 
        /// </summary>
        /// <param name="find"></param>
        /// <param name="offset"></param>
        /// <param name="tab"></param>
        /// <returns></returns>
        public string eOffset(string find, int offset, string[] tab)
        {
            if (offset < 0)
                return "Error, offset < 0";
            else if (offset + Array.IndexOf(tab, find) > tab.Count())
                return "Error, offset value is to high";
            else
                return tab[offset + Array.IndexOf(tab, find)];
        }

        /// <summary>
        /// Returns a string that is a specified number of strings from a string. 
        /// </summary>
        /// <param name="find"></param>
        /// <param name="offset"></param>
        /// <param name="tab"></param>
        /// <returns></returns>
        public string eOffset(int find, int offset, int[] tab)
        {
            if (offset < 0)
                return "Error, offset < 0";
            else if (offset + Array.IndexOf(tab, find) > tab.Count())
                return "Error, offset value is to high";
            else
                return tab[offset + Array.IndexOf(tab, find)].ToString();
        }

    }
}
