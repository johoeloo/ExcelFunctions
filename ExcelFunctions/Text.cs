using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFunctions
{
    class Text
    {
        /// <summary>
        /// Extracts the substring from the middle of a text string 
        /// </summary>
        /// <param name="text"></param>
        /// <param name="startIndex"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public string eMid(string text, int startIndex, int length)
        {
            if (startIndex < 0)
                return "Error, startIndex < 0";
            else if (length < 0)
                return "Error, length < 0";
            else if (text.Length < startIndex)
                return "Error, startIndex value is too high";
            else if (text.Length < length || text.Length < length - startIndex)
            {
                length = text.Length - startIndex;
                return text.Substring(startIndex, length);
            }
            else
                return text.Substring(startIndex, length);
        }

        /// <summary>
        /// Extracts the substring from the middle of a text string
        /// </summary>
        /// <param name="integer"></param>
        /// <param name="startIndex"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public string eMid(int integer, int startIndex, int length)
        {
            string text = integer.ToString();
            if (startIndex < 0)
                return "Error, startIndex < 0";
            else if (length < 0)
                return "Error, length < 0";
            else if (text.Length < startIndex)
                return "Error, startIndex value is too high";
            else if (text.Length < length || text.Length < length - startIndex)
            {
                length = text.Length - startIndex;
                return text.Substring(startIndex, length);
            }
            else
                return text.Substring(startIndex, length);
        }

        /// <summary>
        /// Extracts a given number of characters from the left side of a text string
        /// </summary>
        /// <param name="text"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public string eLeft(string text, int length)
        {
            if (length < 0)
                return "Error, length < 0";
            if (text.ToString().Length < length)
                return "Error, length value is too high";
            else
                return text.Substring(1, length);
        }

        /// <summary>
        /// Extracts a given number of characters from the left side of a text string
        /// </summary>
        /// <param name="text"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public string eLeft(int text, int length)
        {
            if (length < 0)
                return "Error, length < 0";
            else if (text.ToString().Length < length)
                return "Error, length value is too high";
            else
                return text.ToString().Substring(1, length);
        }

        /// <summary>
        /// Extracts a given number of characters from the right side of a text string
        /// </summary>
        /// <param name="text"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public string eRight(string text, int length)
        {
            if (length < 0)
                return "Error, length < 0";
            else if (text.Length < length)
                return "Error, length value is too high";
            else
                return text.Substring(text.Length - length, length);
        }

        /// <summary>
        /// Extracts a given number of characters from the right side of a text string
        /// </summary>
        /// <param name="text"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public string eRight(int text, int length)
        {
            if (length < 0)
                return "Error, length < 0";
            else if (text.ToString().Length < length)
                return "Error, length value is too high";
            else
                return text.ToString().Substring(text.ToString().Length - length, length);
        }

        /// <summary>
        /// Returns length of a text string
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public int eLength(string text)
        {
            return text.Length;
        }

        /// <summary>
        /// Returns length of a text string
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public int eLength(int text)
        {
            return text.ToString().Length;
        }

        /// <summary>
        ///  Converts text string to all lower-case
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public string eLower(string text)
        {
            return text.ToLower();
        }

        /// <summary>
        /// Converts text string to all upper-case
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public string eUpper(string text)
        {
            return text.ToUpper();
        }

        /// <summary>
        /// Capitalize first letters in words or sentences
        /// </summary>
        /// <param name="text">string to capitalize letters</param>
        /// <param name="word">capitalize first letter in each: true - word, false - sentence</param>
        /// <returns></returns>
        public string eProper(string text, bool word = true)
        {
            string result = "";
            if (word)
            {
                result = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(text.ToLower());
            }
            else
            {
                bool newWord = true;
                char last = 'x';
                foreach (char c in text)
                {
                    if (newWord)
                    {
                        result = result + Char.ToUpper(c);
                        newWord = false;
                    }
                    else
                        result = result + Char.ToLower(c);

                    if (c == ' ' & last == '.') newWord = true;
                    last = c;
                }
            }
            return result;
        }

        /// <summary>
        /// Replaces a set of characters with another.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public string eSubstitute(string text, string oldValue, string newValue)
        {
            return text.Replace(oldValue, newValue);
        }

        /// <summary>
        /// Replaces a set of characters with another.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public string eSubstitute(long text, int oldValue, int newValue)
        {
            return text.ToString().Replace(oldValue.ToString(), newValue.ToString());
        }

        /// <summary>
        /// Replaces a set of characters with another.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public string eSubstitute(int text, int oldValue, string newValue)
        {
            return text.ToString().Replace(oldValue.ToString(), newValue);
        }

        /// <summary>
        /// Replaces a set of characters with another.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public string eSubstitute(string text, string oldValue, int newValue)
        {
            return text.Replace(oldValue, newValue.ToString());
        }

        /// <summary>
        /// Removes all spaces from text except for single spaces between words
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public string eTrim(string text)
        {
            return text.Trim().Replace("  ", "");
        }

        /// <summary>
        ///  Joins two or more text strings into one
        /// </summary>
        /// <param name="separator"></param>
        /// <param name="strings"></param>
        /// <returns></returns>
        public string eConcatenate(string separator, params object[] strings)
        {
            return String.Join(separator, strings);
        }
    }
}

