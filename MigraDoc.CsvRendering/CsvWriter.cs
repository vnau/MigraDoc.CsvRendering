#region MigraDoc - Creating Documents on the Fly
//
// Authors:
//   Klaus Potzesny (mailto:Klaus.Potzesny@pdfsharp.com)
//
// Copyright (c) 2001-2009 empira Software GmbH, Cologne (Germany)
//
// http://www.pdfsharp.com
// http://www.migradoc.com
// http://sourceforge.net/projects/pdfsharp
//
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included
// in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
// DEALINGS IN THE SOFTWARE.
#endregion

using System.IO;

namespace MigraDoc.CsvRendering
{
    /// <summary>
    /// Class to write RTF output.
    /// </summary>
    internal class CsvWriter
    {
        /// <summary>
        /// Initializes a new instance of the RtfWriter class.
        /// </summary>
        public CsvWriter(TextWriter textWriter)
        {
            this.textWriter = textWriter;
        }

        /// <summary>
        /// Writes a left brace.
        /// </summary>
        public void StartContent()
        {
            this.lastWasControl = false;
        }

        /// <summary>
        /// Writes a right brace.
        /// </summary>
        public void EndContent()
        {
            this.lastWasControl = false;
        }

        /// <summary>
        /// Writes the given text, handling special characters before.
        /// </summary>
        public void WriteText(string text)
        {
            if (this.lastWasControl)
                this.textWriter.Write(" ");
            //    strBuilder.Append(" ");
            //StringBuilder strBuilder = new StringBuilder(text.Length);
            //if (this.lastWasControl)
            //    strBuilder.Append(" ");

            //int lengh = text.Length;
            //for (int idx = 0; idx < lengh; idx++)
            //{
            //    char ch = text[idx];
            //    switch (ch)
            //    {
            //        case '\\':
            //            strBuilder.Append(@"\\");
            //            break;

            //        case '{':
            //            strBuilder.Append(@"\{");
            //            break;

            //        case '}':
            //            strBuilder.Append(@"\}");
            //            break;

            //        case '­': //character 173, softhyphen
            //            strBuilder.Append(@"\-");
            //            break;

            //        default:
            //            if (IsCp1252Char(ch))
            //                strBuilder.Append(ch);
            //            else
            //            {
            //                strBuilder.Append(@"\u");
            //                strBuilder.Append(((int)ch).ToString(CultureInfo.InvariantCulture));
            //                strBuilder.Append('?');
            //            }
            //            break;
            //    }
            //}
            //this.textWriter.Write(strBuilder.ToString());
            //lastWasControl = false;

            text.Replace("\"", "\"\""); // Replace quotes with doublequotes
            if (text.Contains(Separator))
                text = "\"" + text + "\"";
            this.textWriter.Write(text);

        }

        ///// <summary>
        ///// Writes the number as hex value. Only numbers &lt;= 255 are allowed.
        ///// </summary>
        //public void WriteHex(uint hex)
        //{
        //    if (hex > 0xFF)
        //        //TODO: Fehlermeldung
        //        return;

        //    this.textWriter.Write(@"\'" + hex.ToString("x"));
        //    lastWasControl = false;
        //    //Dahinter darf kein zusätzliches Blank stehen.
        //}

        /// <summary>
        /// Writes a blank in paragraph text.
        /// </summary>
        public void WriteBlank()
        {
            this.textWriter.Write(" ");
        }

        private string Separator
        {
            get
            {
                return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
                //return ",";
            }
        }

        /// <summary>
        /// Writes a semicolon as separator e.g. in in font tables.
        /// </summary>
        public void WriteSeparator()
        {

            this.textWriter.Write(Separator);
            lastWasControl = false;
        }

        /// <summary>
        /// Writes a new line.
        /// </summary>
        public void WriteNewLine()
        {
            this.textWriter.Write("\r\n");
            lastWasControl = false;
        }

        private bool lastWasControl;// = false;
        private TextWriter textWriter;
    }
}
