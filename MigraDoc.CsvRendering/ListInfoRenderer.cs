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

using MigraDoc.DocumentObjectModel;
using System.Collections;
using System.Collections.Generic;

namespace MigraDoc.CsvRendering
{
    /// <summary>
    /// ListInfoRenderer.
    /// </summary>
    internal class ListInfoRenderer : RendererBase
    {
        public ListInfoRenderer(DocumentObject domObj, CsvDocumentRenderer docRenderer)
          : base(domObj, docRenderer)
        {
            this.listInfo = domObj as ListInfo;
        }

        public static void Clear()
        {
            idList.Clear();
            listID = 1;
            templateID = 2;
        }
        /// <summary>
        /// Renders a ListIfo to CSV.
        /// </summary>
        internal override void Render()
        {
            if (prevListInfoID.Key != null && listInfo.ContinuePreviousList)
            {
                idList.Add(this.listInfo, prevListInfoID.Value);
                return;
            }
            idList.Add(listInfo, listID);

            csvWriter.StartContent();
            // rtfWriter.WriteControl("listtemplateid", templateID.ToString());
            WriteListLevel();
            // csvWriter.WriteControl("listid", listID.ToString(CultureInfo.InvariantCulture));
            this.csvWriter.EndContent();

            prevListInfoID = new KeyValuePair<ListInfo, int>(listInfo, listID);
            listID += 2;
            templateID += 2;
        }
        private static KeyValuePair<ListInfo, int> prevListInfoID = new KeyValuePair<ListInfo, int>();
        private ListInfo listInfo;
        private static int listID = 1;
        private static int templateID = 2;
        private static Hashtable idList = new Hashtable();

        /// <summary>
        /// Gets the corresponding List ID of the ListInfo Object.
        /// </summary>
        internal static int GetListID(ListInfo li)
        {
            if (idList.ContainsKey(li))
                return (int)idList[li];

            return -1;
        }

        private void WriteListLevel()
        {
            ListType listType = this.listInfo.ListType;
            string levelText1 = "";
            string levelText2 = "";
            string levelNumbers = "";
            switch (listType)
            {
                case ListType.NumberList1:
                    levelText1 = "'02";
                    levelText2 = "'00.";
                    levelNumbers = "'01";
                    break;

                case ListType.NumberList2:
                case ListType.NumberList3:
                    levelText1 = "'02";
                    levelText2 = "'00)";
                    levelNumbers = "'01";
                    break;

                //levelText1 = "'02";
                //levelText2 = "'00)";
                //levelNumbers = "'01";
                //break;

                case ListType.BulletList1:
                    levelText1 = "'01";
                    levelText2 = "u-3913 ?";
                    break;

                case ListType.BulletList2:
                    levelText1 = "'01o";
                    levelText2 = "";
                    break;

                case ListType.BulletList3:
                    levelText1 = "'01";
                    levelText2 = "u-3929 ?";
                    break;
            }
            WriteListLevel(levelText1, levelText2, levelNumbers);
        }

        private void WriteListLevel(string levelText1, string levelText2, string levelNumbers)
        {
            csvWriter.StartContent();
            // Start
            //Translate("ListType", "levelnfcn", RtfUnit.Undefined, "4", false);
            //Translate("ListType", "levelnfc", RtfUnit.Undefined, "4", false);


            csvWriter.StartContent();

            if (levelText2 != "")
            {
                //   csvWriter.WriteControl(levelText2);
            }

            this.csvWriter.WriteSeparator();

            csvWriter.EndContent();
            csvWriter.StartContent();
            //csvWriter.WriteControl("levelnumbers");
            if (levelNumbers != "")
            {
                //   csvWriter.WriteControl(levelNumbers);
            }

            this.csvWriter.WriteSeparator();
            csvWriter.EndContent();

            //csvWriter.WriteControl("levelfollow", 0);

            csvWriter.EndContent();
        }
    }
}
