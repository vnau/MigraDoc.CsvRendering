#region MigraDoc - Creating Documents on the Fly
//
// Authors:
//   Stefan Lange (mailto:Stefan.Lange@pdfsharp.com)
//   Klaus Potzesny (mailto:Klaus.Potzesny@pdfsharp.com)
//   David Stephensen (mailto:David.Stephensen@pdfsharp.com)
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

using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using MigraDoc.DocumentObjectModel.Tables;

namespace MigraDoc.DocumentObjectModel.InternalVisitors
{
    /// <summary>
    /// Summary description for VisitorBase.
    /// </summary>
   internal abstract class VisitorBase : DocumentObjectVisitor
    {
        public VisitorBase()
        {
        }

        public override void Visit(DocumentObject documentObject)
        {
            IVisitable visitable = documentObject as IVisitable;
            if (visitable != null)
                visitable.AcceptVisitor(this, true);
        }

        protected void FlattenParagraphFormat(ParagraphFormat format, ParagraphFormat refFormat)
        {
        }

        protected void FlattenListInfo(ListInfo listInfo, ListInfo refListInfo)
        {
            //if (listInfo.continuePreviousList.IsNull)
            //  listInfo.continuePreviousList = refListInfo.continuePreviousList;
            //if (listInfo.listType.IsNull)
            //  listInfo.listType = refListInfo.listType;
            //if (listInfo.numberPosition.IsNull)
            //  listInfo.numberPosition = refListInfo.numberPosition;
        }

        protected void FlattenFont(Font font, Font refFont)
        {
        }

        protected void FlattenShading(Shading shading, Shading refShading)
        {
        }

        protected Border FlattenedBorderFromBorders(Border border, Borders parentBorders)
        {
            return border;
        }

        protected void FlattenBorders(Borders borders, Borders refBorders)
        {
        }

        protected void FlattenBorder(Border border, Border refBorder)
        {
        }

        protected void FlattenTabStops(TabStops tabStops, TabStops refTabStops)
        {
        }

        protected void FlattenPageSetup(PageSetup pageSetup, PageSetup refPageSetup)
        {
        }

        protected void FlattenHeaderFooter(HeaderFooter headerFooter, bool isHeader)
        {
        }

        protected void FlattenFillFormat(FillFormat fillFormat)
        {
        }

        protected void FlattenLineFormat(LineFormat lineFormat, LineFormat refLineFormat)
        {
        }

        protected void FlattenAxis(Axis axis)
        {
        }

        protected void FlattenPlotArea(PlotArea plotArea)
        {
        }

        protected void FlattenDataLabel(DataLabel dataLabel)
        {
        }


        #region Chart
        internal override void VisitChart(Chart chart)
        {
        }
        #endregion

        #region Document
        internal override void VisitDocument(Document document)
        {
        }

        internal override void VisitDocumentElements(DocumentElements elements)
        {
        }
        #endregion

        #region Format
        internal override void VisitStyle(Style style)
        {
        }

        internal override void VisitStyles(Styles styles)
        {
        }
        #endregion

        #region Paragraph
        internal override void VisitFootnote(Footnote footnote)
        {
        }

        internal override void VisitParagraph(Paragraph paragraph)
        {
        }
        #endregion

        #region Section
        internal override void VisitHeaderFooter(HeaderFooter headerFooter)
        {
        }

        internal override void VisitHeadersFooters(HeadersFooters headersFooters)
        {
        }

        internal override void VisitSection(Section section)
        {
        }

        internal override void VisitSections(Sections sections)
        {
        }
        #endregion

        #region Shape
        internal override void VisitTextFrame(TextFrame textFrame)
        {
        }
        #endregion

        #region Table
        internal override void VisitCell(Cell cell)
        {
            // format, shading and borders are already processed.
        }

        internal override void VisitColumns(Columns columns)
        {
        }

        internal override void VisitRow(Row row)
        {
        }

        internal override void VisitRows(Rows rows)
        {
        }
        /// <summary>
        /// Returns a paragraph format object initialized by the given style.
        /// It differs from style.ParagraphFormat if style is a character style.
        /// </summary>
        ParagraphFormat ParagraphFormatFromStyle(Style style)
        {
            return null;
        }

        internal override void VisitTable(Table table)
        {
        }
        #endregion


        internal override void VisitLegend(Legend legend)
        {
        }

        internal override void VisitTextArea(TextArea textArea)
        {
        }


        private DocumentObject GetDocumentElementHolder(DocumentObject docObj)
        {
            return null;
            //DocumentElements docEls = (DocumentElements)docObj.parent;
            //return docEls.parent;
        }
    }
}
