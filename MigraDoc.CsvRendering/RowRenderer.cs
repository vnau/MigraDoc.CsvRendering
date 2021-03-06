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
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Visitors;

namespace MigraDoc.CsvRendering
{
    /// <summary>
    /// Class to render a Row to CSV.
    /// </summary>
    internal class RowRenderer : RendererBase
    {
        internal RowRenderer(DocumentObject domObj, CsvDocumentRenderer docRenderer)
          : base(domObj, docRenderer)
        {
            this.row = domObj as Row;
        }

        /// <summary>
        /// Render a Row to CSV.
        /// </summary>
        internal override void Render()
        {
            this.useEffectiveValue = true;
            new RowsRenderer(DocumentRelations.GetParent(this.row) as Rows, this.docRenderer).Render();

            int thisRowIdx = this.row.Index;

            foreach (Cell cell in this.row.Cells)
            {
                CellRenderer cellRndrr = new CellRenderer(cell, this.docRenderer);
                cellRndrr.CellList = this.cellList;
                cellRndrr.Render();
            }

            this.csvWriter.WriteNewLine();
        }

        /// <summary>
        /// Sets the merged cell list. This property is set by the table renderer.
        /// </summary>
        internal MergedCellList CellList
        {
            set
            {
                this.cellList = value;
            }
        }
        MergedCellList cellList = null;
        Row row;
    }
}
