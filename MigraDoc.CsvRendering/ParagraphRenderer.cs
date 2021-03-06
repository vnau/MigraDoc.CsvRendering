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

namespace MigraDoc.CsvRendering
{
    /// <summary>
    /// Class to render a Paragraph to CSV.
    /// </summary>
    internal class ParagraphRenderer : RendererBase
    {
        public ParagraphRenderer(DocumentObject domObj, CsvDocumentRenderer docRenderer)
          : base(domObj, docRenderer)
        {
            this.paragraph = domObj as Paragraph;
        }

        /// <summary>
        /// Renders the paragraph to CSV.
        /// </summary>
        internal override void Render()
        {
            useEffectiveValue = true;
            DocumentElements elements = DocumentRelations.GetParent(this.paragraph) as DocumentElements;


            bool isCellParagraph = DocumentRelations.GetParent(elements) is Cell;
            bool isFootnoteParagraph = isCellParagraph ? false : DocumentRelations.GetParent(elements) is Footnote;

            if (!this.paragraph.IsNull("Elements"))
                RenderContent();

            if ((!isCellParagraph && !isFootnoteParagraph) || this.paragraph != elements.LastObject)
                this.csvWriter.WriteNewLine();
        }

        /// <summary>
        /// Renders the paragraph content to CSV.
        /// </summary>
        private void RenderContent()
        {
            DocumentElements elements = DocumentRelations.GetParent(this.paragraph) as DocumentElements;
            //First paragraph of a footnote writes the reference symbol:
            if (DocumentRelations.GetParent(elements) is Footnote && this.paragraph == elements.First)
            {
                FootnoteRenderer ftntRenderer = new FootnoteRenderer(DocumentRelations.GetParent(elements) as Footnote, this.docRenderer);
                ftntRenderer.RenderReference();
            }
            foreach (DocumentObject docObj in this.paragraph.Elements)
            {
                if (docObj == paragraph.Elements.LastObject)
                {
                    if (docObj is Character)
                    {
                        if (((Character)docObj).SymbolName == SymbolName.LineBreak)
                            continue; //Ignore last linebreak.
                    }
                }
                RendererBase rndrr = RendererFactory.CreateRenderer(docObj, this.docRenderer);
                if (rndrr != null)
                    rndrr.Render();
            }
        }
        Paragraph paragraph;
    }
}
