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

using MigraDoc.CsvRendering.Resources;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using System.Diagnostics;
namespace MigraDoc.CsvRendering
{
    /// <summary>
    /// Class to render a TextFrame to CSV.
    /// </summary>
    internal class TextFrameRenderer : ShapeRenderer
    {
        internal TextFrameRenderer(DocumentObject domObj, CsvDocumentRenderer docRenderer) : base(domObj, docRenderer)
        {
            this.textFrame = domObj as TextFrame;
        }

        /// <summary>
        /// Renders a TextFrame to CSV.
        /// </summary>
        internal override void Render()
        {
            DocumentElements elms = DocumentRelations.GetParent(this.textFrame) as DocumentElements;
            bool renderInParagraph = RenderInParagraph();
            if (renderInParagraph)
                StartDummyParagraph();

            StartShapeArea();

            //Properties
            RenderNameValuePair("shapeType", "202");//202 entspr. Textrahmen.

            if (this.textFrame.IsNull("Elements") ||
                !CollectionContainsObjectAssignableTo(this.textFrame.Elements,
                   typeof(Shape), typeof(Table)))
            {
            }
            else
            {
                TextOrientation orient = this.textFrame.Orientation;
                if (orient != TextOrientation.Horizontal && orient != TextOrientation.HorizontalRotatedFarEast)
                    Trace.WriteLine(Messages.TextframeContentsNotTurned, "warning");
            }
            csvWriter.StartContent();
            csvWriter.StartContent();
            foreach (DocumentObject docObj in this.textFrame.Elements)
            {
                RendererBase rndrr = RendererFactory.CreateRenderer(docObj, this.docRenderer);
                if (rndrr != null)
                {
                    rndrr.Render();
                }
            }
            //Text fields need to close with a paragraph.
            RenderTrailingParagraph(this.textFrame.Elements);

            this.csvWriter.EndContent();
            this.csvWriter.EndContent();
            EndShapeArea();
            if (renderInParagraph)
            {
                RenderLayoutPicture();
                EndDummyParagraph();
            }
        }


        /// <summary>
        /// Gets the user defined shape height if given, else 1 inch.
        /// </summary>
        protected override Unit GetShapeHeight()
        {
            if (this.textFrame.IsNull("Height"))
                return Unit.FromInch(1.0);

            return base.GetShapeHeight();
        }


        /// <summary>
        /// Gets the user defined shape width if given, else 1 inch.
        /// </summary>
        protected override Unit GetShapeWidth()
        {
            if (this.textFrame.IsNull("Width"))
                return Unit.FromInch(1.0);

            return base.GetShapeWidth();
        }

        /// <summary>
        /// Renders an empty dummy picture that allows the textframe to be placed in the dummy paragraph.
        /// (A bit obscure, but the only possiblity.)
        /// </summary>
        private void RenderLayoutPicture()
        {
            this.csvWriter.StartContent();
            this.csvWriter.StartContent();
            RenderNameValuePair("shapeType", "75");
            RenderNameValuePair("fPseudoInline", "1");
            RenderNameValuePair("fLockPosition", "1");
            RenderNameValuePair("fLockRotation", "1");
            this.csvWriter.EndContent();
            this.csvWriter.EndContent();
        }

        /// <summary>
        /// Starts a dummy paragraph to put a shape in, which is wrapped TopBottom style.
        /// </summary>
        protected override void StartDummyParagraph()
        {
            base.StartDummyParagraph();
            this.csvWriter.StartContent();
            this.csvWriter.StartContent();
            this.csvWriter.StartContent();
            this.csvWriter.WriteText("SHAPE");
            this.csvWriter.WriteText(@" \*MERGEFORMAT");
            this.csvWriter.EndContent();
            this.csvWriter.EndContent();
            this.csvWriter.StartContent();
        }

        /// <summary>
        /// Ends a dummy paragraph to put a shape in, which is wrapped TopBottom style.
        /// </summary>
        protected override void EndDummyParagraph()
        {
            this.csvWriter.EndContent();
            this.csvWriter.EndContent();
            base.EndDummyParagraph();
        }

        private TextFrame textFrame;
    }
}
