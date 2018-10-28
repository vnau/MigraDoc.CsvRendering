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

using System;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.IO;

namespace MigraDoc.CsvRendering
{
  /// <summary>
  /// Renders formatted text to CSV.
  /// </summary>
  internal class FormattedTextRenderer : RendererBase
  {
    internal FormattedTextRenderer(DocumentObject domObj, CsvDocumentRenderer docRenderer)
      : base(domObj, docRenderer)
    {
      this.formattedText = domObj as FormattedText;
    }

    /// <summary>
    /// Renders a formatted text to CSV.
    /// </summary>
    internal override void Render()
    {
      useEffectiveValue = true;

      this.csvWriter.StartContent();

      foreach (DocumentObject docObj in this.formattedText.Elements)
        RendererFactory.CreateRenderer(docObj, this.docRenderer).Render();

      this.csvWriter.EndContent();
    }

    private FormattedText formattedText;
  }
}