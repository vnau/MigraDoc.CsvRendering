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
using MigraDoc.DocumentObjectModel.Internals;
using MigraDoc.DocumentObjectModel.InternalVisitors;
using System;
using System.Collections;
using System.IO;

namespace MigraDoc.CsvRendering
{
    /// <summary>
    /// Class to render a MigraDoc document to CSV format.
    /// </summary>
    public class CsvDocumentRenderer : RendererBase
    {
        /// <summary>
        /// Initializes a new instance of the DocumentRenderer class.
        /// </summary>
        public CsvDocumentRenderer()
        {
        }

        /// <summary>
        /// This function is declared only for technical reasons!
        /// </summary>
        internal override void Render()
        {
            throw new InvalidOperationException();
        }

        /// <summary>
        /// Renders a MigraDoc document to the specified file.
        /// </summary>
        public void Render(Document doc, string file)
        {
            StreamWriter strmWrtr = null;
            try
            {
                this.document = doc;
                this.docObject = doc;
                string path = file;

                strmWrtr = new StreamWriter(path, false, System.Text.Encoding.Default);
                this.csvWriter = new CsvWriter(strmWrtr);
                WriteDocument();
            }
            finally
            {
                if (strmWrtr != null)
                {
                    strmWrtr.Flush();
                    strmWrtr.Close();
                }
            }
        }

        /// <summary>
        /// Renders a MigraDoc document to the specified stream.
        /// </summary>
        public void Render(Document document, Stream stream)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            StreamWriter strmWrtr = null;
            try
            {
                strmWrtr = new StreamWriter(stream, System.Text.Encoding.Default);
                this.document = document;
                this.docObject = document;
                this.csvWriter = new CsvWriter(strmWrtr);
                WriteDocument();
            }
            finally
            {
                if (strmWrtr != null)
                {
                    strmWrtr.Flush();
                    strmWrtr.Close();
                }
            }
        }

        /// <summary>
        /// Renders a MigraDoc to CSV and returns the result as string.
        /// </summary>
        public string RenderToString(Document document)
        {
            if (document == null)
                throw new ArgumentNullException("document");

            this.document = document;
            this.docObject = document;
            StringWriter writer = null;
            try
            {
                writer = new StringWriter();
                this.csvWriter = new CsvWriter(writer);
                WriteDocument();
                writer.Flush();
                return writer.GetStringBuilder().ToString();
            }
            finally
            {
                if (writer != null)
                    writer.Close();
            }
        }

        /// <summary>
        /// Renders a MigraDoc document with help of the internal RtfWriter.
        /// </summary>
        private void WriteDocument()
        {
            CsvFlattenVisitor flattener = new CsvFlattenVisitor();
            flattener.Visit(this.document);
            Prepare();
            this.csvWriter.StartContent();
            RenderHeader();
            RenderDocumentArea();
            this.csvWriter.EndContent();
        }

        /// <summary>
        /// Prepares this renderer by collecting Information for font and color table.
        /// </summary>
        private void Prepare()
        {
            this.listList.Clear();
            ListInfoRenderer.Clear();
            ListInfoOverrideRenderer.Clear();
            CollectTables(this.document);
        }

        /// <summary>
        /// Renders the RTF Header.
        /// </summary>
        private void RenderHeader()
        {
            //Document properties can occur before and between the header tables.
            //Lists are not yet implemented.
            RenderListTable();
        }

        /// <summary>
        /// Fills the font, color and (later!) list hashtables so they can be rendered and used by other renderers.
        /// </summary>
        private void CollectTables(DocumentObject dom)
        {
            ValueDescriptorCollection vds = Meta.GetMeta(dom).ValueDescriptors;
            int count = vds.Count;
            for (int idx = 0; idx < count; idx++)
            {
                ValueDescriptor vd = vds[idx];
                if (!vd.IsRefOnly && !vd.IsNull(dom))
                {
                    if (vd.ValueType == typeof(Color))
                    {
                    }
                    else if (vd.ValueType == typeof(Font))
                    {
                    }
                    else if (vd.ValueType == typeof(ListInfo))
                    {
                        ListInfo lst = vd.GetValue(dom, GV.ReadWrite) as ListInfo; //ReadOnly
                        if (!this.listList.Contains(lst))
                            this.listList.Add(lst);
                    }
                    if (typeof(DocumentObject).IsAssignableFrom(vd.ValueType))
                    {
                        CollectTables(vd.GetValue(dom, GV.ReadWrite) as DocumentObject); //ReadOnly
                        if (typeof(DocumentObjectCollection).IsAssignableFrom(vd.ValueType))
                        {
                            DocumentObjectCollection coll = vd.GetValue(dom, GV.ReadWrite) as DocumentObjectCollection; //ReadOnly
                            if (coll != null)
                            {
                                foreach (DocumentObject obj in coll)
                                {
                                    //In SeriesCollection kann null vorkommen.
                                    if (obj != null)
                                        CollectTables(obj);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Renders the list hashtable within the RTF header.
        /// </summary>
        private void RenderListTable()
        {
            if (this.listList.Count == 0)
                return;

            csvWriter.StartContent();
            foreach (object obj in this.listList)
            {
                ListInfo lst = (ListInfo)obj;
                ListInfoRenderer lir = new ListInfoRenderer(lst, this);
                lir.Render();
            }
            csvWriter.EndContent();

            csvWriter.StartContent();
            foreach (object obj in this.listList)
            {
                ListInfo lst = (ListInfo)obj;
                ListInfoOverrideRenderer lir =
                    new ListInfoOverrideRenderer(lst, this);
                lir.Render();
            }
            csvWriter.EndContent();
        }

        /// <summary>
        /// Renders the RTF document area, which is all except the header.
        /// </summary>
        private void RenderDocumentArea()
        {
            RenderInfo();
            foreach (Section sect in this.document.Sections)
            {
                RendererFactory.CreateRenderer(sect, this).Render();
            }
        }

        /// <summary>
        /// Renders global document properties, such as mirror margins and unicode treatment.
        /// Note that a section specific margin mirroring does not work in Word.
        /// </summary>
        private void RenderGlobalPorperties()
        {
        }

        /// <summary>
        /// Renders the document information of title, author etc..
        /// </summary>
        private void RenderInfo()
        {
            if (document.IsNull("Info"))
                return;

            this.csvWriter.StartContent();

            DocumentInfo info = this.document.Info;
            /*
            if (!info.IsNull("Title"))
            {
                this.csvWriter.StartContent();
                this.csvWriter.WriteControl("title");
                this.csvWriter.WriteText(info.Title);
                this.csvWriter.EndContent();
            }
            if (!info.IsNull("Subject"))
            {
                this.csvWriter.StartContent();
                this.csvWriter.WriteControl("subject");
                this.csvWriter.WriteText(info.Subject);
                this.csvWriter.EndContent();
            }
            if (!info.IsNull("Author"))
            {
                this.csvWriter.StartContent();
                this.csvWriter.WriteControl("author");
                this.csvWriter.WriteText(info.Author);
                this.csvWriter.EndContent();
            }
            if (!info.IsNull("Keywords"))
            {
                this.csvWriter.StartContent();
                this.csvWriter.WriteControl("keywords");
                this.csvWriter.WriteText(info.Keywords);
                this.csvWriter.EndContent();
            }
            */
            this.csvWriter.EndContent();
        }

        /// <summary>
        /// Gets the MigraDoc document that is currently rendered.
        /// </summary>
        internal Document Document
        {
            get { return this.document; }
        }
        private Document document;

        /// <summary>
        /// Gets the RtfWriter the document is rendered with.
        /// </summary>
        internal CsvWriter CsvWriter
        {
            get { return this.csvWriter; }
        }

        private ArrayList listList = new ArrayList();
    }
}
