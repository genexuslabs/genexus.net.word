using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;

namespace Genexus.Word.Shapes
{
    public abstract class CustomShape
    {
        public uint Id { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }
        public string Text { get; set; }

        protected Dictionary<string, string> RequiredImports = new Dictionary<string, string>();

        public virtual OpenXmlCompositeElement Build()
        {
            throw new NotImplementedException();
        }

        public void AddRequiredNamespaces(MainDocumentPart documentPart)
        {
            foreach (var item in RequiredImports)
            {
                documentPart.Document.RemoveNamespaceDeclaration(item.Key);
                documentPart.Document.AddNamespaceDeclaration(item.Key, item.Value);
            }
        }
    }
}
