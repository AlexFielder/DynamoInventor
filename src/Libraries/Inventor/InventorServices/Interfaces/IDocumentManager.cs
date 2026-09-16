using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Inventor;

namespace InventorServices.Persistence
{
    public interface IDocumentManager
    {
        AssemblyDocument ActiveAssemblyDoc { get; set; }
        DrawingDocument ActiveDrawingDoc { get; set; }
        PartDocument ActivePartDoc { get; set; }
        Document ActiveDocument { get; }

        /// <summary>
        /// The component definition of the active part or assembly. Creates a new assembly when no
        /// document is open; throws InvalidOperationException for drawings and presentations.
        /// </summary>
        ComponentDefinition ActiveComponentDefinition { get; }

        /// <summary>Reference-key manager of the document behind <see cref="ActiveComponentDefinition"/>.</summary>
        ReferenceKeyManager ActiveReferenceKeyManager { get; }
        Inventor.Application InventorApplication { get; set; }
        ApprenticeServerComponent ActiveApprenticeServer { get; set; }             
    }
}
