using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Inventor;

namespace InventorServices.Persistence
{
    public class DocumentManager : IDocumentManager
    {
        private Inventor.ApplicationEvents appEvents; 
        private ApprenticeServerComponentClass apprenticeServer;
        private Inventor.Application invApp;
        private AssemblyDocument assDoc;
        private PartDocument partDoc;

        public DocumentManager()
        {
            appEvents = InventorApplication.ApplicationEvents;
            appEvents.OnActivateDocument += appEvents_OnActivateDocument;
            appEvents.OnDeactivateDocument += appEvents_OnDeactivateDocument;
        }

        public Inventor.AssemblyDocument ActiveAssemblyDoc
        {
            get
            {
                //TODO: This is not good.  The convention in the DynamoInventor is that if we don't have an active
                //assembly document, we set this to null.  This is just for RTC demo.  Change this back.
                if (assDoc == null && ActiveDocument != null && ActiveDocument.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    assDoc = (AssemblyDocument)ActiveDocument;
                }
                else if (ActiveDocument == null)
                {
                    assDoc = (AssemblyDocument)InventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject);
                }
                return assDoc;
            }
            set { assDoc = value; }
        }

        public Inventor.DrawingDocument ActiveDrawingDoc { get; set; }

        public Inventor.PartDocument ActivePartDoc
        {
            get
            {
                //TODO: This is not good.  The convention in the DynamoInventor is that if we don't have an active
                //part document, we set this to null.  This is just for RTC demo.  Change this back.
                if (partDoc == null && ActiveDocument != null && ActiveDocument.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    partDoc = (PartDocument)ActiveDocument;
                }
                return partDoc;
            }
            set { partDoc = value; }
        }

        public Inventor.Document ActiveDocument
        {
            get { return InventorApplication.ActiveDocument; }
        }

        /// <summary>
        /// Part or assembly component definition of the active document. Nodes that create geometry go
        /// through this rather than ActiveAssemblyDoc/ActivePartDoc so they work in both document types.
        /// Assemblies only accept fixed work points/planes; parts accept every definition type.
        /// </summary>
        public Inventor.ComponentDefinition ActiveComponentDefinition
        {
            get
            {
                var doc = ActiveDocument;
                if (doc == null)
                {
                    // Nothing open: keep the historical behaviour and start an assembly.
                    doc = InventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject);
                }

                switch (doc.DocumentType)
                {
                    case DocumentTypeEnum.kPartDocumentObject:
                        return (ComponentDefinition)((PartDocument)doc).ComponentDefinition;
                    case DocumentTypeEnum.kAssemblyDocumentObject:
                        return (ComponentDefinition)((AssemblyDocument)doc).ComponentDefinition;
                    default:
                        throw new InvalidOperationException(
                            "The active Inventor document is a " + doc.DocumentType + "; open a part or an assembly.");
                }
            }
        }

        public Inventor.ReferenceKeyManager ActiveReferenceKeyManager
        {
            get
            {
                // Resolve the component definition first so the no-document case creates the assembly.
                var definition = ActiveComponentDefinition;
                return ((Document)definition.Document).ReferenceKeyManager;
            }
        }

        public Inventor.Application InventorApplication
        {
            get
            {
                if (invApp == null)
                {
                    try
                    {
                        invApp = (Inventor.Application)InventorServices.Utilities.ComInterop.GetActiveObject("Inventor.Application");
                    }
                    catch
                    {
                        Type invAppType = System.Type.GetTypeFromProgID("Inventor.Application");
                        invApp = (Inventor.Application)System.Activator.CreateInstance(invAppType);
                        invApp.Visible = true;
                    }
                }
                return invApp;
            }
            set { invApp = value; }
        }

        public Inventor.ApprenticeServerComponent ActiveApprenticeServer
        {
            get
            {
                if (apprenticeServer == null)
                {
                    apprenticeServer = new ApprenticeServerComponentClass();
                }
                return apprenticeServer;
            }

            set { value = apprenticeServer; }
        }

        public void ResetOnDocumentActivate(_Document activeDoc)
        {
            try
            {
                if (activeDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    PersistenceManager.ActiveAssemblyDoc = (AssemblyDocument)activeDoc;
                    ReferenceManager.KeyManager = PersistenceManager.ActiveAssemblyDoc.ReferenceKeyManager;
                }

                else if (activeDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    PersistenceManager.ActiveDrawingDoc = (DrawingDocument)activeDoc;
                    ReferenceManager.KeyManager = PersistenceManager.ActiveDrawingDoc.ReferenceKeyManager;
                }

                else if (activeDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    PersistenceManager.ActivePartDoc = (PartDocument)activeDoc;
                    ReferenceManager.KeyManager = PersistenceManager.ActivePartDoc.ReferenceKeyManager;
                }

                else
                {
                    ResetOnDocumentDeactivate();
                }
            }
            catch (Exception e)
            {
                throw;
            }
        }

        public void ResetOnDocumentDeactivate()
        {
            ActiveAssemblyDoc = null;
            ActiveDrawingDoc = null;
            ActivePartDoc = null;
        }

        void appEvents_OnActivateDocument(_Document documentObject, EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
        {
            handlingCode = HandlingCodeEnum.kEventNotHandled;
            if (beforeOrAfter == EventTimingEnum.kAfter)
            {
                ResetOnDocumentActivate(documentObject);
            }     
        }

        void appEvents_OnDeactivateDocument(_Document documentObject, EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
        {
            handlingCode = HandlingCodeEnum.kEventNotHandled;
            if (beforeOrAfter == EventTimingEnum.kBefore)
            {
                ResetOnDocumentDeactivate();
            }
        }
    }
}
