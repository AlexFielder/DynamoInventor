using System;
using System.Collections.Generic;
using Autodesk.DesignScript.Runtime;

namespace InventorLibrary.API
{
    [IsVisibleInDynamoLibrary(false)]
    public class InvAssemblyComponentDefinitions : IEnumerable<InvAssemblyComponentDefinition>
    {
        #region Internal properties
        List<InvAssemblyComponentDefinition> compDefList;
        
        internal Inventor.AssemblyComponentDefinitions InternalAssemblyComponentDefinitions { get; set; }

        internal Object InternalApplication
        {
            get { return AssemblyComponentDefinitionsInstance.Application; }
        }

        internal int InternalCount
        {
            get { return compDefList.Count; }
        }

        //internal Inv_AssemblyDocument InternalParent
        //{
        //    get { return Inv_AssemblyDocument.ByInv_AssemblyDocument(AssemblyComponentDefinitionsInstance.Parent); }
        //}


        internal InvObjectTypeEnum InternalType
        {
            get { return AssemblyComponentDefinitionsInstance.Type.As<InvObjectTypeEnum>(); }
        }


        #endregion

        #region Private constructors
        private InvAssemblyComponentDefinitions(InvAssemblyComponentDefinitions invAssemblyComponentDefinitions)
        {
            InternalAssemblyComponentDefinitions = invAssemblyComponentDefinitions.InternalAssemblyComponentDefinitions;
            compDefList = new List<InvAssemblyComponentDefinition>();
            foreach (var compDef in InternalAssemblyComponentDefinitions)
            {
                compDefList.Add(InvAssemblyComponentDefinition.ByInvAssemblyComponentDefinition((Inventor.AssemblyComponentDefinition)compDef));
            }
        }

        private InvAssemblyComponentDefinitions(Inventor.AssemblyComponentDefinitions invAssemblyComponentDefinitions)
        {
            InternalAssemblyComponentDefinitions = invAssemblyComponentDefinitions;
            compDefList = new List<InvAssemblyComponentDefinition>();
            foreach (var compDef in InternalAssemblyComponentDefinitions)
            {
                compDefList.Add(InvAssemblyComponentDefinition.ByInvAssemblyComponentDefinition((Inventor.AssemblyComponentDefinition)compDef));
            }
        }
        #endregion

        #region Private methods
        #endregion

        #region Public properties
        public Inventor.AssemblyComponentDefinitions AssemblyComponentDefinitionsInstance
        {
            get { return InternalAssemblyComponentDefinitions; }
            set { InternalAssemblyComponentDefinitions = value; }
        }

        public InvAssemblyComponentDefinition this[int index]
        {
            get { return compDefList[index]; }
            set { compDefList.Insert(index, value); }
        }

        public Object Application
        {
            get { return InternalApplication; }
        }

        public int Count
        {
            get { return InternalCount; }
        }

        //public Inv_AssemblyDocument Parent
        //{
        //    get { return InternalParent; }
        //}

        public InvObjectTypeEnum Type
        {
            get { return InternalType; }
        }

        #endregion

        #region Public static constructors
        public static InvAssemblyComponentDefinitions ByInvAssemblyComponentDefinitions(InvAssemblyComponentDefinitions invAssemblyComponentDefinitions)
        {
            return new InvAssemblyComponentDefinitions(invAssemblyComponentDefinitions);
        }
        public static InvAssemblyComponentDefinitions ByInvAssemblyComponentDefinitions(Inventor.AssemblyComponentDefinitions invAssemblyComponentDefinitions)
        {
            return new InvAssemblyComponentDefinitions(invAssemblyComponentDefinitions);
        }
        #endregion

        #region Public methods
        #endregion

        public void Add(InvAssemblyComponentDefinition invAssemblyCompDef)
        {
            compDefList.Add(invAssemblyCompDef);
        }

        public IEnumerator<InvAssemblyComponentDefinition> GetEnumerator()
        {
            return compDefList.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
