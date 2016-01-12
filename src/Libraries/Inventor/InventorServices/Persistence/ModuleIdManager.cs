﻿using DynamoServices;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace InventorServices.Persistence
{
    public class ModuleIdManager : ISerializableModuleIdManager
    {
        public bool GetTraceData(string key, out ISerializableId<List<Tuple<string, int, int, byte[]>>> id)
        {
            id = TraceUtils.GetTraceData(key) as ISerializableId<List<Tuple<string, int, int, byte[]>>>;
            if (id != null)
            {
                return true;    
            }
            else
            {
                return false;
            }
        }

        public void SetTraceData(string key, ISerializable value)
        {
            TraceUtils.SetTraceData(key, value);
        }
    }
}
