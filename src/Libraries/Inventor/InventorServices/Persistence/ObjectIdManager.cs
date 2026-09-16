using System;
using DynamoServices;

namespace InventorServices.Persistence
{
    /// <summary>
    /// Stores a single Inventor reference key in Dynamo's per-node trace slot.
    /// Dynamo 3+/4 trace data is a string, so the key bytes travel as base64.
    /// </summary>
    public class ObjectIdManager : ISerializableIdManager
    {
        public bool GetTraceData(string key, out ISerializableId<byte[]> id)
        {
            id = null;
            var stored = TraceUtils.GetTraceData(key);
            if (string.IsNullOrEmpty(stored))
            {
                return false;
            }

            try
            {
                id = new ObjectId { Id = Convert.FromBase64String(stored) };
                return true;
            }
            catch (FormatException)
            {
                // Trace written by an older (BinaryFormatter-era) build, or by another library: ignore it.
                return false;
            }
        }

        public void SetTraceData(string key, System.Runtime.Serialization.ISerializable value)
        {
            if (value is not ISerializableId<byte[]> id)
            {
                throw new ArgumentException("Expected an ISerializableId<byte[]> (ObjectId).", nameof(value));
            }

            TraceUtils.SetTraceData(key, Convert.ToBase64String(id.Id ?? Array.Empty<byte>()));
        }
    }
}
