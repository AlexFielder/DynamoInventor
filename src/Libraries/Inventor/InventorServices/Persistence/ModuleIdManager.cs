using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.Json;
using DynamoServices;

namespace InventorServices.Persistence
{
    /// <summary>
    /// Stores a module's set of Inventor reference keys in Dynamo's per-node trace slot.
    /// Dynamo 3+/4 trace data is a string, so the list is JSON with base64 key bytes.
    /// </summary>
    public class ModuleIdManager : ISerializableModuleIdManager
    {
        private sealed class Entry
        {
            public string Name { get; set; }
            public int Row { get; set; }
            public int Column { get; set; }
            public string Key { get; set; }
        }

        public bool GetTraceData(string key, out ISerializableId<List<Tuple<string, int, int, byte[]>>> id)
        {
            id = null;
            var stored = TraceUtils.GetTraceData(key);
            if (string.IsNullOrEmpty(stored))
            {
                return false;
            }

            try
            {
                var entries = JsonSerializer.Deserialize<List<Entry>>(stored) ?? new List<Entry>();
                id = new ModuleId
                {
                    Id = entries.Select(e => Tuple.Create(e.Name, e.Row, e.Column,
                            string.IsNullOrEmpty(e.Key) ? Array.Empty<byte>() : Convert.FromBase64String(e.Key)))
                        .ToList()
                };
                return true;
            }
            catch (Exception ex) when (ex is JsonException || ex is FormatException)
            {
                // Trace written by an older (BinaryFormatter-era) build, or by another library: ignore it.
                return false;
            }
        }

        public void SetTraceData(string key, ISerializable value)
        {
            if (value is not ISerializableId<List<Tuple<string, int, int, byte[]>>> id)
            {
                throw new ArgumentException("Expected an ISerializableId<List<Tuple<string,int,int,byte[]>>> (ModuleId).", nameof(value));
            }

            var entries = (id.Id ?? new List<Tuple<string, int, int, byte[]>>())
                .Select(t => new Entry
                {
                    Name = t.Item1,
                    Row = t.Item2,
                    Column = t.Item3,
                    Key = t.Item4 == null ? null : Convert.ToBase64String(t.Item4)
                })
                .ToList();

            TraceUtils.SetTraceData(key, JsonSerializer.Serialize(entries));
        }
    }
}
