/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/
using LiteDB;

namespace OpenXMLOffice.Excel
{
    public class Record
    {
        public int Id { get; set; }
        public string Value { get; set; }
        public Record(string Value)
        {
            this.Value = Value;
        }
    }
    internal class ShareString : IDisposable
    {
        private static ShareString? instance = null;

        // Public method to access the instance of the class
        public static ShareString Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ShareString();
                }
                return instance;
            }
        }
        static readonly LiteDatabase LiteDatabase = new(Path.ChangeExtension(Path.GetTempFileName(), "db"));
        readonly ILiteCollection<Record> Collection;
        private ShareString()
        {
            Collection = LiteDatabase.GetCollection<Record>("Records");
            Collection.EnsureIndex("Record.Value");
        }

        public void Dispose()
        {
            LiteDatabase.Dispose();
        }

        public void Insert(string Data)
        {
            Collection.Insert(new Record(Data));
        }

        public int InsertUnique(string Data)
        {
            int? Index = GetIndex(Data);
            if (Index != null)
            {
                return (int)Index;
            }
            BsonValue DocId = Collection.Insert(new Record(Data));
            return DocId.AsInt32 - 1;
        }

        public void InsertBulk(List<string> Data)
        {
            Collection.InsertBulk(Data.Select(item => new Record(item)));
        }

        public int? GetIndex(string Value)
        {
            return Collection.Query().Where(col => col.Value == Value).FirstOrDefault()?.Id - 1;
        }
        public string? GetValue(int Index)
        {
            return Collection.Query().Where(col => col.Id == Index).FirstOrDefault()?.Value;
        }

        public List<string> GetRecords()
        {
            return Collection.Query().OrderBy(x => x.Id).Select(x => x.Value).ToList();
        }
    }

}