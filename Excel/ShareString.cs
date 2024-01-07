/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using LiteDB;

namespace OpenXMLOffice.Excel
{
    public class Record
    {
        #region Internal Constructors

        internal Record(string Value)
        {
            this.Value = Value;
        }

        #endregion Internal Constructors

        #region Public Properties

        public int Id { get; set; }
        public string Value { get; set; }

        #endregion Public Properties
    }

    internal class ShareString : IDisposable
    {
        #region Private Fields

        private static readonly LiteDatabase LiteDatabase = new(Path.ChangeExtension(Path.GetTempFileName(), "db"));
        private static ShareString? instance = null;

        private readonly ILiteCollection<Record> Collection;

        #endregion Private Fields

        #region Private Constructors

        private ShareString()
        {
            Collection = LiteDatabase.GetCollection<Record>("Records");
            Collection.EnsureIndex("Record.Value");
        }

        #endregion Private Constructors

        #region Public Properties

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

        #endregion Public Properties

        #region Public Methods

        public void Dispose()
        {
            LiteDatabase.Dispose();
        }

        public int? GetIndex(string Value)
        {
            return Collection.Query().Where(col => col.Value == Value).FirstOrDefault()?.Id - 1;
        }

        public List<string> GetRecords()
        {
            return Collection.Query().OrderBy(x => x.Id).Select(x => x.Value).ToList();
        }

        public string? GetValue(int Index)
        {
            return Collection.Query().Where(col => col.Id == Index).FirstOrDefault()?.Value;
        }

        public void Insert(string Data)
        {
            Collection.Insert(new Record(Data));
        }

        public void InsertBulk(List<string> Data)
        {
            Collection.InsertBulk(Data.Select(item => new Record(item)));
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

        #endregion Public Methods
    }
}