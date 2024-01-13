/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using LiteDB;

namespace OpenXMLOffice.Excel
{
    /// <summary>
    /// Represents a class that manages shared string values in Excel.
    /// </summary>
    internal class ShareString : IDisposable
    {
        #region Private Fields

        private static readonly LiteDatabase LiteDatabase = new(Path.ChangeExtension(Path.GetTempFileName(), "db"));
        private static ShareString? instance = null;

        private readonly ILiteCollection<Record> Collection;

        #endregion Private Fields

        #region Private Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ShareString"/> class.
        /// </summary>
        private ShareString()
        {
            Collection = LiteDatabase.GetCollection<Record>("Records");
            Collection.EnsureIndex("Record.Value");
        }

        #endregion Private Constructors

        #region Public Properties

        /// <summary>
        /// Gets the instance of the <see cref="ShareString"/> class.
        /// </summary>
        public static ShareString Instance
        {
            get
            {
                instance ??= new ShareString();
                return instance;
            }
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        /// Releases all resources used by the <see cref="ShareString"/> class.
        /// </summary>
        public void Dispose()
        {
            LiteDatabase.Dispose();
        }

        /// <summary>
        /// Gets the index of the specified value in the shared string collection.
        /// </summary>
        /// <param name="Value">The value to search for.</param>
        /// <returns>The index of the value if found; otherwise, null.</returns>
        public int? GetIndex(string Value)
        {
            return Collection.Query().Where(col => col.Value == Value).FirstOrDefault()?.Id - 1;
        }

        /// <summary>
        /// Gets all the records in the shared string collection.
        /// </summary>
        /// <returns>A list of all the records.</returns>
        public List<string> GetRecords()
        {
            return Collection.Query().OrderBy(x => x.Id).Select(x => x.Value).ToList();
        }

        /// <summary>
        /// Gets the value at the specified index in the shared string collection.
        /// </summary>
        /// <param name="Index">The index of the value to retrieve.</param>
        /// <returns>The value at the specified index if found; otherwise, null.</returns>
        public string? GetValue(int Index)
        {
            return Collection.Query().Where(col => col.Id == Index).FirstOrDefault()?.Value;
        }

        /// <summary>
        /// Inserts a new value into the shared string collection.
        /// </summary>
        /// <param name="Data">The value to insert.</param>
        public void Insert(string Data)
        {
            Collection.Insert(new Record(Data));
        }

        /// <summary>
        /// Inserts multiple values into the shared string collection.
        /// </summary>
        /// <param name="Data">The list of values to insert.</param>
        public void InsertBulk(List<string> Data)
        {
            Collection.InsertBulk(Data.Select(item => new Record(item)));
        }

        /// <summary>
        /// Inserts a unique value into the shared string collection.
        /// </summary>
        /// <param name="Data">The value to insert.</param>
        /// <returns>The index of the inserted value.</returns>
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