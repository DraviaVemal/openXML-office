// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using LiteDB;

namespace OpenXMLOffice.Excel_2013
{
	/// <summary>
	/// Represents a class that manages shared string values in Excel.
	/// </summary>
	internal class ShareString : IDisposable
	{

		private static readonly LiteDatabase liteDatabase = new(Path.ChangeExtension(Path.GetTempFileName(), "db"));
		private static ShareString? instance = null;
		private readonly ILiteCollection<Record> collection;

		/// <summary>
		/// Initializes a new instance of the <see cref="ShareString"/> class.
		/// </summary>
		private ShareString()
		{
			collection = liteDatabase.GetCollection<Record>("Records");
			collection.EnsureIndex("Record.Value");
		}

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

		/// <summary>
		/// Releases all resources used by the <see cref="ShareString"/> class.
		/// </summary>
		public void Dispose()
		{
			liteDatabase.Dispose();
		}

		/// <summary>
		/// Gets the index of the specified value in the shared string collection.
		/// </summary>
		/// <param name="value">
		/// The value to search for.
		/// </param>
		/// <returns>
		/// The index of the value if found; otherwise, null.
		/// </returns>
		public int? GetIndex(string value)
		{
			return collection.FindOne(col => col.Value == value)?.Id - 1;
		}

		/// <summary>
		/// Gets all the records in the shared string collection.
		/// </summary>
		/// <returns>
		/// A list of all the records.
		/// </returns>
		public List<string> GetRecords()
		{
			return collection.Query().OrderBy(x => x.Id).Select(x => x.Value).ToList();
		}

		/// <summary>
		/// Gets the value at the specified index in the shared string collection.
		/// </summary>
		/// <param name="index">
		/// The index of the value to retrieve.
		/// </param>
		/// <returns>
		/// The value at the specified index if found; otherwise, null.
		/// </returns>
		public string? GetValue(int index)
		{
			return collection.FindOne(col => col.Id == index)?.Value;
		}

		/// <summary>
		/// Inserts a new value into the shared string collection.
		/// </summary>
		/// <param name="Data">
		/// The value to insert.
		/// </param>
		public void Insert(string Data)
		{
			collection.Insert(new Record(Data));
		}

		/// <summary>
		/// Inserts multiple values into the shared string collection.
		/// </summary>
		/// <param name="data">
		/// The list of values to insert.
		/// </param>
		public void InsertBulk(List<string> data)
		{
			collection.InsertBulk(data.Select(item => new Record(item)));
		}

		/// <summary>
		/// Inserts a unique value into the shared string collection.
		/// </summary>
		/// <param name="data">
		/// The value to insert.
		/// </param>
		/// <returns>
		/// The index of the inserted value.
		/// </returns>
		public int InsertUnique(string data)
		{
			int? Index = GetIndex(data);
			if (Index != null)
			{
				return (int)Index;
			}
			BsonValue DocId = collection.Insert(new Record(data));
			return (int)DocId.AsInt64 - 1;
		}


	}
}
