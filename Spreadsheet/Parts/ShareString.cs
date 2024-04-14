// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LiteDB;
namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// Represents a class that manages shared string values in Excel.
	/// </summary>
	internal class ShareStringService : IDisposable
	{
		private static readonly LiteDatabase liteDatabase = new LiteDatabase(Path.ChangeExtension(Path.GetTempFileName(), "db"));
		private readonly ILiteCollection<StringRecord> stringCollection;
		/// <summary>
		/// Initializes a new instance of the <see cref="ShareStringService"/> class.
		/// </summary>
		internal ShareStringService()
		{
			stringCollection = liteDatabase.GetCollection<StringRecord>("StringRecord");
			stringCollection.EnsureIndex("StringRecord.Value");
		}
		/// <summary>
		/// Releases all resources used by the <see cref="ShareStringService"/> class.
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
			return stringCollection.FindOne(col => col.Value == value).Id - 1;
		}
		/// <summary>
		/// Gets all the records in the shared string collection.
		/// </summary>
		/// <returns>
		/// A list of all the records.
		/// </returns>
		public List<string> GetRecords()
		{
			return stringCollection.Query().OrderBy(x => x.Id).Select(x => x.Value).ToList();
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
		public string GetValue(int index)
		{
			return stringCollection.FindById(index).Value;
		}
		/// <summary>
		/// Inserts a new value into the shared string collection.
		/// </summary>
		/// <param name="Data">
		/// The value to insert.
		/// </param>
		public void Insert(string Data)
		{
			stringCollection.Insert(new StringRecord(Data));
		}
		/// <summary>
		/// Inserts multiple values into the shared string collection.
		/// </summary>
		/// <param name="data">
		/// The list of values to insert.
		/// </param>
		public void InsertBulk(List<string> data)
		{
			stringCollection.InsertBulk(data.Select(item => new StringRecord(item)));
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
			BsonValue DocId = stringCollection.Insert(new StringRecord(data));
			return (int)DocId.AsInt64 - 1;
		}
	}
}
