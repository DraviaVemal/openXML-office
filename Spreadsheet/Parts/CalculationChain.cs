// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LiteDB;

namespace OpenXMLOffice.Spreadsheet_2007
{
	/// <summary>
	/// 
	/// </summary>
	internal class CalculationChainService : IDisposable
	{
		private readonly LiteDatabase liteDatabase = new LiteDatabase(Path.ChangeExtension(Path.GetTempFileName(), "db"));
		private readonly ILiteCollection<CalculationRecord> calculationCollection;
		/// <summary>
		/// 
		/// </summary>
		internal CalculationChainService()
		{
			calculationCollection = liteDatabase.GetCollection<CalculationRecord>("CalculationRecord");
			calculationCollection.EnsureIndex("CalculationRecord.Value");
		}
		/// <summary>
		/// 
		/// </summary>
		public void Dispose()
		{
			liteDatabase.Dispose();
		}
		/// <summary>
		/// 
		/// </summary>
		public void InsertBulk(List<CalculationRecord> data)
		{
			calculationCollection.InsertBulk(data);
		}
		/// <summary>
		/// 
		/// </summary>
		public List<CalculationRecord> GetRecords()
		{
			return calculationCollection.Query().OrderBy(x => x.Id).ToList();
		}

		/// <summary>
		/// Add Calculation chain data
		/// </summary>
		/// <param name="CellId">Cell the formula applied</param>
		/// <param name="SheetIndex">Sheet Id of cell location</param>
		public void AddRecord(string CellId, uint SheetIndex)
		{
			CalculationRecord record = calculationCollection.FindOne(item => item.CellId == CellId && item.SheetIndex == SheetIndex);
			// Check if record already exist and insert
			if (record == null)
			{
				calculationCollection.Insert(new CalculationRecord(CellId, (int)SheetIndex));
			}
		}

		/// <summary>
		/// Remove Calculation chain
		/// </summary>
		/// <param name="CellId">Cell the formula applied</param>
		/// <param name="SheetIndex">Sheet Id of cell location</param>
		public void RemoveRecord(string CellId, uint SheetIndex)
		{
			// Remove if record exist
			CalculationRecord record = calculationCollection.FindOne(item => item.CellId == CellId && item.SheetIndex == SheetIndex);
			if (record != null)
			{
				calculationCollection.Delete(record.Id);
			}
		}
	}
}
