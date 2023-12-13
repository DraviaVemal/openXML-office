using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OpenXMLOffice.Global
{
    public class BarFamilyChart : ChartBase
    {
        #region Protected Methods

        protected NumberingCache AddNumberCacheValue(List<ChartData> Cells, ChartSeriesSetting? ChartSeriesSetting)
        {
            try
            {
                NumberingCache NumberingCache = new()
                {
                    FormatCode = new FormatCode(ChartSeriesSetting?.NumberFormat ?? "General"),
                    PointCount = new PointCount()
                    {
                        Val = (UInt32Value)(uint)Cells.Count
                    },
                };
                int count = 0;
                foreach (ChartData Cell in Cells)
                {
                    StringPoint StringPoint = new()
                    {
                        Index = (UInt32Value)(uint)count
                    };
                    StringPoint.AppendChild(new NumericValue(Cell.Value ?? ""));
                    NumberingCache.AppendChild(StringPoint);
                    ++count;
                }
                return NumberingCache;
            }
            catch
            {
                throw new Exception("Chart. Numeric Ref Error");
            }
        }

        protected StringCache AddStringCacheValue(List<ChartData> Cells)
        {
            try
            {
                StringCache StringCache = new()
                {
                    PointCount = new PointCount()
                    {
                        Val = (UInt32Value)(uint)Cells.Count
                    },
                };
                int count = 0;
                foreach (ChartData Cell in Cells)
                {
                    StringPoint StringPoint = new()
                    {
                        Index = (UInt32Value)(uint)count
                    };
                    StringPoint.AppendChild(new NumericValue(Cell.Value ?? ""));
                    StringCache.AppendChild(StringPoint);
                    ++count;
                }
                return StringCache;
            }
            catch
            {
                throw new Exception("Chart. String Ref Error");
            }
        }

        #endregion Protected Methods
    }
}