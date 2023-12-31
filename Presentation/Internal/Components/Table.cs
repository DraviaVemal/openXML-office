using System.Data;
using OpenXMLOffice.Global;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation

{
    public class Table
    {
        private int X = 0;
        private int Y = 0;
        private int Width = 8128000;
        private int Height = 741680;
        private readonly P.GraphicFrame GraphicFrame = new();
        private readonly TableSetting TableSetting;

        public Table(TableRow[] TableRows, TableSetting TableSetting)
        {
            this.TableSetting = TableSetting;
            CreateTableGraphicFrame(TableRows);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns>X,Y</returns>
        public (int, int) GetPosition()
        {
            return (X, Y);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns>Width,Height</returns>
        public (int, int) GetSize()
        {
            return (Width, Height);
        }

        public void UpdatePosition(int X, int Y)
        {
            this.X = X;
            this.Y = Y;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = X, Y = Y },
                    Extents = new A.Extents { Cx = Width, Cy = Height }
                };
            }
        }

        public void UpdateSize(int Width, int Height)
        {
            if (this.Width != Width && TableSetting != null)
            {
                ReCalculateColumnWidth();
            }
            this.Width = Width;
            this.Height = Height;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = X, Y = Y },
                    Extents = new A.Extents { Cx = Width, Cy = Height }
                };
            }
        }

        public P.GraphicFrame GetTableGraphicFrame()
        {
            return GraphicFrame;
        }

        private A.Table CreateTable(TableRow[] TableRows)
        {
            if (TableRows.Length < 1 || TableRows[0].TableCells.Count < 1)
            {
                throw new DataException("No Table Data Provided");
            }
            if (TableSetting.WidthType != TableSetting.eWidthType.AUTO && TableSetting.TableColumnwidth.Count != TableRows[0].TableCells.Count)
            {
                throw new ArgumentException("Column With Setting Does Not Match Data");
            }
            A.Table Table = new()
            {
                TableProperties = new A.TableProperties()
                {
                    FirstRow = true,
                    BandRow = true
                },
                TableGrid = CreateTableGrid(TableRows[0].TableCells.Count)
            };
            // Add Table Data Row
            foreach (TableRow row in TableRows)
            {
                Table.Append(CreateTableRow(row));
            }
            return Table;
        }

        private A.TableRow CreateTableRow(TableRow Row)
        {
            A.TableRow TableRow = new()
            {
                Height = Row.Height
            };
            foreach (TableCell cell in Row.TableCells)
            {
                TableRow.Append(CreateTableCell(cell));
            }
            return TableRow;
        }

        private A.TableGrid CreateTableGrid(int ColumnCount)
        {
            A.TableGrid TableGrid = new();
            if (TableSetting.WidthType == TableSetting.eWidthType.AUTO)
            {
                for (int i = 0; i < ColumnCount; i++)
                {
                    TableGrid.Append(new A.GridColumn() { Width = Width / ColumnCount });
                }
            }
            else
            {
                for (int i = 0; i < ColumnCount; i++)
                {
                    TableGrid.Append(new A.GridColumn() { Width = CalculateColumnWidth(TableSetting.WidthType, TableSetting.TableColumnwidth[i]) });
                }
            }
            return TableGrid;
        }

        private void ReCalculateColumnWidth()
        {
            A.Table? Table = GraphicFrame!.Graphic!.GraphicData!.GetFirstChild<A.Table>();
            if (Table != null)
            {
                List<A.GridColumn> GridColumn = Table.TableGrid!.Elements<A.GridColumn>().ToList();
                if (TableSetting.WidthType == TableSetting.eWidthType.AUTO)
                {
                    GridColumn.ForEach(Column => Column.Width = Width / GridColumn.Count);
                }
                else
                {
                    GridColumn.Select((item, index) => (item, index)).ToList().ForEach(Column =>
                        Column.item.Width = CalculateColumnWidth(TableSetting.WidthType, TableSetting.TableColumnwidth[Column.index]));
                }
            }
        }

        private long CalculateColumnWidth(TableSetting.eWidthType widthType, float InputWidth)
        {
            return widthType switch
            {
                TableSetting.eWidthType.PIXEL => ConverterUtils.PixelsToEmu(Convert.ToInt32(InputWidth)),
                TableSetting.eWidthType.PERCENTAGE => Convert.ToInt32(Width / 100 * InputWidth),
                TableSetting.eWidthType.RATIO => Convert.ToInt32(Width / 100 * (InputWidth * 10)),
                _ => Convert.ToInt32(InputWidth)
            };
        }

        private A.TableCell CreateTableCell(TableCell Cell)
        {
            A.TableCell TableCell = new();
            A.Paragraph Paragraph = new();
            if (Cell.Value == null)
            {
                Paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-IN" });
            }
            else
            {
                Paragraph.Append(new TextBox().CreateTextRun(new TextBoxSetting()
                {
                    Text = Cell.Value,
                    TextBackground = Cell.TextBackground,
                    TextColor = Cell.TextColor,
                    FontFamily = Cell.FontFamily,
                    FontSize = Cell.FontSize,
                    IsBold = Cell.IsBold,
                    IsItalic = Cell.IsItalic,
                    IsUnderline = Cell.IsUnderline,
                }));
            }
            TableCell.Append(new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                Paragraph
            ));
            A.TableCellProperties TableCellProperties = new();
            TableCellProperties.Append(new A.LeftBorderLineProperties(
                Cell.LeftBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.RightBorderLineProperties(
                Cell.RightBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.TopBorderLineProperties(
                Cell.TopBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.BottomBorderLineProperties(
                Cell.BottomBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.TopLeftToBottomRightBorderLineProperties(
                Cell.BottomBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.BottomLeftToTopRightBorderLineProperties(
                Cell.BottomBorder ? new A.SolidFill(new A.RgbColorModelHex() { Val = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex() { Val = Cell.CellBackground }));
            TableCell.Append(TableCellProperties);
            return TableCell;
        }

        private void CreateTableGraphicFrame(TableRow[] TableRows)
        {
            A.GraphicData GraphicData = new(CreateTable(TableRows))
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            GraphicFrame.NonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties(
               new P.NonVisualDrawingProperties()
               {
                   Id = 1,
                   Name = "Table 1"
               },
               new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoGrouping = true }),
               new P.ApplicationNonVisualDrawingProperties());
            GraphicFrame.Graphic = new A.Graphic()
            {
                GraphicData = GraphicData
            };
            GraphicFrame.Transform = new P.Transform()
            {
                Offset = new A.Offset()
                {
                    X = X,
                    Y = Y
                },
                Extents = new A.Extents()
                {
                    Cx = Width,
                    Cy = Height
                }
            };
        }
    }
}