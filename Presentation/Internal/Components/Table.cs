/*
* Copyright (c) DraviaVemal. All Rights Reserved. Licensed under the MIT License.
* See License in the project root for license information.
*/

using System.Data;
using A = DocumentFormat.OpenXml.Drawing;
using G = OpenXMLOffice.Global;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLOffice.Presentation
{
    /// <summary>
    /// Represents Table Class
    /// </summary>
    public class Table : G.CommonProperties
    {
        #region Private Fields

        private readonly P.GraphicFrame GraphicFrame = new();
        private readonly TableSetting TableSetting;

        #endregion Private Fields

        #region Public Constructors
        /// <summary>
        /// Create Table with provided settings
        /// </summary>
        /// <param name="TableRows"></param>
        /// <param name="TableSetting"></param>
        public Table(TableRow[] TableRows, TableSetting TableSetting)
        {
            this.TableSetting = TableSetting;
            CreateTableGraphicFrame(TableRows);
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// </summary>
        /// <returns>
        /// X,Y
        /// </returns>
        public (uint, uint) GetPosition()
        {
            return (TableSetting.X, TableSetting.Y);
        }

        /// <summary>
        /// </summary>
        /// <returns>
        /// Width,Height
        /// </returns>
        public (uint, uint) GetSize()
        {
            return (TableSetting.Width, TableSetting.Height);
        }
        /// <summary>
        /// Get Table Graphic Frame
        /// </summary>
        /// <returns></returns>
        public P.GraphicFrame GetTableGraphicFrame()
        {
            return GraphicFrame;
        }
        /// <summary>
        /// Update Table Position
        /// </summary>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        public void UpdatePosition(uint X, uint Y)
        {
            TableSetting.X = X;
            TableSetting.Y = Y;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = TableSetting.X, Y = TableSetting.Y },
                    Extents = new A.Extents { Cx = TableSetting.Width, Cy = TableSetting.Height }
                };
            }
        }
        /// <summary>
        /// Update Table Size
        /// </summary>
        /// <param name="Width"></param>
        /// <param name="Height"></param>
        public void UpdateSize(uint Width, uint Height)
        {
            ReCalculateColumnWidth();
            TableSetting.Width = Width;
            TableSetting.Height = Height;
            if (GraphicFrame != null)
            {
                GraphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = TableSetting.X, Y = TableSetting.Y },
                    Extents = new A.Extents { Cx = TableSetting.Width, Cy = TableSetting.Height }
                };
            }
        }

        #endregion Public Methods

        #region Private Methods

        private long CalculateColumnWidth(TableSetting.WidthOptionValues widthType, float InputWidth)
        {
            return widthType switch
            {
                TableSetting.WidthOptionValues.PIXEL => G.ConverterUtils.PixelsToEmu(Convert.ToInt32(InputWidth)),
                TableSetting.WidthOptionValues.PERCENTAGE => Convert.ToInt32(TableSetting.Width / 100 * InputWidth),
                TableSetting.WidthOptionValues.RATIO => Convert.ToInt32(TableSetting.Width / 100 * (InputWidth * 10)),
                _ => Convert.ToInt32(InputWidth)
            };
        }

        private A.Table CreateTable(TableRow[] TableRows)
        {
            if (TableRows.Length < 1 || TableRows[0].TableCells.Count < 1)
            {
                throw new DataException("No Table Data Provided");
            }
            if (TableSetting.WidthType != TableSetting.WidthOptionValues.AUTO && TableSetting.TableColumnWidth.Count != TableRows[0].TableCells.Count)
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
                Paragraph.Append(new TextBox(new G.TextBoxSetting()
                {
                    Text = Cell.Value,
                    TextBackground = Cell.TextBackground,
                    TextColor = Cell.TextColor,
                    FontFamily = Cell.FontFamily,
                    FontSize = Cell.FontSize,
                    IsBold = Cell.IsBold,
                    IsItalic = Cell.IsItalic,
                    IsUnderline = Cell.IsUnderline,
                }).GetTextBoxRun());
            }
            TableCell.Append(new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                Paragraph
            ));
            A.TableCellProperties TableCellProperties = new();
            TableCellProperties.Append(new A.LeftBorderLineProperties(
                Cell.LeftBorder ? CreateSolidFill(new List<string>() { "000000" }, 0) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.RightBorderLineProperties(
                Cell.RightBorder ? CreateSolidFill(new List<string>() { "000000" }, 0) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.TopBorderLineProperties(
                Cell.TopBorder ? CreateSolidFill(new List<string>() { "000000" }, 0) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.BottomBorderLineProperties(
                Cell.BottomBorder ? CreateSolidFill(new List<string>() { "000000" }, 0) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.TopLeftToBottomRightBorderLineProperties(
                Cell.BottomBorder ? CreateSolidFill(new List<string>() { "000000" }, 0) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.BottomLeftToTopRightBorderLineProperties(
                Cell.BottomBorder ? CreateSolidFill(new List<string>() { "000000" }, 0) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(Cell.CellBackground != null ? CreateSolidFill(new List<string>() { Cell.CellBackground }, 0) : new A.NoFill());
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
                    X = TableSetting.X,
                    Y = TableSetting.Y
                },
                Extents = new A.Extents()
                {
                    Cx = TableSetting.Width,
                    Cy = TableSetting.Height
                }
            };
        }

        private A.TableGrid CreateTableGrid(int ColumnCount)
        {
            A.TableGrid TableGrid = new();
            if (TableSetting.WidthType == TableSetting.WidthOptionValues.AUTO)
            {
                for (int i = 0; i < ColumnCount; i++)
                {
                    TableGrid.Append(new A.GridColumn() { Width = TableSetting.Width / ColumnCount });
                }
            }
            else
            {
                for (int i = 0; i < ColumnCount; i++)
                {
                    TableGrid.Append(new A.GridColumn() { Width = CalculateColumnWidth(TableSetting.WidthType, TableSetting.TableColumnWidth[i]) });
                }
            }
            return TableGrid;
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

        private void ReCalculateColumnWidth()
        {
            A.Table? Table = GraphicFrame!.Graphic!.GraphicData!.GetFirstChild<A.Table>();
            if (Table != null)
            {
                List<A.GridColumn> GridColumn = Table.TableGrid!.Elements<A.GridColumn>().ToList();
                if (TableSetting.WidthType == TableSetting.WidthOptionValues.AUTO)
                {
                    GridColumn.ForEach(Column => Column.Width = TableSetting.Width / GridColumn.Count);
                }
                else
                {
                    GridColumn.Select((item, index) => (item, index)).ToList().ForEach(Column =>
                        Column.item.Width = CalculateColumnWidth(TableSetting.WidthType, TableSetting.TableColumnWidth[Column.index]));
                }
            }
        }

        #endregion Private Methods
    }
}