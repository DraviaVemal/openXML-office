// Copyright (c) DraviaVemal. Licensed under the MIT License. See License in the project root.

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

        private readonly P.GraphicFrame graphicFrame = new();
        private readonly TableSetting tableSetting;

        #endregion Private Fields

        #region Public Constructors

        /// <summary>
        /// Create Table with provided settings
        /// </summary>
        /// <param name="TableRows">
        /// </param>
        /// <param name="TableSetting">
        /// </param>
        public Table(TableRow[] TableRows, TableSetting TableSetting)
        {
            tableSetting = TableSetting;
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
            return (tableSetting.x, tableSetting.y);
        }

        /// <summary>
        /// </summary>
        /// <returns>
        /// Width,Height
        /// </returns>
        public (uint, uint) GetSize()
        {
            return (tableSetting.width, tableSetting.height);
        }

        /// <summary>
        /// Get Table Graphic Frame
        /// </summary>
        /// <returns>
        /// </returns>
        public P.GraphicFrame GetTableGraphicFrame()
        {
            return graphicFrame;
        }

        /// <summary>
        /// Update Table Position
        /// </summary>
        /// <param name="X">
        /// </param>
        /// <param name="Y">
        /// </param>
        public void UpdatePosition(uint X, uint Y)
        {
            tableSetting.x = X;
            tableSetting.y = Y;
            if (graphicFrame != null)
            {
                graphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = tableSetting.x, Y = tableSetting.y },
                    Extents = new A.Extents { Cx = tableSetting.width, Cy = tableSetting.height }
                };
            }
        }

        /// <summary>
        /// Update Table Size
        /// </summary>
        /// <param name="Width">
        /// </param>
        /// <param name="Height">
        /// </param>
        public void UpdateSize(uint Width, uint Height)
        {
            ReCalculateColumnWidth();
            tableSetting.width = Width;
            tableSetting.height = Height;
            if (graphicFrame != null)
            {
                graphicFrame.Transform = new P.Transform
                {
                    Offset = new A.Offset { X = tableSetting.x, Y = tableSetting.y },
                    Extents = new A.Extents { Cx = tableSetting.width, Cy = tableSetting.height }
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
                TableSetting.WidthOptionValues.PERCENTAGE => Convert.ToInt32(tableSetting.width / 100 * InputWidth),
                TableSetting.WidthOptionValues.RATIO => Convert.ToInt32(tableSetting.width / 100 * (InputWidth * 10)),
                _ => Convert.ToInt32(InputWidth)
            };
        }

        private A.Table CreateTable(TableRow[] TableRows)
        {
            if (TableRows.Length < 1 || TableRows[0].tableCells.Count < 1)
            {
                throw new DataException("No Table Data Provided");
            }
            if (tableSetting.widthType != TableSetting.WidthOptionValues.AUTO && tableSetting.tableColumnWidth.Count != TableRows[0].tableCells.Count)
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
                TableGrid = CreateTableGrid(TableRows[0].tableCells.Count)
            };
            // Add Table Data Row
            foreach (TableRow row in TableRows)
            {
                Table.Append(CreateTableRow(row));
            }
            return Table;
        }

        private A.TableCell CreateTableCell(TableCell Cell, TableRow Row)
        {
            A.Paragraph Paragraph = new();
            if (Cell.alignment != null)
            {
                Paragraph.Append(new A.ParagraphProperties()
                {
                    Alignment = Cell.alignment switch
                    {
                        TableCell.AlignmentValues.CENTER => A.TextAlignmentTypeValues.Center,
                        TableCell.AlignmentValues.LEFT => A.TextAlignmentTypeValues.Left,
                        TableCell.AlignmentValues.JUSTIFY => A.TextAlignmentTypeValues.Justified,
                        _ => A.TextAlignmentTypeValues.Left
                    }
                });
            }
            if (Cell.value == null)
            {
                Paragraph.Append(new A.EndParagraphRunProperties() { Language = "en-IN" });
            }
            else
            {
                Paragraph.Append(new TextBox(new G.TextBoxSetting()
                {
                    text = Cell.value,
                    textBackground = Cell.textBackground,
                    textColor = Cell.textColor,
                    fontFamily = Cell.fontFamily,
                    fontSize = Cell.fontSize,
                    isBold = Cell.isBold,
                    isItalic = Cell.isItalic,
                    isUnderline = Cell.isUnderline,
                }).GetTextBoxRun());
            }
            A.TableCell TableCellXML = new();
            TableCellXML.Append(new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                Paragraph
            ));
            A.TableCellProperties TableCellProperties = new();
            TableCellProperties.Append(new A.LeftBorderLineProperties(
                Cell.leftBorder ? G.CommonProperties.CreateSolidFill(new() { hexColor = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.RightBorderLineProperties(
                Cell.rightBorder ? G.CommonProperties.CreateSolidFill(new() { hexColor = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.TopBorderLineProperties(
                Cell.topBorder ? G.CommonProperties.CreateSolidFill(new() { hexColor = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.BottomBorderLineProperties(
                Cell.bottomBorder ? G.CommonProperties.CreateSolidFill(new() { hexColor = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.TopLeftToBottomRightBorderLineProperties(
                Cell.topLeftToBottomRightBorder ? G.CommonProperties.CreateSolidFill(new() { hexColor = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append(new A.BottomLeftToTopRightBorderLineProperties(
                Cell.bottomLeftToTopRightBorder ? G.CommonProperties.CreateSolidFill(new() { hexColor = "000000" }) : new A.NoFill(),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid }
            )
            { Width = 12700, CompoundLineType = A.CompoundLineValues.Single });
            TableCellProperties.Append((Cell.cellBackground != null || Row.rowBackground != null) ? G.CommonProperties.CreateSolidFill(new() { hexColor = Cell.cellBackground ?? Row.rowBackground! }) : new A.NoFill());
            TableCellXML.Append(TableCellProperties);
            return TableCellXML;
        }

        private void CreateTableGraphicFrame(TableRow[] TableRows)
        {
            A.GraphicData GraphicData = new(CreateTable(TableRows))
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
            };
            graphicFrame.NonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties(
               new P.NonVisualDrawingProperties()
               {
                   Id = 1,
                   Name = "Table 1"
               },
               new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoGrouping = true }),
               new P.ApplicationNonVisualDrawingProperties());
            graphicFrame.Graphic = new A.Graphic()
            {
                GraphicData = GraphicData
            };
            graphicFrame.Transform = new P.Transform()
            {
                Offset = new A.Offset()
                {
                    X = tableSetting.x,
                    Y = tableSetting.y
                },
                Extents = new A.Extents()
                {
                    Cx = tableSetting.width,
                    Cy = tableSetting.height
                }
            };
        }

        private A.TableGrid CreateTableGrid(int ColumnCount)
        {
            A.TableGrid TableGrid = new();
            if (tableSetting.widthType == TableSetting.WidthOptionValues.AUTO)
            {
                for (int i = 0; i < ColumnCount; i++)
                {
                    TableGrid.Append(new A.GridColumn() { Width = tableSetting.width / ColumnCount });
                }
            }
            else
            {
                for (int i = 0; i < ColumnCount; i++)
                {
                    TableGrid.Append(new A.GridColumn() { Width = CalculateColumnWidth(tableSetting.widthType, tableSetting.tableColumnWidth[i]) });
                }
            }
            return TableGrid;
        }

        private A.TableRow CreateTableRow(TableRow Row)
        {
            A.TableRow TableRow = new()
            {
                Height = Row.height
            };
            foreach (TableCell cell in Row.tableCells)
            {
                TableRow.Append(CreateTableCell(cell, Row));
            }
            return TableRow;
        }

        private void ReCalculateColumnWidth()
        {
            A.Table? Table = graphicFrame!.Graphic!.GraphicData!.GetFirstChild<A.Table>();
            if (Table != null)
            {
                List<A.GridColumn> GridColumn = Table.TableGrid!.Elements<A.GridColumn>().ToList();
                if (tableSetting.widthType == TableSetting.WidthOptionValues.AUTO)
                {
                    GridColumn.ForEach(Column => Column.Width = tableSetting.width / GridColumn.Count);
                }
                else
                {
                    GridColumn.Select((item, index) => (item, index)).ToList().ForEach(Column =>
                        Column.item.Width = CalculateColumnWidth(tableSetting.widthType, tableSetting.tableColumnWidth[Column.index]));
                }
            }
        }

        #endregion Private Methods
    }
}