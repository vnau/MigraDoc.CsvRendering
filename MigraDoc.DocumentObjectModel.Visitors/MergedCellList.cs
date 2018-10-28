#region MigraDoc - Creating Documents on the Fly
//
// Authors:
//   Stefan Lange (mailto:Stefan.Lange@pdfsharp.com)
//   Klaus Potzesny (mailto:Klaus.Potzesny@pdfsharp.com)
//   David Stephensen (mailto:David.Stephensen@pdfsharp.com)
//
// Copyright (c) 2001-2009 empira Software GmbH, Cologne (Germany)
//
// http://www.pdfsharp.com
// http://www.migradoc.com
// http://sourceforge.net/projects/pdfsharp
//
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included
// in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
// THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
// DEALINGS IN THE SOFTWARE.
//
// Andrew Tsekhansky (mailto:pakeha07@gmail.com): Table rendering optimization in 2010
#endregion

using MigraDoc.DocumentObjectModel.Tables;
using System;
using System.Collections;
using System.Collections.Generic;

namespace MigraDoc.DocumentObjectModel.InternalVisitors
{
    /// <summary>
    /// Represents a merged list of cells of a table.
    /// </summary>
    internal class MergedCellList
    {

        /// <summary>
        /// Enumeration of neighbor positions of cells in a table.
        /// </summary>
        private enum NeighborPosition
        {
            Top,
            Left,
            Right,
            Bottom
        }

        /// <summary>
        /// Initializes a new instance of the MergedCellList class.
        /// </summary>
        public MergedCellList(Table table)
        {
            Init(table);
        }

        private int RowCount;
        private int ColCount;
        private Table Table;
        private CellInfo[,] CellInfos;
        private Hashtable CellIndex;

        /// <summary>
        /// Initializes this instance from a table.
        /// </summary>
        private void Init(Table table)
        {
            Table = table;
            RowCount = Table.Rows.Count;
            ColCount = Table.Columns.Count;

            CellIndex = new Hashtable();
            CellInfos = new CellInfo[RowCount, ColCount];
            for (int rwIdx = 0; rwIdx < RowCount; rwIdx++)
            {
                for (int clmIdx = 0; clmIdx < ColCount; clmIdx++)
                {
                    Cell cell = table[rwIdx, clmIdx];
                    if (CellInfos[rwIdx, clmIdx] == null)
                    {
                        for (int mx = 0; mx <= cell.MergeRight; mx++)
                            for (int my = 0; my <= cell.MergeDown; my++)
                            {
                                bool isMergedCell = mx > 0 || my > 0;
                                CellInfo cellInfo = new CellInfo();
                                cellInfo.TableCell = table[rwIdx + my, clmIdx + mx];
                                cellInfo.Cell = cell;
                                cellInfo.IsMergedCell = isMergedCell;
                                cellInfo.MergedWith = isMergedCell ? CellInfos[rwIdx, clmIdx] : null;
                                cellInfo.Row = rwIdx;
                                cellInfo.Col = clmIdx;

                                CellInfos[rwIdx + my, clmIdx + mx] = cellInfo;
                                CellIndex[cellInfo.TableCell] = CellInfos[rwIdx, clmIdx];
                            }
                    }
                }
            }
        }

        public IEnumerable<Cell> GetCells()
        {
            return GetCells(0, this.RowCount - 1);
        }

        public IEnumerable<Cell> GetRowCells(int row)
        {
            return GetCells(row, row);
        }

        public IEnumerable<Cell> GetCells(int startRow, int endRow)
        {
            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                for (int colIndex = 0; colIndex < ColCount; colIndex++)
                {
                    CellInfo cellInfo = CellInfos[rowIndex, colIndex];
                    if (!cellInfo.IsMergedCell)
                        yield return cellInfo.Cell;
                    colIndex += cellInfo.Cell.MergeRight;
                }
        }

        /// <summary>
        /// Gets the cell that covers the given cell by merging. Usually the cell itself if not merged.
        /// </summary>
        public Cell GetCoveringCell(Cell cell)
        {
            return ((CellInfo)CellIndex[cell]).Cell;
        }

        /// <summary>
        /// Gets the specified cell's uppermost neighbor at the specified position.
        /// </summary>
        private Cell GetNeighbor(CellInfo cellInfo, NeighborPosition position)
        {
            Cell cell = cellInfo.Cell;

            switch (position)
            {
                case NeighborPosition.Left:
                    if (cellInfo.BlockCol > 0)
                        return CellInfos[cellInfo.BlockRow, cellInfo.BlockCol - 1].Cell;
                    break;

                case NeighborPosition.Right:
                    if (cellInfo.BlockCol + cell.MergeRight < ColCount - 1)
                        return CellInfos[cellInfo.BlockRow, cellInfo.BlockCol + cell.MergeRight + 1].Cell;
                    break;

                case NeighborPosition.Top:
                    if (cellInfo.BlockRow > 0)
                        return CellInfos[cellInfo.BlockRow - 1, cellInfo.BlockCol].Cell;
                    break;

                case NeighborPosition.Bottom:
                    if (cellInfo.BlockRow + cell.MergeDown < RowCount - 1)
                        return CellInfos[cellInfo.BlockRow + cell.MergeDown + 1, cellInfo.BlockCol].Cell;
                    break;
            }

            return null;
        }

        public int CalcLastConnectedRow(int row)
        {
            int lastConnectedRow = row;
            for (int rowIndex = row; rowIndex <= lastConnectedRow && rowIndex < this.RowCount; rowIndex++)
            {
                int downConnection = rowIndex;
                for (int colIndex = 0; colIndex < this.ColCount; colIndex++)
                {
                    CellInfo cellInfo = CellInfos[rowIndex, colIndex];
                    downConnection = Math.Max(downConnection,
                      cellInfo.BlockRow + Math.Max(cellInfo.Cell.Row.KeepWith, cellInfo.Cell.MergeDown));
                    colIndex += cellInfo.Cell.MergeRight;
                }
                lastConnectedRow = Math.Max(lastConnectedRow, downConnection);
            }

            return Math.Min(this.RowCount - 1, lastConnectedRow);
        }

        public int CalcLastConnectedColumn(int column)
        {
            int lastConnectedColumn = column;

            for (int colIndex = column; colIndex <= lastConnectedColumn && colIndex < this.ColCount; colIndex++)
            {
                int rightConnection = column;
                for (int rowIndex = 0; rowIndex < this.RowCount; rowIndex++)
                {
                    CellInfo cellInfo = CellInfos[rowIndex, colIndex];
                    cellInfo = cellInfo.MergedWith ?? cellInfo;
                    rightConnection = Math.Max(rightConnection,
                      cellInfo.BlockCol + Math.Max(cellInfo.Cell.Column.KeepWith, cellInfo.Cell.MergeRight));
                    rowIndex += cellInfo.Cell.MergeDown;
                }
                lastConnectedColumn = Math.Max(lastConnectedColumn, rightConnection);
            }
            return Math.Min(lastConnectedColumn, this.ColCount);
        }

        public int GetFirstRowMergedWithRow(int row)
        {
            if (row == 0 || row >= this.RowCount)
                return row;

            var result = row;

            for (int colIndex = 0; colIndex < ColCount; colIndex++)
            {
                CellInfo cellInfo = CellInfos[row, colIndex];
                result = Math.Min(result, cellInfo.BlockRow);
                colIndex += cellInfo.Cell.MergeRight;
            }

            return result;
        }
    }

    internal class CellInfo
    {

        // Cell from table at given Row, Col
        public Cell TableCell;

        // Cell which fills given Row, Col
        public Cell Cell;

        // Whether cell is merged with another cell
        public bool IsMergedCell;

        // CellInfo this cell is merged with
        public CellInfo MergedWith;

        // Cell's Row 
        public int Row;

        // Cells' Col
        public int Col;

        // Row where merged area starts
        public int BlockRow
        {
            get { return MergedWith == null ? Row : MergedWith.Row; }
        }

        // Col where merged area starts
        public int BlockCol
        {
            get { return MergedWith == null ? Col : MergedWith.Col; }
        }

    }
}


