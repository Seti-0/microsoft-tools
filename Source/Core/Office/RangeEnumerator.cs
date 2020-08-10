using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace Red.Core.Office
{
    using ExcelRange = Microsoft.Office.Interop.Excel.Range;

    public class RangeEnumerable : IEnumerable<ExcelRange>
    {
        private ExcelRange _range;

        public RangeEnumerable(ExcelRange range)
        {
            _range = range;
        }

        public IEnumerator<ExcelRange> GetEnumerator()
        {
            return new RangeEnumerator(_range);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new RangeEnumerator(_range);
        }
    }

    public class RangeEnumerator : IEnumerator<ExcelRange>
    {
        private ExcelRange totalRange;

        private bool[,] _mask;

        public int Width { get; private set; }
        public int Height { get; private set; }

        public int RowIndex { get; private set; }
        public int ColumnIndex { get; private set; }

        public float Progress
        {
            get
            {
                if (RowIndex < 0 || ColumnIndex < 0)
                    return 0;

                if (Width * Height == 0)
                    return 1;

                return ((float)((Width * RowIndex) + ColumnIndex)) / (Width * Height);
            }
        }

        public ExcelRange Current
        {
            get
            {
                if (RowIndex == -1 || ColumnIndex == -1)
                    return null;

                if (RowIndex >= Height || ColumnIndex >= Width)
                    return null;

                return totalRange.Cells[RowIndex + 1, ColumnIndex + 1];
            }
        }

        object IEnumerator.Current => Current;

        public RangeEnumerator(ExcelRange range)
        {
            if (range == null)
                throw new ArgumentNullException(nameof(range));

            totalRange = range;

            Width = range.Columns.Count;
            Height = range.Rows.Count;

            Reset();
        }

        public void ApplyMask(bool[,] mask)
        {
            if (mask.GetLength(0) != Height)
                throw new ArgumentException($"Mask width {mask.GetLength(0)} does not match range height {Height}");

            if (mask.GetLength(1) != Width)
                throw new ArgumentException($"Mask width {mask.GetLength(1)} does not match range width {Width}");

            _mask = mask;
        }

        public void Dispose(){}

        public bool MoveNext()
        {
            if (Width * Height == 0)
                return false;

            if (RowIndex == -1)
            {
                RowIndex = 0;
                ColumnIndex = 0;
                return true;
            }

            ColumnIndex++;
            if (ColumnIndex >= Width)
            {
                ColumnIndex = 0;
                RowIndex++;
            }

            if (RowIndex >= Height)
                return false;

            if (_mask != null && !_mask[RowIndex, ColumnIndex])
                return MoveNext();

            return true;
        }

        public void Reset()
        {
            ColumnIndex = -1;
            RowIndex = -1;
        }
    }
}
