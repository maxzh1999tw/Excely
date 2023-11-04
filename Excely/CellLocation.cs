using System.Diagnostics.CodeAnalysis;

namespace Excely
{
    /// <summary>
    /// 代表 Cell 位置的結構。
    /// </summary>
    public struct CellLocation
    {
        /// <summary>
        /// 從 0 開始。
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// 從 0 開始。
        /// </summary>
        public int Column { get; set; }

        public CellLocation(int row, int column)
        {
            Row = row;
            Column = column;
        }

        public override readonly bool Equals([NotNullWhen(true)] object? obj)
        {
            if (obj is CellLocation otherCell)
            {
                return Row == otherCell.Row && Column == otherCell.Column;
            }

            return false;
        }

        public static bool operator ==(CellLocation left, CellLocation right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(CellLocation left, CellLocation right)
        {
            return !(left == right);
        }

        public override readonly int GetHashCode()
        {
            return HashCode.Combine(Row, Column);
        }
    }
}
