using System.Diagnostics.CodeAnalysis;

namespace Excely
{
    public struct CellLocation
    {
        public int Row { get; set; }
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
