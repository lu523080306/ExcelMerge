namespace ExcelMerge
{
    public class ExcelCell
    {
        public string Value { get; private set; }
        public int OriginalColumnIndex { get; private set; }
        public int OriginalRowIndex { get; private set; }

        /// <summary>合成格：对齐/列结构补位产生，没有真实物理坐标，禁止用于定位写回。</summary>
        public bool IsSynthetic { get; private set; }

        public ExcelCell(string value, int originalColumnIndex, int originalRowIndex)
            : this(value, originalColumnIndex, originalRowIndex, false)
        {
        }

        public ExcelCell(string value, int originalColumnIndex, int originalRowIndex, bool isSynthetic)
        {
            Value = value;
            OriginalColumnIndex = originalColumnIndex;
            OriginalRowIndex = originalRowIndex;
            IsSynthetic = isSynthetic;
        }
    }
}
