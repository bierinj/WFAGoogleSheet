namespace WFAGoolgeSheet
{
    internal class DataGridViewCellMouseClickArgs
    {
        private int rowIdx;
        private int v;

        public DataGridViewCellMouseClickArgs(int rowIdx, int v)
        {
            this.rowIdx = rowIdx;
            this.v = v;
        }
    }
}