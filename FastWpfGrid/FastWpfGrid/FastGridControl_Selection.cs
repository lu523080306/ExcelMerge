using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace FastWpfGrid
{
    partial class FastGridControl
    {
        public event EventHandler<SelectionChangedEventArgs> SelectedCellsChanged;

        private HashSet<FastGridCellAddress> _selectedCells = new HashSet<FastGridCellAddress>();
        private Dictionary<int, int> _selectedRows = new Dictionary<int, int>();
        private Dictionary<int, int> _selectedColumns = new Dictionary<int, int>();

        public IEnumerable<FastGridCellAddress> SelectedCells { get { return _selectedCells.AsEnumerable(); } }

        int? _selectedRealRowCountLimit;
        bool _selectedRealRowCountLimitLoaded;
        public int? SelectedRealRowCountLimit
        {
            get
            {
                if (_selectedRealRowCountLimitLoaded) return _selectedRealRowCountLimit;
                _selectedRealRowCountLimitLoaded = true;
                if (Model == null)
                {
                    _selectedRealRowCountLimit = null;
                }
                else
                {
                    _selectedRealRowCountLimit = IsTransposed ? Model.SelectedColumnCountLimit : Model.SelectedRowCountLimit;
                }
                return _selectedRealRowCountLimit;
            }
        }

        int? _selectedRealColumnCountLimit;
        bool _selectedRealColumnCountLimitLoaded;
        public int? SelectedRealColumnCountLimit
        {
            get
            {
                if (_selectedRealColumnCountLimitLoaded) return _selectedRealColumnCountLimit;
                _selectedRealColumnCountLimitLoaded = true;
                if (Model == null)
                {
                    _selectedRealColumnCountLimit = null;
                }
                else
                {
                    _selectedRealColumnCountLimit = IsTransposed ? Model.SelectedRowCountLimit : Model.SelectedColumnCountLimit;
                }
                return _selectedRealColumnCountLimit;
            }
        }

        private bool _isLimitedSelection = false;

        private void CheckChangedLimitedSelection()
        {
            if (IsLimitedSelection != _isLimitedSelection)
            {
                InvalidateAll();
                _isLimitedSelection = IsLimitedSelection;
            }
        }

        private void ClearSelectedCells()
        {
            _selectedCells.Clear();
            _selectedRows.Clear();
            _selectedColumns.Clear();

            CheckChangedLimitedSelection();
        }

        public bool AddSelectedCell(FastGridCellAddress cell)
        {
            if (!cell.IsCell) return false;

            if (SelectedRealRowCountLimit.HasValue && _selectedRows.Count >= SelectedRealRowCountLimit.Value && !_selectedRows.ContainsKey(cell.Row.Value)) return false;
            if (SelectedRealColumnCountLimit.HasValue && _selectedColumns.Count >= SelectedRealColumnCountLimit.Value && !_selectedColumns.ContainsKey(cell.Column.Value)) return false;

            if (_selectedCells.Contains(cell)) return false;

            _selectedCells.Add(cell);

            if (!_selectedRows.ContainsKey(cell.Row.Value)) _selectedRows[cell.Row.Value] = 0;
            _selectedRows[cell.Row.Value]++;

            if (!_selectedColumns.ContainsKey(cell.Column.Value)) _selectedColumns[cell.Column.Value] = 0;
            _selectedColumns[cell.Column.Value]++;

            CheckChangedLimitedSelection();

            return true;
        }

        /// <summary>
        /// 用给定的 real 坐标整体替换当前选择。
        /// 不触发 SelectedCellsChanged：供两侧网格镜像选择时调用，
        /// 避免“设置当前格 → 触发同步 → 再清空选择”形成递归并互相覆盖。
        /// </summary>
        public void SetSelectedCells(IEnumerable<FastGridCellAddress> cells)
        {
            _selectedCells.ToList().ForEach(InvalidateCell);
            ClearSelectedCells();

            foreach (var cell in cells)
            {
                // 越界 real 坐标（点击/拖拽到数据区外产生）不能进入选择集合：
                // 它们渲染不出来，却会让两侧镜像选择看起来不一致。
                if (!IsRealCellInRange(cell))
                    continue;

                AddSelectedCell(cell);
                InvalidateCell(cell);
            }
        }

        private void RemoveSelectedCell(FastGridCellAddress cell)
        {
            if (!cell.IsCell) return;

            if (!_selectedCells.Contains(cell)) return;

            _selectedCells.Remove(cell);

            if (_selectedRows.ContainsKey(cell.Row.Value))
            {
                _selectedRows[cell.Row.Value]--;
                if (_selectedRows[cell.Row.Value] == 0) _selectedRows.Remove(cell.Row.Value);
            }

            if (_selectedColumns.ContainsKey(cell.Column.Value))
            {
                _selectedColumns[cell.Column.Value]--;
                if (_selectedColumns[cell.Column.Value] == 0) _selectedColumns.Remove(cell.Column.Value);
            }

            CheckChangedLimitedSelection();
        }

        public bool IsLimitedSelection
        {
            get
            {
                return (SelectedRealRowCountLimit.HasValue && _selectedRows.Count >= SelectedRealRowCountLimit.Value)
                    ||
                     (SelectedRealColumnCountLimit.HasValue && _selectedColumns.Count >= SelectedRealColumnCountLimit.Value);
            }
        }

        private void SetSelectedRectangle(FastGridCellAddress origin, FastGridCellAddress cell)
        {
            var newSelected = GetCellRange(origin, cell);
            foreach (var added in newSelected)
            {
                if (_selectedCells.Contains(added)) continue;
                InvalidateCell(added);
                AddSelectedCell(added);
            }
            foreach (var removed in _selectedCells.ToList())
            {
                if (newSelected.Contains(removed)) continue;
                InvalidateCell(removed);
                RemoveSelectedCell(removed);
            }
            SetCurrentCell(cell);
            OnChangeSelectedCells(true);
        }

        private void OnChangeSelectedCells(bool isInvokedByUser)
        {
            if (SelectedCellsChanged != null) SelectedCellsChanged(this, new SelectionChangedEventArgs { IsInvokedByUser = isInvokedByUser });
        }
    }
}
