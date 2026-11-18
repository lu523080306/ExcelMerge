using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Practices.Unity;
using FastWpfGrid;
using NetDiff;
using SKCore.Collection;
using ExcelMerge.GUI.ViewModels;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Models;
using ExcelMerge.GUI.Styles;

namespace ExcelMerge.GUI.Views
{
    public partial class DiffView : UserControl
    {
        private ExcelSheetDiffConfig diffConfig = new ExcelSheetDiffConfig();
        private ExcelSheetDiff currentSheetDiff;
        private IUnityContainer container;
        private FastGridControl pendingColumnGrid;
        private int? pendingColumnIndex;
        private bool isLock;
        private const string srcKey = "src";
        private const string dstKey = "dst";

        private FastGridControl copyTargetGrid;

        public DiffView()
        {
            InitializeComponent();
            InitializeContainer();
            InitializeEventListeners();

            SrcDataGrid.ColumnHeaderClick += (s, e) => SetPendingColumn(SrcDataGrid, e.Column);
            DstDataGrid.ColumnHeaderClick += (s, e) => SetPendingColumn(DstDataGrid, e.Column);

            App.Instance.OnSettingUpdated += OnApplicationSettingUpdated;

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            // In order to enable Ctrl + F immediately after startup.
            ToolExpander.IsExpanded = true;
        }

        private DiffViewModel GetViewModel()
        {
            return DataContext as DiffViewModel;
        }

        private void InitializeContainer()
        {
            container = new UnityContainer();
            container
                .RegisterInstance(srcKey, SrcDataGrid)
                .RegisterInstance(dstKey, DstDataGrid)
                .RegisterInstance(srcKey, SrcLocationGrid)
                .RegisterInstance(dstKey, DstLocationGrid)
                .RegisterInstance(srcKey, SrcViewRectangle)
                .RegisterInstance(dstKey, DstViewRectangle)
                .RegisterInstance(srcKey, SrcValueTextBox)
                .RegisterInstance(dstKey, DstValueTextBox);
        }

        private void InitializeEventListeners()
        {
            var srcEventHandler = new DiffViewEventHandler(srcKey);
            var dstEventHandler = new DiffViewEventHandler(dstKey);

            DataGridEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            DataGridEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            LocationGridEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            LocationGridEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            ViewportEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            ViewportEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(dstEventHandler);
        }

        private void OnApplicationSettingUpdated()
        {
            var e = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchApplicationSettingUpdateEvent(e);
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchParentLoadEvent(args);

            ExecuteDiff(isStartup: true);

            // In order to enable Ctrl + F immediately after startup.
            ToolExpander.IsExpanded = false;
        }

        private ExcelSheetDiffConfig CreateDiffConfig(FileSetting srcFileSetting, FileSetting dstFileSetting, bool isStartup)
        {
            var config = new ExcelSheetDiffConfig();

            config.SrcSheetIndex = SrcSheetCombobox.SelectedIndex;
            config.DstSheetIndex = DstSheetCombobox.SelectedIndex;

            if (srcFileSetting != null)
            {
                if (isStartup)
                    config.SrcSheetIndex = GetSheetIndex(srcFileSetting, SrcSheetCombobox.Items);

                config.SrcHeaderIndex = srcFileSetting.ColumnHeaderIndex;
            }

            if (dstFileSetting != null)
            {
                if (isStartup)
                    config.DstSheetIndex = GetSheetIndex(dstFileSetting, DstSheetCombobox.Items);

                config.DstHeaderIndex = dstFileSetting.ColumnHeaderIndex;
            }

            return config;
        }

        private int GetSheetIndex(FileSetting fileSetting, ItemCollection sheetNames)
        {
            if (fileSetting == null)
                return -1;

            var index = fileSetting.SheetIndex;
            if (!string.IsNullOrEmpty(fileSetting.SheetName))
                index = sheetNames.IndexOf(fileSetting.SheetName);

            if (index < 0 || index >= sheetNames.Count)
            {
                MessageBox.Show(Properties.Resources.Msg_OutofSheetRange);
                index = 0;
            }

            return index;
        }

        private void LocationGrid_MouseDown(object sender, MouseEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
        }

        private void LocationGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
                LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
            }
        }

        private void LocationGrid_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseWheelEvent(args, e);
        }

        private void DataGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSizeChangeEvent(args, e);
        }

        private void LocationGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchSizeChangeEvent(args, e);
        }

        private void DataGrid_SelectedCellsChanged(object sender, FastWpfGrid.SelectionChangedEventArgs e)
        {
            var grid = copyTargetGrid = sender as FastGridControl;
            if (grid == null)
                return;

            // 点数据格 = 放弃之前的列头列意图
            if (pendingColumnGrid == grid && grid.CurrentCell != null && !grid.CurrentCell.IsColumnHeader)
                pendingColumnIndex = null;

            copyTargetGrid = grid;

            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSelectedCellChangeEvent(args);

            if (!SrcDataGrid.CurrentCell.Row.HasValue || !DstDataGrid.CurrentCell.Row.HasValue)
                return;

            if (!SrcDataGrid.CurrentCell.Column.HasValue || !DstDataGrid.CurrentCell.Column.HasValue)
                return;

            if (SrcDataGrid.Model == null || DstDataGrid.Model == null)
                return;

            var srcValue =
                (SrcDataGrid.Model as DiffGridModel).GetCellText(SrcDataGrid.CurrentCell.Row.Value, SrcDataGrid.CurrentCell.Column.Value, true);
            var dstValue =
                (DstDataGrid.Model as DiffGridModel).GetCellText(DstDataGrid.CurrentCell.Row.Value, DstDataGrid.CurrentCell.Column.Value, true);

            UpdateValueDiff(srcValue, dstValue);

            if (App.Instance.Setting.AlwaysExpandCellDiff)
            {
                var a = new DiffViewEventArgs<RichTextBox>(null, container, TargetType.First);
                ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(a);
            }
        }

        private void ValueTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container, TargetType.First);
            ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(args);
        }

        private void ValueTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container, TargetType.First);
            ValueTextBoxEventDispatcher.Instance.DispatchLostFocusEvent(args);
        }

        private string GetRichTextString(RichTextBox textBox)
        {
            var textRange = new TextRange(textBox.Document.ContentStart, textBox.Document.ContentEnd);

            return textRange.Text;
        }

        private IEnumerable<DiffResult<string>> DiffCellValue(IEnumerable<string> src, IEnumerable<string> dst)
        {
            var r = DiffUtil.Diff(src, dst);
            r = DiffUtil.Order(r, DiffOrderType.LazyDeleteFirst);
            return DiffUtil.OptimizeCaseDeletedFirst(r);
        }

        private string ConvertWhiteSpaces(string str)
        {
            return new string(str.Select(c =>
            {
                if (Encoding.UTF8.GetByteCount(c.ToString()) == 1)
                    return ' ';
                else
                    return '　';

            }).ToArray());
        }

        private string ConvertWhiteSpaces(char c)
        {
            if (Encoding.UTF8.GetByteCount(c.ToString()) == 1)
                return " ";
            else
                return "　";
        }

        private void DiffModifiedLine(IEnumerable<DiffResult<char>> results, List<Tuple<string, Color?>> ranges, bool isSrc)
        {
            var splited = results.SplitByRegularity((items, current) => items.Last().Status.Equals(current.Status)).ToList();

            foreach (var sr in splited)
            {
                var status = sr.First().Status;
                if (status == DiffStatus.Equal)
                {
                    ranges.Add(Tuple.Create<string, Color?>(new string(sr.Select(r => r.Obj1).ToArray()), null));
                }
                else if (status == DiffStatus.Modified)
                {
                    var str = new string(sr.Select(r => isSrc ? r.Obj1 : r.Obj2).ToArray());
                    ranges.Add(Tuple.Create<string, Color?>(str, EMColor.LightOrange));
                }
                else if (status == DiffStatus.Deleted)
                {
                    var str = new string(sr.Select(r => r.Obj1).ToArray());
                    ranges.Add(Tuple.Create<string, Color?>(str, EMColor.LightGray));
                }
                else if (status == DiffStatus.Inserted)
                {
                    var str = new string(sr.Select(r => r.Obj2).ToArray());
                    ranges.Add(Tuple.Create<string, Color?>(str, EMColor.Orange));
                }
            }

            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void DiffEqualLine(DiffResult<string> lineDiffResult, List<Tuple<string, Color?>> ranges)
        {
            ranges.Add(Tuple.Create<string, Color?>(lineDiffResult.Obj1, null));
            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void DiffDeletedLine(DiffResult<string> lineDiffResult, List<Tuple<string, Color?>> ranges, bool isSrc)
        {
            var str = isSrc ? lineDiffResult.Obj1 : ConvertWhiteSpaces(lineDiffResult.Obj1.ToString());
            ranges.Add(Tuple.Create<string, Color?>(str, isSrc ? EMColor.LightGray : EMColor.LightGray));
            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void DiffInsertedLine(DiffResult<string> lineDiffResult, List<Tuple<string, Color?>> ranges, bool isSrc)
        {
            var str = isSrc ? ConvertWhiteSpaces(lineDiffResult.Obj2) : lineDiffResult.Obj2;
            ranges.Add(Tuple.Create<string, Color?>(str, isSrc ? EMColor.LightGray : EMColor.Orange));
            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void UpdateValueDiff(string srcValue, string dstValue)
        {
            SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Clear();
            DstValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Clear();

            var srcLines = srcValue.Split('\n').Select(s => s.TrimEnd());
            var dstLines = dstValue.Split('\n').Select(s => s.TrimEnd());

            var lineDiffResults = DiffCellValue(srcLines, dstLines).ToList();

            var srcRange = new List<Tuple<string, Color?>>();
            var dstRange = new List<Tuple<string, Color?>>();
            foreach (var lineDiffResult in lineDiffResults)
            {
                if (lineDiffResult.Status == DiffStatus.Equal)
                {
                    DiffEqualLine(lineDiffResult, srcRange);
                    DiffEqualLine(lineDiffResult, dstRange);
                }
                else if (lineDiffResult.Status == DiffStatus.Modified)
                {
                    var charDiffResults = DiffUtil.Diff(lineDiffResult.Obj1, lineDiffResult.Obj2);
                    charDiffResults = DiffUtil.Order(charDiffResults, DiffOrderType.LazyDeleteFirst);
                    charDiffResults = DiffUtil.OptimizeCaseDeletedFirst(charDiffResults);

                    DiffModifiedLine(charDiffResults.Where(r => r.Status != DiffStatus.Inserted), srcRange, true);
                    DiffModifiedLine(charDiffResults.Where(r => r.Status != DiffStatus.Deleted), dstRange, false);
                }
                else if (lineDiffResult.Status == DiffStatus.Deleted)
                {
                    DiffDeletedLine(lineDiffResult, srcRange, true);
                    DiffDeletedLine(lineDiffResult, dstRange, false);
                }
                else if (lineDiffResult.Status == DiffStatus.Inserted)
                {
                    DiffInsertedLine(lineDiffResult, srcRange, true);
                    DiffInsertedLine(lineDiffResult, dstRange, false);
                }
            }

            foreach (var r in srcRange)
            {
                var bc = r.Item2.HasValue ? new SolidColorBrush(r.Item2.Value) : new SolidColorBrush();
                SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Add(new Run(r.Item1) { Background = bc });
            }

            foreach (var r in dstRange)
            {
                var bc = r.Item2.HasValue ? new SolidColorBrush(r.Item2.Value) : new SolidColorBrush();
                DstValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Add(new Run(r.Item1) { Background = bc });
            }
        }

        private void DiffButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteDiff();
        }

        private void ApplyToDstButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyDiff(ExcelDiffDirection.SrcToDst);
        }

        private void ApplyToSrcButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyDiff(ExcelDiffDirection.DstToSrc);
        }

        private void ApplyDiff(ExcelDiffDirection direction)
        {
            if (currentSheetDiff == null)
            {
                MessageBox.Show("请先执行一次对比。", "应用改动", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var srcToDst = direction == ExcelDiffDirection.SrcToDst;
            var targetPath = srcToDst ? DstPathTextBox.Text : SrcPathTextBox.Text;
            var targetSheet = (srcToDst ? DstSheetCombobox.SelectedItem : SrcSheetCombobox.SelectedItem)?.ToString();

            var applier = new ExcelDiffApplier();

            // ① 整列选中 → 只同步这些列（列优先于行级：拖选整列会顺带触到个别结构行）
            var colInfo = CollectSelectedColumns(direction);

            // 列头点击＝明确列意图：优先于选择检测（消除"新增多个行"时误判为行级）
            if (colInfo == null && pendingColumnIndex.HasValue && pendingColumnGrid != null)
            {
                colInfo = Tuple.Create(new HashSet<int> { pendingColumnIndex.Value }, true);
                pendingColumnIndex = null;
            }

            if (colInfo != null)
            {
                var colNames = string.Join("、", colInfo.Item1.OrderBy(c => c).Select(c => DiffGridModel.GetExcelColumnLetter(c) + "列"));
                var confirmCol = MessageBox.Show(
                    string.Format("从 {0} 同步列 {1} 到 {2}。\n{3}\n\n目标文件: {4}",
                        srcToDst ? "左侧" : "右侧", colNames, srcToDst ? "右侧" : "左侧",
                        DescribeColumnOps(colInfo.Item1, direction), targetPath),
                    "同步选中列", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (confirmCol != MessageBoxResult.Yes)
                    return;

                var rc = applier.ApplyColumns(currentSheetDiff, targetPath, targetSheet, direction, colInfo.Item1);
                MessageBox.Show(rc.Message, "应用结果", MessageBoxButton.OK,
                    rc.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Error);

                if (rc.IsSuccess)
                    ExecuteDiff();

                return;
            }

            // ② 行级：选中整行 或 选中结构行(新增/删除行) → 只同步这些行
            var rowInfo = CollectSelectedRowsAndMode(direction);
            if (rowInfo != null && rowInfo.Item2)
            {
                var confirm = MessageBox.Show(
                    string.Format("从 {0} 同步选中的 {1} 行到 {2}。\n{3}\n\n目标文件: {4}",
                        srcToDst ? "左侧" : "右侧", rowInfo.Item1.Count, srcToDst ? "右侧" : "左侧",
                        DescribeRowOperations(rowInfo.Item1, direction), targetPath),
                    "同步选中行", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (confirm != MessageBoxResult.Yes)
                    return;

                var r = applier.ApplyRows(currentSheetDiff, targetPath, targetSheet, direction, rowInfo.Item1);
                MessageBox.Show(r.Message, "应用结果", MessageBoxButton.OK,
                    r.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Error);

                if (r.IsSuccess)
                    ExecuteDiff();

                return;
            }

            // ③ 格子级：选中部分值差异格 → 只应用这些格
            var selectedChanges = CollectSelectedCellChanges(direction);
            if (selectedChanges != null && selectedChanges.Count > 0)
            {
                var confirm = MessageBox.Show(
                    string.Format("从 {0} 应用选中的 {1} 个格子的值到 {2}。\n\n目标文件: {3}",
                        srcToDst ? "左侧" : "右侧", selectedChanges.Count, srcToDst ? "右侧" : "左侧", targetPath),
                    "应用选中格", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (confirm != MessageBoxResult.Yes)
                    return;

                var r = applier.ApplyExplicit(targetPath, targetSheet, selectedChanges);
                MessageBox.Show(r.Message, "应用结果", MessageBoxButton.OK,
                    r.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Error);

                if (r.IsSuccess)
                    ExecuteDiff();

                return;
            }

            // 有选中但无法归类 → 提示而不是默默全量同步
            if (SrcDataGrid.SelectedCells.Any() || DstDataGrid.SelectedCells.Any())
            {
                MessageBox.Show(
                    "当前选择无法识别为要同步的行/列/格子。\n\n操作方法：\n• 同步整列 → 点击该列列头\n• 同步整行 → 选中整行\n• 同步格子 → 选中值差异格\n再点 ← / →。\n\n若确需全量同步，请先取消所有选中。",
                    "无法识别选择", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // ③ 无选中：全量同步（含行列结构级）
            var confirmAll = MessageBox.Show(
                string.Format("从 {0} 应用全部差异到 {1}：\n{2}\n\n目标文件: {3}\n\n将覆盖目标文件中的对应格子/行列，确定继续？",
                    srcToDst ? "左侧" : "右侧", srcToDst ? "右侧" : "左侧",
                    DescribeFullDiff(direction), targetPath),
                "应用全部差异", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (confirmAll != MessageBoxResult.Yes)
                return;

            var result = applier.Apply(currentSheetDiff, targetPath, targetSheet, direction);

            MessageBox.Show(result.Message, "应用结果", MessageBoxButton.OK,
                result.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Error);

            if (result.IsSuccess)
                ExecuteDiff();
        }

        /// <summary>把源侧网格上选中的格子收集为待应用改动（只取 Modified 格）；无选中返回 null。</summary>
        private List<CellChange> CollectSelectedCellChanges(ExcelDiffDirection direction)
        {
            var srcToDst = direction == ExcelDiffDirection.SrcToDst;
            var sourceGrid = GetActiveSelectionGrid(direction);
            var model = sourceGrid.Model as DiffGridModel;

            var changes = new List<CellChange>();
            foreach (var address in sourceGrid.SelectedCells)
            {
                var cellDiff = model?.GetCellDiffForVisualAddress(address);
                if (cellDiff == null)
                    continue;

                // 只有值差异格(Modified)在目标侧有真实落点；结构级(Added/Removed)不适用格子应用
                if (cellDiff.Status != ExcelCellStatus.Modified)
                    continue;

                changes.Add(srcToDst
                    ? new CellChange(
                          cellDiff.DstCell.OriginalRowIndex,
                          cellDiff.DstCell.OriginalColumnIndex,
                          cellDiff.SrcCell.Value)
                    : new CellChange(
                          cellDiff.SrcCell.OriginalRowIndex,
                          cellDiff.SrcCell.OriginalColumnIndex,
                          cellDiff.DstCell.Value));
            }

            return changes.Count == 0 ? null : changes;
        }

        /// <summary>
        /// 收集选中范围所触达的 diff 行索引。第二项=true 表示应走"整行同步"：
        /// 只要有任何一行是整行选中 或 结构行(新增/删除行)，就按行级处理。
        /// 未选中任何行返回 null。
        /// </summary>
        private Tuple<HashSet<int>, bool> CollectSelectedRowsAndMode(ExcelDiffDirection direction)
        {
            var srcToDst = direction == ExcelDiffDirection.SrcToDst;
            var grid = GetActiveSelectionGrid(direction);
            var model = grid.Model as DiffGridModel;
            if (model == null)
                return null;

            var rowIndexes = new HashSet<int>();
            var rowLevel = false;

            foreach (var group in grid.SelectedCells.GroupBy(c => c.Row.Value))
            {
                var diffIndex = model.GetRowDiffIndexForVisualRow(group.Key);
                if (diffIndex < 0)
                    continue;

                rowIndexes.Add(diffIndex);

                ExcelRowDiff rowDiff;
                if (!currentSheetDiff.Rows.TryGetValue(diffIndex, out rowDiff))
                    continue;

                var isStructural = rowDiff.Cells.Count > 0 && (rowDiff.IsAdded() || rowDiff.IsRemoved());
                var fullRow = group.Count() >= model.ColumnCount;
                if (isStructural || fullRow)
                    rowLevel = true;
            }

            return rowIndexes.Count == 0 ? null : Tuple.Create(rowIndexes, rowLevel);
        }

        /// <summary>
        /// 整列选中检测：某列的选中格数 ≥ 网格可视行数（用户拖选整列时会覆盖所有可见行）。
        /// 未选中整列返回 null。
        /// </summary>
        private Tuple<HashSet<int>, bool> CollectSelectedColumns(ExcelDiffDirection direction)
        {
            var srcToDst = direction == ExcelDiffDirection.SrcToDst;
            var grid = GetActiveSelectionGrid(direction);
            var model = grid.Model as DiffGridModel;
            if (model == null)
                return null;

            var cells = grid.SelectedCells
                .Where(c => c.Row.HasValue && c.Column.HasValue)
                .ToList();

            if (cells.Count < 2)
                return null;

            var distinctColumns = cells.Select(c => c.Column.Value).Distinct().Count();
            var distinctRows = cells.Select(c => c.Row.Value).Distinct().Count();

            // 同一列、跨多行，且(该列为结构列 或 覆盖全部行) → 列意图
            if (distinctColumns == 1 && distinctRows >= 2)
            {
                var col = cells[0].Column.Value;
                var isStructural = currentSheetDiff != null && currentSheetDiff.ColumnStatusMap != null
                    && currentSheetDiff.ColumnStatusMap.ContainsKey(col)
                    && (currentSheetDiff.ColumnStatusMap[col] == ExcelColumnStatus.Inserted
                        || currentSheetDiff.ColumnStatusMap[col] == ExcelColumnStatus.Deleted);

                if (isStructural || distinctRows >= model.RowCount)
                    return Tuple.Create(new HashSet<int> { col }, true);
            }

            return null;
        }

        /// <summary>
        /// 选中的来源网格：优先用用户最后点选/选中的网格，方向不再决定"必须选在哪侧"。
        /// </summary>
        private FastGridControl GetActiveSelectionGrid(ExcelDiffDirection direction)
        {
            if (copyTargetGrid != null && copyTargetGrid.SelectedCells.Any())
                return copyTargetGrid;

            return direction == ExcelDiffDirection.SrcToDst ? SrcDataGrid : DstDataGrid;
        }

        /// <summary>列头点击：记录明确的列同步意图。</summary>
        private void SetPendingColumn(FastGridControl grid, int column)
        {
            pendingColumnGrid = grid;
            pendingColumnIndex = column;
        }

        private void ApplyColumnToDst_Click(object sender, RoutedEventArgs e)
        {
            ApplySingleColumn(SrcDataGrid, ExcelDiffDirection.SrcToDst);
        }

        private void ApplyColumnToSrc_Click(object sender, RoutedEventArgs e)
        {
            ApplySingleColumn(DstDataGrid, ExcelDiffDirection.DstToSrc);
        }

        private void ApplySingleColumn(FastGridControl grid, ExcelDiffDirection direction)
        {
            if (currentSheetDiff == null)
            {
                MessageBox.Show("请先执行一次对比。", "同步本列", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var col = grid.CurrentCell.Column;
            if (!col.HasValue)
            {
                MessageBox.Show("无法确定要同步的列，请先点击该列的表头单元格。", "同步本列", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var srcToDst = direction == ExcelDiffDirection.SrcToDst;
            var targetPath = srcToDst ? DstPathTextBox.Text : SrcPathTextBox.Text;
            var targetSheet = (srcToDst ? DstSheetCombobox.SelectedItem : SrcSheetCombobox.SelectedItem)?.ToString();
            var alignedCol = col.Value;

            var model = grid.Model as DiffGridModel;
            var colName = DiffGridModel.GetExcelColumnLetter(alignedCol) + "列";

            var msg = string.Format("从 {0} 同步列 “{1}” 到 {2}。\n{3}\n\n目标文件: {4}",
                srcToDst ? "左侧" : "右侧", colName,
                srcToDst ? "右侧" : "左侧",
                DescribeColumnOps(new HashSet<int> { alignedCol }, direction), targetPath);

            if (MessageBox.Show(msg, "同步本列", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            var applier = new ExcelDiffApplier();
            var r = applier.ApplyColumns(currentSheetDiff, targetPath, targetSheet, direction, new List<int> { alignedCol });
            MessageBox.Show(r.Message, "应用结果", MessageBoxButton.OK,
                r.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Error);

            if (r.IsSuccess)
                ExecuteDiff();
        }

        /// <summary>列的同步摘要（"同步本列/整列"确认弹窗用，不展示数据明细）。</summary>
        private string DescribeColumnOps(HashSet<int> columns, ExcelDiffDirection direction)
        {
            var structural = columns.Any(c => currentSheetDiff.ColumnStatusMap != null
                && currentSheetDiff.ColumnStatusMap.ContainsKey(c)
                && (currentSheetDiff.ColumnStatusMap[c] == ExcelColumnStatus.Inserted
                    || currentSheetDiff.ColumnStatusMap[c] == ExcelColumnStatus.Deleted));

            if (structural)
                return "  含新增/删除列 → 将按方向整体插入/删除到目标文件";

            var total = 0;
            foreach (var c in columns)
            {
                foreach (var row in currentSheetDiff.Rows.Values)
                {
                    ExcelCellDiff cell;
                    if (row.Cells.TryGetValue(c, out cell) && cell.Status == ExcelCellStatus.Modified)
                        total++;
                }
            }

            return "  共 " + total + " 个差异格，将值同步到目标文件";
        }

        /// <summary>行的同步摘要（"同步选中行"确认弹窗用，只统计操作类型，不展示数据）。</summary>
        private string DescribeRowOperations(HashSet<int> diffRowIndexes, ExcelDiffDirection direction)
        {
            var srcToDst = direction == ExcelDiffDirection.SrcToDst;
            var insert = 0;
            var del = 0;
            var modify = 0;

            foreach (var idx in diffRowIndexes)
            {
                ExcelRowDiff rowDiff;
                if (!currentSheetDiff.Rows.TryGetValue(idx, out rowDiff) || rowDiff.Cells.Count == 0)
                    continue;

                if (rowDiff.IsAdded() || rowDiff.IsRemoved())
                {
                    if (srcToDst)
                    {
                        if (rowDiff.IsAdded()) del++;
                        else insert++;
                    }
                    else
                    {
                        if (rowDiff.IsAdded()) insert++;
                        else del++;
                    }
                }
                else
                {
                    modify++;
                }
            }

            var parts = new List<string>();
            if (insert > 0) parts.Add("插入 " + insert + " 行");
            if (del > 0) parts.Add("删除 " + del + " 行");
            if (modify > 0) parts.Add("整行值回写 " + modify + " 行");

            return parts.Count == 0 ? "  (仅值差异格)" : "  " + string.Join("；", parts);
        }

        private string DescribeFullDiff(ExcelDiffDirection direction)
        {
            var srcToDst = direction == ExcelDiffDirection.SrcToDst;
            var modCells = 0;
            var addRows = 0;
            var remRows = 0;
            var addCols = 0;
            var remCols = 0;

            foreach (var row in currentSheetDiff.Rows.Values)
            {
                if (row.IsAdded()) addRows++;
                else if (row.IsRemoved()) remRows++;

                foreach (var c in row.Cells.Values)
                    if (c.Status == ExcelCellStatus.Modified) modCells++;
            }

            if (currentSheetDiff.ColumnStatusMap != null)
            {
                foreach (var kv in currentSheetDiff.ColumnStatusMap)
                {
                    if (kv.Value == ExcelColumnStatus.Inserted) addCols++;
                    else if (kv.Value == ExcelColumnStatus.Deleted) remCols++;
                }
            }

            var parts = new List<string>();
            if (modCells > 0) parts.Add("值差异 " + modCells + " 格");
            if (addRows > 0) parts.Add((srcToDst ? "删除右侧 " : "插入左侧 ") + addRows + " 行");
            if (remRows > 0) parts.Add((srcToDst ? "插入右侧 " : "删除左侧 ") + remRows + " 行");
            if (addCols > 0) parts.Add((srcToDst ? "删除右侧 " : "插入左侧 ") + addCols + " 列");
            if (remCols > 0) parts.Add((srcToDst ? "插入右侧 " : "删除左侧 ") + remCols + " 列");

            return parts.Count == 0 ? "\n  (无差异)" : "\n  " + string.Join("；", parts);
        }

        private ExcelSheetReadConfig CreateReadConfig()
        {
            var setting = ((App)Application.Current).Setting;

            return new ExcelSheetReadConfig()
            {
                TrimFirstBlankRows = setting.SkipFirstBlankRows,
                TrimFirstBlankColumns = setting.SkipFirstBlankColumns,
                TrimLastBlankRows = setting.TrimLastBlankRows,
                TrimLastBlankColumns = setting.TrimLastBlankColumns,
            };
        }

        private Tuple<ExcelWorkbook, ExcelWorkbook> ReadWorkbooks()
        {
            ExcelWorkbook swb = null;
            ExcelWorkbook dwb = null;
            var srcPath = SrcPathTextBox.Text;
            var dstPath = DstPathTextBox.Text;
            ProgressWindow.DoWorkWithModal(progress =>
            {
                progress.Report(Properties.Resources.Msg_ReadingFiles);

                var config = CreateReadConfig();
                swb = ExcelWorkbook.Create(srcPath, config);
                dwb = ExcelWorkbook.Create(dstPath, config);
            });

            return Tuple.Create(swb, dwb);
        }

        private Tuple<FileSetting, FileSetting> FindFileSettings(bool isStartup)
        {
            FileSetting srcSetting = null;
            FileSetting dstSetting = null;
            var srcPath = SrcPathTextBox.Text;
            var dstPath = DstPathTextBox.Text;
            if (!IgnoreFileSettingCheckbox.IsChecked.Value)
            {
                srcSetting =
                    FindFilseSetting(Path.GetFileName(srcPath), SrcSheetCombobox.SelectedIndex, SrcSheetCombobox.SelectedItem.ToString(), isStartup);

                dstSetting =
                    FindFilseSetting(Path.GetFileName(dstPath), DstSheetCombobox.SelectedIndex, DstSheetCombobox.SelectedItem.ToString(), isStartup);

                diffConfig = CreateDiffConfig(srcSetting, dstSetting, isStartup);
            }
            else
            {
                diffConfig = new ExcelSheetDiffConfig();

                diffConfig.SrcSheetIndex = Math.Max(SrcSheetCombobox.SelectedIndex, 0);
                diffConfig.DstSheetIndex = Math.Max(DstSheetCombobox.SelectedIndex, 0);
            }

            return Tuple.Create(srcSetting, dstSetting);
        }

        private void ExecuteDiff(bool isStartup = false, ExcelSheetDiff specificSheetDiff = null)
        {
            if (!File.Exists(SrcPathTextBox.Text) || !File.Exists(DstPathTextBox.Text))
                return;

            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchPreExecuteDiffEvent(args);

            ExcelSheetDiff diff = specificSheetDiff;

            FileSetting srcFileSetting = null;
            FileSetting dstFileSetting = null;

            if (diff == null)
            {
                var workbooks = ReadWorkbooks();
                var srcWorkbook = workbooks.Item1;
                var dstWorkbook = workbooks.Item2;

                var fileSettings = FindFileSettings(isStartup);
                srcFileSetting = fileSettings.Item1;
                dstFileSetting = fileSettings.Item2;

                SrcSheetCombobox.SelectedIndex = diffConfig.SrcSheetIndex;
                DstSheetCombobox.SelectedIndex = diffConfig.DstSheetIndex;

                // 批量提取：对全部分组的 sheet 对逐个 diff，供右上角 sheet 状态下拉着色
                GetViewModel().SheetDiffInfoList.Clear();
                foreach (var names in GetViewModel().GetSheetNamePairs())
                {
                    ExcelSheet srcSheet;
                    ExcelSheet dstSheet;
                    if (!srcWorkbook.Sheets.TryGetValue(names.Item1, out srcSheet) ||
                        !dstWorkbook.Sheets.TryGetValue(names.Item2, out dstSheet))
                        continue;

                    if (srcSheet.Rows.Count > 10000 || dstSheet.Rows.Count > 10000)
                        MessageBox.Show(Properties.Resources.Msg_WarnSize);

                    GetViewModel().AddSheetDiff(ExcelSheet.Diff(srcSheet, dstSheet, diffConfig));
                }

                var srcName = SrcSheetCombobox.SelectedItem?.ToString();
                var dstName = DstSheetCombobox.SelectedItem?.ToString();
                diff = GetViewModel().SheetDiffInfoList
                    .FirstOrDefault(i => i.SheetDiff.SrcSheet.Name == srcName &&
                                         i.SheetDiff.DstSheet.Name == dstName)?.SheetDiff;

                // 当前对没被按位配对覆盖（如右侧 sheet 数量不同）→ 单独 diff 一次补上
                if (diff == null && srcName != null && dstName != null)
                {
                    ExcelSheet srcSheet;
                    ExcelSheet dstSheet;
                    if (srcWorkbook.Sheets.TryGetValue(srcName, out srcSheet) &&
                        dstWorkbook.Sheets.TryGetValue(dstName, out dstSheet))
                    {
                        diff = ExcelSheet.Diff(srcSheet, dstSheet, diffConfig);
                        GetViewModel().AddSheetDiff(diff);
                    }
                }

                if (diff == null)
                    return;
            }
            else
            {
                var fileSettings = FindFileSettings(isStartup);
                srcFileSetting = fileSettings.Item1;
                dstFileSetting = fileSettings.Item2;

                SrcSheetCombobox.SelectedIndex = diffConfig.SrcSheetIndex;
                DstSheetCombobox.SelectedIndex = diffConfig.DstSheetIndex;
            }

            currentSheetDiff = diff;
            SrcSheetCombobox.SelectedIndex = GetViewModel().GetSrcSheetIndex(diff.SrcSheet.Name);
            DstSheetCombobox.SelectedIndex = GetViewModel().GetDstSheetIndex(diff.DstSheet.Name);
            diffConfig.SrcSheetIndex = SrcSheetCombobox.SelectedIndex;
            diffConfig.DstSheetIndex = DstSheetCombobox.SelectedIndex;

            SrcDataGrid.Model = new DiffGridModel(diff, DiffType.Source);
            DstDataGrid.Model = new DiffGridModel(diff, DiffType.Dest);

            args = new DiffViewEventArgs<FastGridControl>(SrcDataGrid, container);
            DataGridEventDispatcher.Instance.DispatchFileSettingUpdateEvent(args, srcFileSetting);

            args = new DiffViewEventArgs<FastGridControl>(DstDataGrid, container);
            DataGridEventDispatcher.Instance.DispatchFileSettingUpdateEvent(args, dstFileSetting);

            args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, ShowOnlyDiffRadioButton.IsChecked.Value);
            DataGridEventDispatcher.Instance.DispatchPostExecuteDiffEvent(args);

            var summary = diff.CreateSummary();
            GetViewModel().UpdateDiffSummary(summary);

            if (!App.Instance.KeepFileHistory)
                App.Instance.UpdateRecentFiles(SrcPathTextBox.Text, DstPathTextBox.Text);

            if (App.Instance.Setting.NotifyEqual && !summary.HasDiff)
                MessageBox.Show(Properties.Resources.Message_NoDiff);

            if (App.Instance.Setting.FocusFirstDiff)
                MoveNextModifiedCell();
        }

        private FileSetting FindFilseSetting(string fileName, int sheetIndex, string sheetName, bool isStartup)
        {
            var results = new List<FileSetting>();
            foreach (var setting in App.Instance.Setting.FileSettings)
            {
                if (setting.UseRegex)
                {
                    var regex = new System.Text.RegularExpressions.Regex(setting.Name);

                    if (regex.IsMatch(fileName))
                        results.Add(setting);
                }
                else
                {
                    if (setting.ExactMatch)
                    {
                        if (setting.Name == fileName)
                            results.Add(setting);
                    }
                    else
                    {
                        if (fileName.Contains(setting.Name))
                            results.Add(setting);
                    }
                }
            }

            if (isStartup)
                return results.FirstOrDefault(r => r.IsStartupSheet) ?? results.FirstOrDefault() ?? null;

            return results.FirstOrDefault(r => r.SheetName == sheetName) ?? results.FirstOrDefault(r => r.SheetIndex == sheetIndex) ?? null;
        }

        private void SetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderChagneEvent(args);
                }
            }
        }

        private void ResetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderResetEvent(args);
                }
            }
        }

        private void SetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderChangeEvent(args);
                }
            }
        }

        private void ResetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderResetEvent(args);
                }
            }
        }

        private void SwapButton_Click(object sender, RoutedEventArgs e)
        {
            Swap();
        }

        private void Swap()
        {
            var srcTmp = SrcSheetCombobox.SelectedIndex;
            var dstTmp = DstSheetCombobox.SelectedIndex;

            var tmp = SrcPathTextBox.Text;
            SrcPathTextBox.Text = DstPathTextBox.Text;
            DstPathTextBox.Text = tmp;

            diffConfig.SrcSheetIndex = dstTmp;
            diffConfig.DstSheetIndex = srcTmp;

            ExecuteDiff();
        }

        private void PrevSheetButton_Click(object sender, RoutedEventArgs e)
        {
            ShowPrevSheetDiff();
        }

        private void NextSheetButton_Click(object sender, RoutedEventArgs e)
        {
            ShowNextSheetDiff();
        }

        private void SelectedSheetDiffCombobox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var diffInfo = (sender as ComboBox)?.SelectedItem as ExcelSheetDiffInfo;
            if (diffInfo == null)
                return;

            SrcSheetCombobox.SelectedItem = diffInfo.SheetDiff.SrcSheet.Name;
            DstSheetCombobox.SelectedItem = diffInfo.SheetDiff.DstSheet.Name;

            ExecuteDiff(specificSheetDiff: diffInfo.SheetDiff);
        }

        private void ShowNextSheetDiff()
        {
            if (isLock)
                return;

            if (SrcSheetCombobox.SelectedIndex + 1 >= SrcSheetCombobox.Items.Count ||
                DstSheetCombobox.SelectedIndex + 1 >= DstSheetCombobox.Items.Count)
                return;

            isLock = true;

            SrcSheetCombobox.SelectedIndex = SrcSheetCombobox.SelectedIndex + 1;
            diffConfig.SrcSheetIndex = SrcSheetCombobox.SelectedIndex;
            DstSheetCombobox.SelectedIndex = DstSheetCombobox.SelectedIndex + 1;
            diffConfig.DstSheetIndex = DstSheetCombobox.SelectedIndex;

            ExecuteDiff();

            isLock = false;
        }

        private void ShowPrevSheetDiff()
        {
            if (isLock)
                return;

            if (SrcSheetCombobox.SelectedIndex - 1 < 0 ||
                DstSheetCombobox.SelectedIndex - 1 < 0)
                return;

            isLock = true;

            SrcSheetCombobox.SelectedIndex = SrcSheetCombobox.SelectedIndex - 1;
            diffConfig.SrcSheetIndex = SrcSheetCombobox.SelectedIndex;
            DstSheetCombobox.SelectedIndex = DstSheetCombobox.SelectedIndex - 1;
            diffConfig.DstSheetIndex = DstSheetCombobox.SelectedIndex;

            ExecuteDiff();

            isLock = false;
        }

        private void DiffByHeaderSrc_Click(object sender, RoutedEventArgs e)
        {
            var headerIndex = SrcDataGrid.CurrentCell.Row.HasValue ? SrcDataGrid.CurrentCell.Row.Value : -1;

            diffConfig.SrcHeaderIndex= headerIndex;

            ExecuteDiff();
        }

        private void DiffByHeaderDst_Click(object sender, RoutedEventArgs e)
        {
            var headerIndex = DstDataGrid.CurrentCell.Row.HasValue ? DstDataGrid.CurrentCell.Row.Value : -1;

            diffConfig.DstSheetIndex = headerIndex;

            ExecuteDiff();
        }

        private void ShowAllRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, false);
        }

        private void ShowOnlyDiffRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, true);
        }

        private bool ValidateDataGrids()
        {
            return SrcDataGrid.Model != null && DstDataGrid.Model != null;
        }

        private void ValuteTextBox_ScrollChanged(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container);
            ValueTextBoxEventDispatcher.Instance.DispatchScrolledEvent(args, (ScrollChangedEventArgs)e);
        }

        private void NextModifiedCellButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextModifiedCell();
        }

        private void MoveNextModifiedCell()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextModifiedCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevModifiedCellButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevModifiedCell();
        }

        private void MovePrevModifiedCell()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousModifiedCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextModifiedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextModifiedRow();
        }

        private void MoveNextModifiedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextModifiedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevModifiedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevModifiedRow();
        }

        private void MovePrevModifiedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousModifiedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextAddedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextAddedRow();
        }

        private void MoveNextAddedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextAddedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevAddedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevAddedRow();
        }

        private void MovePrevAddedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousAddedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextRemovedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextRemovedRow();
        }

        private void MoveNextRemovedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextRemovedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevRemovedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevRemovedRow();
        }

        private void MovePrevRemovedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousRemovedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevMatchCellButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevMatchCell();
        }

        private void MovePrevMatchCell()
        {
            if (!ValidateDataGrids())
                return;

            var text = SearchTextCombobox.Text;
            if (string.IsNullOrEmpty(text))
                return;

            var history = App.Instance.Setting.SearchHistory.ToList();
            if (history.Contains(text))
                history.Remove(text);

            history.Insert(0, text);
            history = history.Take(10).ToList();

            App.Instance.Setting.SearchHistory = new ObservableCollection<string>(history);
            App.Instance.Setting.Save();

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousMatchCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell, text,
                ExactMatchCheckBox.IsChecked.Value, CaseSensitiveCheckBox.IsChecked.Value, RegexCheckBox.IsChecked.Value, ShowOnlyDiffRadioButton.IsChecked.Value);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextMatchCellButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextMatchCell();
        }

        private void MoveNextMatchCell()
        {
            if (!ValidateDataGrids())
                return;

            var text = SearchTextCombobox.Text;
            if (string.IsNullOrEmpty(text))
                return;

            var history = App.Instance.Setting.SearchHistory.ToList();
            if (history.Contains(text))
                history.Remove(text);

            history.Insert(0, text);
            history = history.Take(10).ToList();

            App.Instance.Setting.SearchHistory = new ObservableCollection<string>(history);
            App.Instance.Setting.Save();

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextMatchCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell, text,
                ExactMatchCheckBox.IsChecked.Value, CaseSensitiveCheckBox.IsChecked.Value, RegexCheckBox.IsChecked.Value, ShowOnlyDiffRadioButton.IsChecked.Value);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void CopyToClipboardSelectedCells(string separator)
        {
            if (copyTargetGrid == null)
                return;

            var model = copyTargetGrid.Model as DiffGridModel;
            if (model == null)
                return;

            var tsv = string.Join(Environment.NewLine,
               copyTargetGrid.SelectedCells
              .GroupBy(c => c.Row.Value)
              .OrderBy(g => g.Key)
              .Select(g => string.Join(separator, g.Select(c => model.GetCellText(c, true)))));

            Clipboard.SetDataObject(tsv);
        }

        private void UserControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Right:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextModifiedCell();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Left:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevModifiedCell();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Down:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextModifiedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Up:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevModifiedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.L:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextRemovedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.O:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevRemovedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.K:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextAddedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.I:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevAddedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.F8:
                    {
                        MovePrevMatchCell();
                        e.Handled = true;
                    }
                    break;
                case Key.F9:
                    {
                        MoveNextMatchCell();
                        e.Handled = true;
                    }
                    break;
                case Key.F:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            ToolExpander.IsExpanded = true;
                            SearchTextCombobox.Focus();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.C:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            CopyToClipboardSelectedCells(Keyboard.IsKeyDown(Key.RightShift) || Keyboard.IsKeyDown(Key.LeftShift) ? "," : "\t");
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.B:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            ShowLog();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.PageDown:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            ShowNextSheetDiff();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.PageUp:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            ShowPrevSheetDiff();
                            e.Handled = true;
                        }
                    }
                    break;
            }
        }

        private void ShowLog()
        {
            var log = BuildCellBaseLog();

            (App.Current.MainWindow as MainWindow).WriteToConsole(log);
        }

        private void BuildCellBaseLog_Click(object sender, RoutedEventArgs e)
        {
            ShowLog();
        }

        private string BuildCellBaseLog()
        {
            var srcModel = SrcDataGrid.Model as DiffGridModel;
            if (srcModel == null)
                return string.Empty;

            var dstModel = DstDataGrid.Model as DiffGridModel;
            if (dstModel == null)
                return string.Empty;

            var builder = new StringBuilder();

            var selectedCells = SrcDataGrid.SelectedCells;

            var modifiedLogFormat = App.Instance.Setting.LogFormat;
            var addedLogFormat = App.Instance.Setting.AddedRowLogFormat;
            var removedLogFormat = App.Instance.Setting.RemovedRowLogFormat;

            foreach (var row in SrcDataGrid.SelectedCells.GroupBy(c => c.Row))
            {
                var rowHeaderText = srcModel.GetRowHeaderText(row.Key.Value);
                if (string.IsNullOrEmpty(rowHeaderText))
                    rowHeaderText = dstModel.GetRowHeaderText(row.Key.Value);

                if (dstModel.IsAddedRow(row.Key.Value, true))
                {
                    var log = addedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText));

                    builder.AppendLine(log);

                    continue;
                }

                if (dstModel.IsRemovedRow(row.Key.Value, true))
                {
                    var log = removedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText));

                    builder.AppendLine(log);

                    continue;
                }

                foreach (var cell in row)
                {
                    if (cell.Row.Value == srcModel.ColumnHeaderIndex)
                        continue;

                    var srcText = srcModel.GetCellText(cell, true);
                    var dstText = dstModel.GetCellText(cell, true);
                    if (srcText == dstText)
                        continue;

                    var colHeaderText = srcModel.GetColumnHeaderText(cell.Column.Value);

                    if (string.IsNullOrEmpty(colHeaderText))
                        colHeaderText = dstModel.GetColumnHeaderText(cell.Column.Value);

                    if (string.IsNullOrEmpty(srcText))
                        srcText = Properties.Resources.Word_Blank;

                    if (string.IsNullOrEmpty(dstText))
                        dstText = Properties.Resources.Word_Blank;

                    if (string.IsNullOrEmpty(rowHeaderText))
                        rowHeaderText = Properties.Resources.Word_Blank;

                    if (string.IsNullOrEmpty(colHeaderText))
                        colHeaderText = Properties.Resources.Word_Blank;

                    var log = modifiedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText))
                        .Replace("${COL}", RemoveMultiLine(colHeaderText))
                        .Replace("${LEFT}", RemoveMultiLine(srcText))
                        .Replace("${RIGHT}", RemoveMultiLine(dstText));

                    builder.AppendLine(log);
                }
            }

            return builder.ToString();
        }

        private string RemoveMultiLine(string log)
        {
            return log.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        }

        private void CopyAsTsv_Click(object sender, RoutedEventArgs e)
        {
            CopyToClipboardSelectedCells("\t");
        }

        private void CopyAsCsv_Click(object sender, RoutedEventArgs e)
        {
            CopyToClipboardSelectedCells(",");
        }
    }
}
