using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.IO;
using System.Windows;
using Prism.Mvvm;
using FastWpfGrid;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Behaviors;

namespace ExcelMerge.GUI.ViewModels
{
    public class DiffViewModel : BindableBase
    {
        private bool showLocationGridLine;
        public bool ShowLocationGridLine
        {
            get { return showLocationGridLine; }
            set { SetProperty(ref showLocationGridLine, value); }
        }

        private string srcPath;
        public string SrcPath
        {
            get { return srcPath; }
            set
            {
                SetProperty(ref srcPath, value);
                Settings.EMEnvironmentValue.Set("SRC", value);
                UpdateExecutableFlag();
            }
        }

        private string dstPath;
        public string DstPath
        {
            get { return dstPath; }
            set
            {
                SetProperty(ref dstPath, value);
                Settings.EMEnvironmentValue.Set("DST", value);
                UpdateExecutableFlag();
            }
        }

        private List<string> srcSheetNames;
        public List<string> SrcSheetNames
        {
            get { return srcSheetNames; }
            private set { SetProperty(ref srcSheetNames, value); }
        }

        private List<string> dstSheetNames;
        public List<string> DstSheetNames
        {
            get { return dstSheetNames; }
            private set { SetProperty(ref dstSheetNames, value); }
        }

        private int selectedSrcSheetIndex;
        public int SelectedSrcSheetIndex
        {
            get { return selectedSrcSheetIndex; }
            set
            {
                SetProperty(ref selectedSrcSheetIndex, value);
                UpdateOtherSheetsExecutableFlag();
            }
        }

        private int selectedDstSheetIndex;
        public int SelectedDstSheetIndex
        {
            get { return selectedDstSheetIndex; }
            set
            {
                SetProperty(ref selectedDstSheetIndex, value);
                UpdateOtherSheetsExecutableFlag();
            }
        }

        private bool executable;
        public bool Executable
        {
            get { return executable; }
            private set
            {
                SetProperty(ref executable, value);
                UpdateOtherSheetsExecutableFlag();
            }
        }

        private bool executableNext;
        public bool ExecutableNext
        {
            get { return executableNext; }
            private set { SetProperty(ref executableNext, value); }
        }

        private bool executablePrev;
        public bool ExecutablePrev
        {
            get { return executablePrev; }
            private set { SetProperty(ref executablePrev, value); }
        }

        private int modifiedCellCount;
        public int ModifiedCellCount
        {
            get { return modifiedCellCount; }
            private set { SetProperty(ref modifiedCellCount, value); }
        }

        private int modifiedRowCount;
        public int ModifiedRowCount
        {
            get { return modifiedRowCount; }
            private set { SetProperty(ref modifiedRowCount, value); }
        }

        private int addedRowCount;
        public int AddedRowCount
        {
            get { return addedRowCount; }
            private set { SetProperty(ref addedRowCount, value); }
        }

        private int removedRowCount;
        public int RemovedRowCount
        {
            get { return removedRowCount; }
            private set { SetProperty(ref removedRowCount, value); }
        }

        private DragAcceptDescription description;
        public DragAcceptDescription Description
        {
            get { return description; }
            private set { SetProperty(ref description, value); }
        }

        private ObservableCollection<ExcelSheetDiffInfo> sheetDiffInfoList =
            new ObservableCollection<ExcelSheetDiffInfo>();
        public ObservableCollection<ExcelSheetDiffInfo> SheetDiffInfoList
        {
            get { return sheetDiffInfoList; }
            private set { SetProperty(ref sheetDiffInfoList, value); }
        }

        public DiffViewModel()
        {
            Description = new DragAcceptDescription();
            Description.DragDrop += DragDrop;
            Description.DragDrop += DragOver;

            SrcPath = string.Empty;
            DstPath = string.Empty;
        }

        public DiffViewModel(string src, string dst, MainWindowViewModel mwv) : this()
        {
            SrcPath = src;
            DstPath = dst;

            mwv.PropertyChanged += Mwv_PropertyChanged;
        }

        public void UpdateDiffSummary(ExcelSheetDiffSummary summary)
        {
            ModifiedCellCount = summary.ModifiedCellCount;
            ModifiedRowCount = summary.ModifiedRowCount;
            AddedRowCount = summary.AddedRowCount;
            RemovedRowCount = summary.RemovedRowCount;
        }

        public void AddSheetDiff(ExcelSheetDiff diff)
        {
            if (SheetDiffInfoList
                .Any(i => i.SheetDiff.SrcSheet.Name == diff.SrcSheet.Name &&
                          i.SheetDiff.DstSheet.Name == diff.DstSheet.Name))
                return;

            SheetDiffInfoList.Add(new ExcelSheetDiffInfo(diff));
        }

        public int GetSrcSheetIndex(string name)
        {
            return SrcSheetNames?.IndexOf(name) ?? -1;
        }

        public int GetDstSheetIndex(string name)
        {
            return DstSheetNames?.IndexOf(name) ?? -1;
        }

        public IEnumerable<Tuple<string, string>> GetSheetNamePairs()
        {
            if (SrcSheetNames == null || DstSheetNames == null)
                return Enumerable.Empty<Tuple<string, string>>();

            return SrcSheetNames.Zip(DstSheetNames, (s, d) => Tuple.Create(s, d));
        }

        private void UpdateOtherSheetsExecutableFlag()
        {
            var srcCount = srcSheetNames != null ? srcSheetNames.Count : 0;
            var dstCount = dstSheetNames != null ? dstSheetNames.Count : 0;

            ExecutableNext = Executable &&
                SelectedSrcSheetIndex + 1 < srcCount &&
                SelectedDstSheetIndex + 1 < dstCount;

            ExecutablePrev = Executable &&
                SelectedSrcSheetIndex - 1 >= 0 &&
                SelectedDstSheetIndex - 1 >= 0;
        }

        private void Mwv_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SrcPath))
            {
                var vm = sender as MainWindowViewModel;
                if (vm != null)
                {
                    var prop = typeof(MainWindowViewModel).GetProperties().FirstOrDefault(p => p.Name == e.PropertyName);
                    if (prop != null)
                    {
                        SrcPath = prop.GetValue(vm) as string;
                    }
                }
            }
            else if (e.PropertyName == nameof(DstPath))
            {
                var vm = sender as MainWindowViewModel;
                if (vm != null)
                {
                    var prop = typeof(MainWindowViewModel).GetProperties().FirstOrDefault(p => p.Name == e.PropertyName);
                    if (prop != null)
                    {
                        DstPath = prop.GetValue(vm) as string;
                    }
                }
            }
        }

        private void DragDrop(DragEventArgs e)
        {
            var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (paths == null || !paths.Any())
                return;

            var target = e.Source as FrameworkElement;
            if (target == null)
                return;

            OnDragDrop(paths, target);
        }

        protected virtual void OnDragDrop(string[] filePath, FrameworkElement target)
        {
            if (filePath.Length > 1)
            {
                SrcPath = filePath[1];
                DstPath = filePath[0];

                return;
            }

            var tag = Convert.ToInt32(target.Tag);
            if (tag == 0)
                SrcPath = filePath[0];
            else if (tag == 1)
                DstPath = filePath[0];
        }

        private void DragOver(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

            e.Handled = true;
        }

        /// <summary>
        /// 由 ReadWorkbooks 成功后回填 sheet 列表。
        /// 启动时文件被占用会导致构造阶段的 sheet 探测失败、列表为空。
        /// </summary>
        public void EnsureSheetNames(IEnumerable<string> srcNames, IEnumerable<string> dstNames)
        {
            if (SrcSheetNames == null || SrcSheetNames.Count == 0)
            {
                SrcSheetNames = srcNames.ToList();
                if (SelectedSrcSheetIndex < 0 && SrcSheetNames.Count > 0)
                    SelectedSrcSheetIndex = 0;
            }

            if (DstSheetNames == null || DstSheetNames.Count == 0)
            {
                DstSheetNames = dstNames.ToList();
                if (SelectedDstSheetIndex < 0 && DstSheetNames.Count > 0)
                    SelectedDstSheetIndex = 0;
            }
        }

        private void UpdateExecutableFlag()
        {
            var existsSrc = File.Exists(SrcPath);
            var existsDst = File.Exists(DstPath);

            if (existsSrc)
            {
                try
                {
                    SrcSheetNames = ExcelWorkbook.GetSheetNames(SrcPath).ToList();
                }
                catch (Exception)
                {
                    // 文件可能正被 WPS/Excel 占用。这里不能抛：此调用发生在主窗口
                    // Show() 之前，异常会让程序直接退出且没有任何提示。
                    // 真正的读取失败由 ReadWorkbooks 的“文件被占用/重试”弹窗处理。
                    SrcSheetNames = new List<string>();
                }

                SelectedSrcSheetIndex = SrcSheetNames.Count > 0 ? 0 : -1;
            }
            else
            {
                SrcSheetNames = new List<string>();
                SelectedSrcSheetIndex = -1;
            }

            if (existsDst)
            {
                try
                {
                    DstSheetNames = ExcelWorkbook.GetSheetNames(DstPath).ToList();
                }
                catch (Exception)
                {
                    DstSheetNames = new List<string>();
                }

                SelectedDstSheetIndex = DstSheetNames.Count > 0 ? 0 : -1;
            }
            else
            {
                DstSheetNames = new List<string>();
                SelectedDstSheetIndex = -1;
            }

            Executable = existsSrc && existsDst;
        }
    }

    public class ExcelSheetDiffInfo
    {
        public ExcelSheetDiff SheetDiff { get; }
        public bool IsNotified { get; set; }

        public ExcelSheetDiffInfo(ExcelSheetDiff sheetDiff)
        {
            SheetDiff = sheetDiff;
        }

        public override string ToString()
        {
            return SheetDiff?.ToString();
        }
    }
}
