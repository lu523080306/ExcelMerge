using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelMerge.GUI.Views
{
    /// <summary>
    /// ProgressWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ProgressWindow : Window
    {
        public ProgressWindow()
        {
            InitializeComponent();
        }

        public static void DoWorkWithModal(Action<IProgress<string>> work)
        {
            var window = new ProgressWindow();
            Exception error = null;

            window.Loaded += (_, args) =>
            {
                var worker = new BackgroundWorker();
                var progress = new Progress<string>(data => window.Message.Content = data);

                // 在后台线程内捕获异常，避免它冒泡到 Dispatcher/AppDomain 级未处理异常；
                // 这样调用方可以在 UI 线程上弹出“文件被占用”之类的友好重试框。
                worker.DoWork += (s, workerArgs) =>
                {
                    try
                    {
                        work(progress);
                    }
                    catch (Exception ex)
                    {
                        error = ex;
                    }
                };
                worker.RunWorkerCompleted += (s, workerArgs) =>
                {
                    window.Close();
                };
                worker.RunWorkerAsync();
            };

            window.ShowDialog();

            // 原样抛出（保留原始堆栈），交给调用方处理
            if (error != null)
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }
    }
}
