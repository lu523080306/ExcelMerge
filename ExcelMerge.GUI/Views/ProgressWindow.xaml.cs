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

            window.Loaded += (_, args) =>
            {
                var worker = new BackgroundWorker();
                var progress = new Progress<string>(data => window.Message.Content = data);

                worker.DoWork += (s, workerArgs) => work(progress);
                worker.RunWorkerCompleted += (s, workerArgs) =>
                {
                    // BackgroundWorker 不会向外抛异常，异常被吞到 workerArgs.Error。
                    // 必须原样重抛（保留原始堆栈），否则 ReadWorkbooks 等调用方
                    // 拿到 null 对象，最终弹的是误导性的 NullReferenceException。
                    if (workerArgs.Error != null)
                        System.Runtime.ExceptionServices.ExceptionDispatchInfo
                            .Capture(workerArgs.Error).Throw();

                    window.Close();
                };
                worker.RunWorkerAsync();
            };

            window.ShowDialog();
        }
    }
}
