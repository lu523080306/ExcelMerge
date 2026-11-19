using System.Windows;

namespace ExcelMerge.GUI.Views
{
    /// <summary>
    /// 轻量提示窗：固定宽度、按内容定高、可拖动、按钮始终可见。
    /// 用于替代"把长路径/堆栈塞进 MessageBox 导致撑满全屏"的场景。
    /// </summary>
    public partial class MessageDialog : Window
    {
        public MessageDialog()
        {
            InitializeComponent();
        }

        /// <summary>重试 / 取消。返回 true 表示重试。</summary>
        public static bool ShowRetry(Window owner, string title, string message, string detail)
        {
            var dialog = Create(owner, title, message, detail, "重试", "取消");
            return dialog.ShowDialog() == true;
        }

        /// <summary>只有"确定"按钮。</summary>
        public static void ShowInfo(Window owner, string title, string message, string detail)
        {
            var dialog = Create(owner, title, message, detail, "确定", null);
            dialog.ShowDialog();
        }

        /// <summary>自定义两个按钮。返回 true 表示点了主按钮。</summary>
        public static bool ShowYesNo(Window owner, string title, string message, string detail,
            string primaryText, string secondaryText)
        {
            var dialog = Create(owner, title, message, detail, primaryText, secondaryText);
            return dialog.ShowDialog() == true;
        }

        private static MessageDialog Create(Window owner, string title, string message, string detail,
            string primaryText, string secondaryText)
        {
            var dialog = new MessageDialog();
            dialog.Title = title;
            dialog.MessageText.Text = message ?? string.Empty;
            dialog.PrimaryButton.Content = primaryText;

            if (owner != null && owner.IsLoaded)
                dialog.Owner = owner;

            if (string.IsNullOrEmpty(detail))
            {
                dialog.DetailExpander.Visibility = Visibility.Collapsed;
            }
            else
            {
                dialog.DetailText.Text = detail;
            }

            if (string.IsNullOrEmpty(secondaryText))
            {
                dialog.CancelButton.Visibility = Visibility.Collapsed;
            }
            else
            {
                dialog.CancelButton.Content = secondaryText;
            }

            return dialog;
        }

        private void PrimaryButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
