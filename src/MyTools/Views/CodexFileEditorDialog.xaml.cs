using System.Text;
using System.Windows;

namespace MyTools.Views
{
    public partial class CodexFileEditorDialog : Window
    {
        /// <summary>编辑后保存的最终明文内容（仅在 DialogResult=true 时有效）。</summary>
        public string EditedText { get; private set; }

        public CodexFileEditorDialog(string fileName, string profileName, string initialText)
        {
            InitializeComponent();
            HeaderTitle.Text = fileName;
            HeaderSub.Text = string.IsNullOrWhiteSpace(profileName) ? string.Empty : "· " + profileName;
            EditorBox.Text = initialText ?? string.Empty;
            Loaded += (s, e) => { EditorBox.Focus(); EditorBox.CaretIndex = 0; };
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            EditedText = EditorBox.Text ?? string.Empty;
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
