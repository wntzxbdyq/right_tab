namespace UnifiedContextMenu.App.WinForms;

internal static class PromptDialog
{
    public static string? Show(IWin32Window owner, string title, string defaultValue)
    {
        using var form = new Form
        {
            Width = 520,
            Height = 160,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            Text = title,
            StartPosition = FormStartPosition.CenterParent,
            MinimizeBox = false,
            MaximizeBox = false
        };

        var textBox = new TextBox
        {
            Left = 16,
            Top = 18,
            Width = 470,
            Text = defaultValue
        };
        var okButton = new Button
        {
            Text = "确定",
            Left = 316,
            Width = 80,
            Top = 58,
            DialogResult = DialogResult.OK
        };
        var cancelButton = new Button
        {
            Text = "取消",
            Left = 406,
            Width = 80,
            Top = 58,
            DialogResult = DialogResult.Cancel
        };

        form.Controls.Add(textBox);
        form.Controls.Add(okButton);
        form.Controls.Add(cancelButton);
        form.AcceptButton = okButton;
        form.CancelButton = cancelButton;

        var result = form.ShowDialog(owner);
        return result == DialogResult.OK ? textBox.Text.Trim() : null;
    }
}
