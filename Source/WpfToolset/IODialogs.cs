using System;
using System.Collections.Generic;
using System.Text;

using System.Windows.Controls;
using Microsoft.Win32;

namespace WpfToolset
{
    public class IODialogs
    {
        public static bool TrySelectFile(out string filePath, string title, string defaultExt)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                CheckFileExists = true,
                CheckPathExists = true,

                Multiselect = false,

                DefaultExt = defaultExt,
                DereferenceLinks = true,
                Title = title
            };

            var result = openFileDialog.ShowDialog();

            if (result.HasValue && result.Value)
            {
                filePath = openFileDialog.FileName;
                return true;
            }
            else
            {
                filePath = null;
                return false;
            }
        }

        public static bool TrySelectFile(TextBox textBox, string title, string defaultExt)
        {
            if (TrySelectFile(out string path, title, defaultExt))
            {
                textBox.Clear();
                textBox.AppendText(path);

                // ScrollToEnd isn't working here, I don't know why. But scrolling to a very big number
                // does it.
                textBox.ScrollToHorizontalOffset(10000);

                return true;
            }

            return false;
        }
    }
}
