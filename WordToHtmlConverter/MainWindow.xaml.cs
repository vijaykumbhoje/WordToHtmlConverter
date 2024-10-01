using Microsoft.Win32;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xceed.Words.NET;

namespace WordToHtmlConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _filePath;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnUpload_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Documents (*.docx)|*.docx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _filePath = openFileDialog.FileName;
                txtFilePath.Text = _filePath;
                btnConvert.IsEnabled = true;
            }
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_filePath))
            {
                MessageBox.Show("Please upload a Word file first.");
                return;
            }

            string htmlPath = System.IO.Path.ChangeExtension(_filePath, ".html");
            ConvertDocxToHtml(_filePath, htmlPath);
            MessageBox.Show($"HTML saved at: {htmlPath}");
        }

        private void ConvertDocxToHtml(string docxPath, string htmlPath)
        {
            using (DocX document = DocX.Load(docxPath))
            {
                StringBuilder htmlBuilder = new StringBuilder();
                htmlBuilder.Append("<html><head><title>Document</title></head><body>");

                // Iterate through paragraphs
                foreach (var paragraph in document.Paragraphs)
                {
                    // Handle headers
                    if (paragraph.StyleName.StartsWith("Heading"))
                    {
                        htmlBuilder.Append($"<h{paragraph.StyleName.Last()}>{paragraph.Text}</h{paragraph.StyleName.Last()}>");
                    }
                    else
                    {
                        htmlBuilder.Append($"<p>{paragraph.Text}</p>");
                    }
                }

                // Handle images
                foreach (var image in document.Images)
                {
                    using (var imageStream = image.GetStream(FileMode.Open, FileAccess.Read))
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            imageStream.CopyTo(memoryStream);
                            string base64Image = Convert.ToBase64String(memoryStream.ToArray());
                            htmlBuilder.Append($"<img src='data:image/png;base64,{base64Image}' alt='{image.FileName}'/>");
                        }
                    }
                }

                htmlBuilder.Append("</body></html>");

                // Write HTML to file
                File.WriteAllText(htmlPath, htmlBuilder.ToString());
            }
        }




    }
}