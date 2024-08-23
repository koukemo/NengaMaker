using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using OfficeOpenXml;
using System.IO;
using Microsoft.Win32;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace NengaMaker
{
    public partial class MainWindow : Window
    {
        private string recipientName;
        private string recipientFurigana;
        private string recipientPostalCode;
        private string recipientAddress1;
        private string recipientAddress2;
        private string excelFilePath;
        private int rowCount;

        private string templeteImagePath = "file:///C://Users/javas/Programs/NengaMaker/NengaMaker/figures/postcard_1.png";

        /// <summary>
        /// Initializes the MainWindow instance and sets up the background image.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            SetBackgroundImage();
        }

        /// <summary>
        /// Sets the background image for the preview canvas.
        /// </summary>
        private void SetBackgroundImage()
        {
            ImageBrush brush = new ImageBrush();
            brush.ImageSource = new BitmapImage(new Uri(templeteImagePath));
            brush.Stretch = Stretch.Uniform;
            brush.ViewportUnits = BrushMappingMode.Absolute;
            brush.Viewport = new Rect(0, 0, 394, 583);
            previewCanvas.Background = brush;
        }

        /// <summary>
        /// Handles the Print button click event. Initiates the printing process.
        /// </summary>
        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Excelファイルを読み込んでください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            Canvas printCanvas = CreatePrintCanvas();
            string tempPngPath = System.IO.Path.GetTempFileName() + ".png";
            SaveCanvasAsImage(printCanvas, tempPngPath);
            ShowPrintDialog(tempPngPath, null, true);
        }

        /// <summary>
        /// Handles the Export Image button click event. Exports the canvas as a PNG image.
        /// </summary>
        private void ExportImageButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Excelファイルを読み込んでください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PNG Files|*.png";
            if (saveFileDialog.ShowDialog() == true)
            {
                Canvas imageCanvas = CreatePrintCanvas();
                SaveCanvasAsImage(imageCanvas, saveFileDialog.FileName);
            }
        }

        /// <summary>
        /// Handles the Print All button click event. Prints all addresses from the Excel file.
        /// </summary>
        private void PrintAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Excelファイルを読み込んでください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                bool isPrintToPdf = printDialog.PrintQueue.Name == "Microsoft Print to PDF";
                string baseFilePath = null;

                if (isPrintToPdf)
                {
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "PDF Files|*.pdf";
                    if (saveFileDialog.ShowDialog() == true)
                    {
                        baseFilePath = saveFileDialog.FileName;
                    }
                    else
                    {
                        return;
                    }
                }

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        recipientName = worksheet.Cells[row, 2].Text;
                        recipientFurigana = worksheet.Cells[row, 3].Text;
                        recipientPostalCode = worksheet.Cells[row, 4].Text;
                        recipientAddress1 = worksheet.Cells[row, 5].Text;
                        recipientAddress2 = worksheet.Cells[row, 6].Text;

                        Canvas printCanvas = CreatePrintCanvas();
                        string tempPngPath = System.IO.Path.GetTempFileName() + ".png";
                        SaveCanvasAsImage(printCanvas, tempPngPath);

                        if (isPrintToPdf)
                        {
                            string directory = Path.GetDirectoryName(baseFilePath);
                            string baseFileName = Path.GetFileNameWithoutExtension(baseFilePath);
                            string pdfFilePath = Path.Combine(directory, $"{baseFileName}_{row - 1}.pdf");
                            ConvertPngToPdf(tempPngPath, pdfFilePath, false);
                        }
                        else
                        {
                            printDialog.PrintVisual(printCanvas, "Nenga Print");
                        }
                    }
                }

                if (isPrintToPdf)
                {
                    MessageBox.Show("すべてのPDFが保存されました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        /// <summary>
        /// Creates a Canvas for printing, with specified dimensions.
        /// </summary>
        /// <returns>A fully prepared Canvas object for printing.</returns>
        private Canvas CreatePrintCanvas()
        {
            Canvas printCanvas = new Canvas
            {
                Width = 1181,
                Height = 1748,
                Background = System.Windows.Media.Brushes.White
            };

            AddTextBlocksToCanvas(printCanvas, 1.0);
            return printCanvas;
        }

        /// <summary>
        /// Adds text blocks to the provided canvas based on the current recipient's information.
        /// </summary>
        /// <param name="canvas">The canvas to add the text blocks to.</param>
        /// <param name="scale">The scale factor for the text blocks.</param>
        private void AddTextBlocksToCanvas(Canvas canvas, double scale)
        {
            double fontSize = 60 * scale;
            double recipientPostalCodeX = 550 * scale;
            double recipientPostalCodeY = 140 * scale;
            double recipientAddressFontSize = 70 * scale;
            double recipientAddressX = 1000 * scale;
            double recipientAddressY = 350 * scale;
            double recipientAddressX2 = recipientAddressX - (80 * scale);
            double recipientAddressY2 = recipientAddressY + (50 * scale);
            double recipientNameFontSize = 90 * scale;
            double recipientNameX = recipientAddressX2 - (200 * scale);
            double recipientNameY = recipientAddressY2 + (100 * scale);

            double senderPostCodeFontSize = 40 * scale;
            double senderAddressFontSize = 50 * scale;
            double senderNameFontSize = 50 * scale;
            double senderPostalCodeX = 80 * scale;
            double senderPostalCodeY = 1450 * scale;
            double senderAddressX = 300 * scale;
            double senderAddressY = 700 * scale;
            double senderAddressX2 = senderAddressX - (80 * scale);
            double senderAddressY2 = senderAddressY;
            double senderNameX = senderAddressX2 - (100 * scale);
            double senderNameY = senderAddressY2 + (100 * scale);

            AddPostalCodeToCanvas(canvas, recipientPostalCode, recipientPostalCodeX, recipientPostalCodeY, fontSize, scale);
            AddVerticalTextToCanvas(canvas, recipientAddress1, recipientAddressFontSize, recipientAddressX, recipientAddressY);
            AddVerticalTextToCanvas(canvas, recipientAddress2, recipientAddressFontSize, recipientAddressX2, recipientAddressY2);
            AddVerticalTextToCanvas(canvas, recipientName + " 様", recipientNameFontSize, recipientNameX, recipientNameY);

            string senderPostalCode = SenderPostalCodeTextBox.Text;
            string senderAddress1 = SenderAddress1TextBox.Text;
            string senderAddress2 = SenderAddress2TextBox.Text;
            string senderName = SenderNameTextBox.Text;

            AddSenderPostalCodeToCanvas(canvas, senderPostalCode, senderPostalCodeX, senderPostalCodeY, senderPostCodeFontSize, scale);
            AddVerticalTextToCanvas(canvas, senderAddress1, senderAddressFontSize, senderAddressX, senderAddressY);
            AddVerticalTextToCanvas(canvas, senderAddress2, senderAddressFontSize, senderAddressX2, senderAddressY2);
            AddVerticalTextToCanvas(canvas, senderName, senderNameFontSize, senderNameX, senderNameY);
        }

        /// <summary>
        /// Adds the postal code to the provided canvas at the specified location.
        /// </summary>
        /// <param name="canvas">The canvas to add the postal code to.</param>
        /// <param name="postalCode">The postal code to add.</param>
        /// <param name="startX">The starting X coordinate.</param>
        /// <param name="startY">The starting Y coordinate.</param>
        /// <param name="fontSize">The font size to use for the postal code.</param>
        /// <param name="scale">The scale factor for positioning and font size.</param>
        private void AddPostalCodeToCanvas(Canvas canvas, string postalCode, double startX, double startY, double fontSize, double scale)
        {
            postalCode = postalCode.Replace("-", "");
            double spacing = 80 * scale;

            for (int i = 0; i < postalCode.Length; i++)
            {
                TextBlock textBlock = new TextBlock
                {
                    Text = postalCode[i].ToString(),
                    FontSize = fontSize,
                    Margin = new Thickness(0),
                    FontWeight = FontWeights.Bold
                };
                canvas.Children.Add(textBlock);
                Canvas.SetLeft(textBlock, startX + i * spacing);
                Canvas.SetTop(textBlock, startY);
            }
        }

        /// <summary>
        /// Adds the sender's postal code to the provided canvas at the specified location.
        /// </summary>
        /// <param name="canvas">The canvas to add the postal code to.</param>
        /// <param name="postalCode">The postal code to add.</param>
        /// <param name="startX">The starting X coordinate.</param>
        /// <param name="startY">The starting Y coordinate.</param>
        /// <param name="fontSize">The font size to use for the postal code.</param>
        /// <param name="scale">The scale factor for positioning and font size.</param>
        private void AddSenderPostalCodeToCanvas(Canvas canvas, string postalCode, double startX, double startY, double fontSize, double scale)
        {
            postalCode = postalCode.Replace("-", "");
            double spacing = 50 * scale;

            for (int i = 0; i < postalCode.Length; i++)
            {
                TextBlock textBlock = new TextBlock
                {
                    Text = postalCode[i].ToString(),
                    FontSize = fontSize,
                    Margin = new Thickness(0),
                    FontWeight = FontWeights.Bold
                };
                canvas.Children.Add(textBlock);
                Canvas.SetLeft(textBlock, startX + i * spacing);
                Canvas.SetTop(textBlock, startY);
            }
        }

        /// <summary>
        /// Adds vertical text to the provided canvas.
        /// </summary>
        /// <param name="canvas">The canvas to add the text to.</param>
        /// <param name="text">The text to add.</param>
        /// <param name="fontSize">The font size to use for the text.</param>
        /// <param name="x">The starting X coordinate.</param>
        /// <param name="y">The starting Y coordinate.</param>
        private void AddVerticalTextToCanvas(Canvas canvas, string text, double fontSize, double x, double y)
        {
            for (int i = 0; i < text.Length; i++)
            {
                TextBlock textBlock = new TextBlock
                {
                    Text = text[i].ToString(),
                    FontSize = fontSize,
                    Margin = new Thickness(0),
                    FontWeight = FontWeights.Bold
                };

                if (text[i] == 'ー' || text[i] == '-')
                {
                    textBlock.RenderTransform = new RotateTransform(90);
                    textBlock.RenderTransformOrigin = new Point(0.5, 0.5);
                }

                canvas.Children.Add(textBlock);
                Canvas.SetLeft(textBlock, x);
                Canvas.SetTop(textBlock, y + i * fontSize);
            }
        }

        /// <summary>
        /// Handles the Load From Excel button click event. Loads recipient data from the selected Excel file.
        /// </summary>
        private void LoadFromExcelButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                excelFilePath = openFileDialog.FileName;

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    recipientName = worksheet.Cells[2, 2].Text;
                    recipientFurigana = worksheet.Cells[2, 3].Text;
                    recipientPostalCode = worksheet.Cells[2, 4].Text;
                    recipientAddress1 = worksheet.Cells[2, 5].Text;
                    recipientAddress2 = worksheet.Cells[2, 6].Text;

                    UpdatePreview();
                }
            }
        }

        /// <summary>
        /// Handles the Update Preview button click event. Updates the preview canvas with the current recipient data.
        /// </summary>
        private void UpdatePreviewButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Excelファイルを読み込んでください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            UpdatePreview();
        }

        /// <summary>
        /// Updates the preview canvas by clearing it and adding the current recipient's data.
        /// </summary>
        private void UpdatePreview()
        {
            previewCanvas.Children.Clear();
            AddTextBlocksToCanvas(previewCanvas, 1.0 / 3.0);
        }

        /// <summary>
        /// Handles the GotFocus event for the text boxes. Clears placeholder text when focused.
        /// </summary>
        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null && (textBox.Text == "送り主の名前" || textBox.Text == "送り主の郵便番号" || textBox.Text == "送り主の住所1" || textBox.Text == "送り主の住所2"))
            {
                textBox.Text = "";
                textBox.Foreground = new SolidColorBrush(Colors.Black);
            }
        }

        /// <summary>
        /// Handles the LostFocus event for the text boxes. Restores placeholder text when the field is empty.
        /// </summary>
        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null && string.IsNullOrWhiteSpace(textBox.Text))
            {
                textBox.Text = textBox.Name == "SenderNameTextBox" ? "送り主の名前" :
                               textBox.Name == "SenderPostalCodeTextBox" ? "送り主の郵便番号" :
                               textBox.Name == "SenderAddress1TextBox" ? "送り主の住所1" :
                               textBox.Name == "SenderAddress2TextBox" ? "送り主の住所2" : "";
                textBox.Foreground = new SolidColorBrush(Colors.Gray);
            }
        }

        /// <summary>
        /// Saves the provided canvas as a PNG image to the specified file path.
        /// </summary>
        /// <param name="canvas">The canvas to save as an image.</param>
        /// <param name="filePath">The file path to save the image to.</param>
        private void SaveCanvasAsImage(Canvas canvas, string filePath)
        {
            RenderTargetBitmap renderBitmap = new RenderTargetBitmap(
                (int)canvas.Width, (int)canvas.Height,
                96d, 96d, PixelFormats.Pbgra32);
            canvas.Measure(new Size((int)canvas.Width, (int)canvas.Height));
            canvas.Arrange(new Rect(new Size((int)canvas.Width, (int)canvas.Height)));
            renderBitmap.Render(canvas);

            using (FileStream outStream = new FileStream(filePath, FileMode.Create))
            {
                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(renderBitmap));
                encoder.Save(outStream);
            }
        }

        /// <summary>
        /// Converts a PNG image to a PDF document and saves it to the specified file path.
        /// </summary>
        /// <param name="pngPath">The file path of the PNG image to convert.</param>
        /// <param name="pdfPath">The file path to save the PDF to. If null, prompts the user to select a location.</param>
        /// <param name="showMessage">Whether to show a message after saving the PDF.</param>
        private void ConvertPngToPdf(string pngPath, string pdfPath = null, bool showMessage = true)
        {
            if (pdfPath == null)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "PDF Files|*.pdf";
                if (saveFileDialog.ShowDialog() == true)
                {
                    pdfPath = saveFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }

            PdfSharp.Pdf.PdfDocument document = new PdfSharp.Pdf.PdfDocument();
            PdfSharp.Pdf.PdfPage page = document.AddPage();
            PdfSharp.Drawing.XGraphics gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page);

            using (FileStream fs = new FileStream(pngPath, FileMode.Open, FileAccess.Read))
            {
                PdfSharp.Drawing.XImage image = PdfSharp.Drawing.XImage.FromStream(fs);
                gfx.DrawImage(image, 0, 0, page.Width, page.Height);
            }

            document.Save(pdfPath);

            if (showMessage)
            {
                MessageBox.Show($"PDFが保存されました: {pdfPath}", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            File.Delete(pngPath);
        }

        /// <summary>
        /// Displays a print dialog and handles the print operation, converting the canvas to PDF if necessary.
        /// </summary>
        /// <param name="tempPngPath">The temporary PNG file path for printing.</param>
        /// <param name="pdfPath">The file path to save the PDF to, if printing to PDF.</param>
        /// <param name="showMessage">Whether to show a message after saving the PDF.</param>
        private void ShowPrintDialog(string tempPngPath, string pdfPath = null, bool showMessage = true)
        {
            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                if (printDialog.PrintQueue.Name == "Microsoft Print to PDF")
                {
                    ConvertPngToPdf(tempPngPath, pdfPath, showMessage);
                }
                else
                {
                    Canvas printCanvas = CreatePrintCanvas();
                    SaveCanvasAsImage(printCanvas, tempPngPath);
                    printDialog.PrintVisual(printCanvas, "Nenga Print");
                }
            }
        }
    }
}
