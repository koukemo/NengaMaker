using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using OfficeOpenXml; // EPPlusの名前空間
using System.IO;
using Microsoft.Win32; // 保存ダイアログのための名前空間
using PdfSharp.Pdf; // PdfSharpの名前空間
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

        public MainWindow()
        {
            InitializeComponent();
            // ライセンスコンテキストの設定
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // プレビューを更新（背景画像を設定）
            SetBackgroundImage();
        }

        private void SetBackgroundImage()
        {
            ImageBrush brush = new ImageBrush();
            brush.ImageSource = new BitmapImage(new Uri("file:///C://Users/javas/Programs/NengaMaker/NengaMaker/figures/postcard_1.png"));
            brush.Stretch = Stretch.Uniform;
            brush.ViewportUnits = BrushMappingMode.Absolute;
            brush.Viewport = new Rect(0, 0, 394, 583); // 1/3スケールのサイズ
            previewCanvas.Background = brush;
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            Canvas printCanvas = CreatePrintCanvas(); // フルサイズの印刷用キャンバスを作成
            string tempPngPath = System.IO.Path.GetTempFileName() + ".png";
            SaveCanvasAsImage(printCanvas, tempPngPath);
            ConvertPngToPdf(tempPngPath);
        }

        private void ExportImageButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PNG Files|*.png";
            if (saveFileDialog.ShowDialog() == true)
            {
                Canvas imageCanvas = CreatePrintCanvas(); // フルサイズのキャンバスを作成
                SaveCanvasAsImage(imageCanvas, saveFileDialog.FileName);
            }
        }

        private Canvas CreatePrintCanvas()
        {
            Canvas printCanvas = new Canvas
            {
                Width = 1181,
                Height = 1748,
                Background = System.Windows.Media.Brushes.White
            };

            AddTextBlocksToCanvas(printCanvas, 1.0); // フルサイズでテキスト要素を追加
            return printCanvas;
        }

        private void AddTextBlocksToCanvas(Canvas canvas, double scale)
        {
            // 文字単位の設定
            double fontSize = 60 * scale; // スケールに応じたフォントサイズ
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

            // sender settings
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

            // 受取人郵便番号（赤枠に収めるための位置設定）
            AddPostalCodeToCanvas(canvas, recipientPostalCode, recipientPostalCodeX, recipientPostalCodeY, fontSize, scale);

            // 受取人住所と氏名（縦書き）
            AddVerticalTextToCanvas(canvas, recipientAddress1, recipientAddressFontSize, recipientAddressX, recipientAddressY);
            AddVerticalTextToCanvas(canvas, recipientAddress2, recipientAddressFontSize, recipientAddressX2, recipientAddressY2);
            AddVerticalTextToCanvas(canvas, recipientName + " 様", recipientNameFontSize, recipientNameX, recipientNameY);

            // 差出人郵便番号（赤枠に収めるための位置設定）
            string senderPostalCode = SenderPostalCodeTextBox.Text;
            string senderAddress1 = SenderAddress1TextBox.Text;
            string senderAddress2 = SenderAddress2TextBox.Text;
            string senderName = SenderNameTextBox.Text;

            AddSenderPostalCodeToCanvas(canvas, senderPostalCode, senderPostalCodeX, senderPostalCodeY, senderPostCodeFontSize, scale);

            // 差出人住所と氏名（縦書き）
            AddVerticalTextToCanvas(canvas, senderAddress1, senderAddressFontSize, senderAddressX, senderAddressY);
            AddVerticalTextToCanvas(canvas, senderAddress2, senderAddressFontSize, senderAddressX2, senderAddressY2);
            AddVerticalTextToCanvas(canvas, senderName, senderNameFontSize, senderNameX, senderNameY);
        }

        private void AddPostalCodeToCanvas(Canvas canvas, string postalCode, double startX, double startY, double fontSize, double scale)
        {
            postalCode = postalCode.Replace("-", ""); // ハイフンを削除
            double spacing = 80 * scale; // 各数字の間隔

            for (int i = 0; i < postalCode.Length; i++)
            {
                TextBlock textBlock = new TextBlock
                {
                    Text = postalCode[i].ToString(),
                    FontSize = fontSize, // 指定されたフォントサイズ
                    Margin = new Thickness(0),
                    FontWeight = FontWeights.Bold
                };
                canvas.Children.Add(textBlock);
                Canvas.SetLeft(textBlock, startX + i * spacing);
                Canvas.SetTop(textBlock, startY);
            }
        }

        private void AddSenderPostalCodeToCanvas(Canvas canvas, string postalCode, double startX, double startY, double fontSize, double scale)
        {
            postalCode = postalCode.Replace("-", ""); // ハイフンを削除
            double spacing = 50 * scale; // 各数字の間隔

            for (int i = 0; i < postalCode.Length; i++)
            {
                TextBlock textBlock = new TextBlock
                {
                    Text = postalCode[i].ToString(),
                    FontSize = fontSize, // 指定されたフォントサイズ
                    Margin = new Thickness(0),
                    FontWeight = FontWeights.Bold
                };
                canvas.Children.Add(textBlock);
                Canvas.SetLeft(textBlock, startX + i * spacing);
                Canvas.SetTop(textBlock, startY);
            }
        }

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
                    textBlock.RenderTransformOrigin = new Point(0.5, 0.5); // 回転の中心を設定
                }

                canvas.Children.Add(textBlock);
                Canvas.SetLeft(textBlock, x);
                Canvas.SetTop(textBlock, y + i * fontSize); // 各文字を縦に配置
            }
        }

        private void LoadFromExcelButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // 2行目からデータを読み込む（見出し行を飛ばす）
                    recipientName = worksheet.Cells[2, 2].Text;
                    recipientFurigana = worksheet.Cells[2, 3].Text;
                    recipientPostalCode = worksheet.Cells[2, 4].Text;
                    recipientAddress1 = worksheet.Cells[2, 5].Text;
                    recipientAddress2 = worksheet.Cells[2, 6].Text;

                    // プレビューを更新
                    UpdatePreview();
                }
            }
        }

        private void UpdatePreviewButton_Click(object sender, RoutedEventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            previewCanvas.Children.Clear();
            AddTextBlocksToCanvas(previewCanvas, 1.0 / 3.0); // 1/3スケールでテキスト要素を追加
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null && (textBox.Text == "送り主の名前" || textBox.Text == "送り主の郵便番号" || textBox.Text == "送り主の住所1" || textBox.Text == "送り主の住所2"))
            {
                textBox.Text = "";
                textBox.Foreground = new SolidColorBrush(Colors.Black);
            }
        }

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

        private void ConvertPngToPdf(string pngPath)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PDF Files|*.pdf";
            if (saveFileDialog.ShowDialog() == true)
            {
                string pdfPath = saveFileDialog.FileName;

                PdfSharp.Pdf.PdfDocument document = new PdfSharp.Pdf.PdfDocument();
                PdfSharp.Pdf.PdfPage page = document.AddPage();
                PdfSharp.Drawing.XGraphics gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page);

                using (FileStream fs = new FileStream(pngPath, FileMode.Open, FileAccess.Read))
                {
                    PdfSharp.Drawing.XImage image = PdfSharp.Drawing.XImage.FromStream(fs);
                    gfx.DrawImage(image, 0, 0, page.Width, page.Height);
                }

                document.Save(pdfPath);
                MessageBox.Show($"PDFが保存されました: {pdfPath}", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);

                // 一時PNGファイルを削除
                File.Delete(pngPath);
            }
        }
    }
}
