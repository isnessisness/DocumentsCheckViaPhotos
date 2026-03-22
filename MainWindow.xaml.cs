using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using OpenCvSharp;
using Sdcb.PaddleInference;
using Sdcb.PaddleOCR;
using Sdcb.PaddleOCR.Models.Local;

namespace WpfOcrApp
{
    public partial class MainWindow : System.Windows.Window
    {
        private PaddleOcrAll? _ocrEngine;

        public MainWindow()
        {
            InitializeComponent();
            _ = InitOcrAsync();
        }

        private async Task InitOcrAsync()
        {
            await Task.Run(() =>
            {
                _ocrEngine = new PaddleOcrAll(LocalFullModels.ChineseV3, PaddleDevice.Blas())
                {
                    AllowRotateDetection = false,
                };

                if (_ocrEngine.Detector != null)
                {
                    _ocrEngine.Detector.MaxSize = null;
                }

                Dispatcher.Invoke(() => Log("✅ OCR 模型加载完成，高精度网格解析引擎已就绪。"));
            });
        }

        private void BtnSelectInput_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog { Title = "选择存放待提取图片的文件夹" };
            if (dialog.ShowDialog() == true) TxtInputFolder.Text = dialog.FolderName;
        }

        private void BtnSelectOutput_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog { Title = "选择整合后的结果存放文件夹" };
            if (dialog.ShowDialog() == true) TxtOutputFolder.Text = dialog.FolderName;
        }

        private async void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TxtInputFolder.Text) || string.IsNullOrEmpty(TxtOutputFolder.Text))
            {
                MessageBox.Show("请先选择图片输入和结果输出的文件夹！");
                return;
            }

            if (_ocrEngine == null) return;

            string inputDir = TxtInputFolder.Text;
            string outputDir = TxtOutputFolder.Text;
            int sortMode = CmbSort.SelectedIndex;

            BtnStart.IsEnabled = false;
            PbProgress.Value = 0;
            TxtProgress.Text = "0 %";

            Log("🚀 开始启动物理网格解析与 Excel 完美无乱码导出...");
            await Task.Run(() => ProcessBatchCore(inputDir, outputDir, sortMode));
            Dispatcher.Invoke(() => BtnStart.IsEnabled = true);
        }

        private void ProcessBatchCore(string inputDir, string outputDir, int sortMode)
        {
            try
            {
                var extensions = new[] { ".jpg", ".jpeg", ".png", ".bmp" };
                var dirInfo = new DirectoryInfo(inputDir);
                var files = dirInfo.GetFiles("*.*")
                    .Where(f => extensions.Contains(f.Extension.ToLower()))
                    .ToList();

                if (files.Count == 0) return;

                if (sortMode == 0) files = files.OrderBy(f => f.Name).ToList();
                else files = files.OrderBy(f => f.LastWriteTime).ToList();

                string timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                string finalTxtPath = Path.Combine(outputDir, $"表格排版_{timeStamp}.txt");
                string finalCsvPath = Path.Combine(outputDir, $"Excel分列数据_{timeStamp}.csv");

                using var swTxt = new StreamWriter(finalTxtPath, false, Encoding.UTF8);
                // 🔥 必须使用 new UTF8Encoding(true) 写入 BOM 头，且绝对不再写入 sep=, 防止触发 Excel 乱码 Bug
                using var swCsv = new StreamWriter(finalCsvPath, false, new UTF8Encoding(true));

                for (int i = 0; i < files.Count; i++)
                {
                    var file = files[i];
                    Log($"⏳ 正在解析网格 ({i + 1}/{files.Count}): {file.Name}");

                    ExtractAndAppend(file.FullName, file.Name, i + 1, swTxt, swCsv);

                    swTxt.WriteLine();
                    swCsv.WriteLine();

                    int currentProgress = (int)((i + 1) * 100.0 / files.Count);
                    Dispatcher.Invoke(() =>
                    {
                        PbProgress.Value = currentProgress;
                        TxtProgress.Text = $"{currentProgress} %";
                    });
                }
                Log($"🎉 批量处理结束！请打开 CSV 电子表格，中文已完美显示且精准分列！");
            }
            catch (Exception ex)
            {
                Log($"❌ 批量处理发生错误: {ex.Message}");
            }
        }

        private void ExtractAndAppend(string imagePath, string fileName, int index, StreamWriter swTxt, StreamWriter swCsv)
        {
            try
            {
                byte[] fileBytes = File.ReadAllBytes(imagePath);
                using Mat src = Cv2.ImDecode(fileBytes, ImreadModes.Color);
                if (src.Empty()) return;

                PaddleOcrResult result;
                lock (_ocrEngine!)
                {
                    result = _ocrEngine.Run(src);
                }

                if (result.Regions.Length == 0)
                {
                    swTxt.WriteLine($"\r\n========== 【图 {index}：{fileName}】 ==========\r\n");
                    // 智取方案：填充 10 个空逗号，骗过 Excel 让它乖乖开启逗号分列模式
                    swCsv.WriteLine($"\"========== 【图 {index}：{fileName}】 ==========\"{new string(',', 10)}");
                    swTxt.WriteLine("（未检测到任何文字）");
                    swCsv.WriteLine("\"（未检测到任何文字）\"");
                    return;
                }

                string[,] table = ReconstructByPhysicalBorders(src, result.Regions);
                int rows = table.GetLength(0);
                int cols = table.GetLength(1);

                swTxt.WriteLine($"\r\n========== 【图 {index}：{fileName}】 ==========\r\n");

                // 🔥 绝杀：根据表格实际列数，在标题行末尾“隐身”填充等量逗号。
                // 完美激活 Excel 逗号嗅探器，彻底解决挤在第一列的问题，同时原汁原味保留 UTF-8 编码！
                string dummyCommas = cols > 1 ? new string(',', cols - 1) : "";
                swCsv.WriteLine($"\"========== 【图 {index}：{fileName}】 ==========\"{dummyCommas}");

                int[] colWidths = new int[cols];
                for (int j = 0; j < cols; j++)
                {
                    for (int i = 0; i < rows; i++)
                    {
                        int w = GetDisplayWidth(table[i, j] ?? "");
                        if (w > colWidths[j]) colWidths[j] = w;
                    }
                }

                for (int i = 0; i < rows; i++)
                {
                    bool emptyRow = true;
                    for (int j = 0; j < cols; j++)
                    {
                        if (!string.IsNullOrEmpty(table[i, j])) { emptyRow = false; break; }
                    }
                    if (emptyRow) continue;

                    StringBuilder txtLine = new StringBuilder();
                    List<string> csvLine = new List<string>();

                    for (int j = 0; j < cols; j++)
                    {
                        string text = table[i, j] ?? "";

                        txtLine.Append(PadRightEx(text, colWidths[j]));
                        if (j < cols - 1) txtLine.Append("    ");

                        // 安全转义内部可能的双引号，避免 CSV 格式错乱
                        csvLine.Add($"\"{text.Replace("\"", "\"\"")}\"");
                    }

                    swTxt.WriteLine(txtLine.ToString());
                    swCsv.WriteLine(string.Join(",", csvLine));
                }
            }
            catch (Exception ex)
            {
                Log($"❌ 单图处理失败 ({fileName}): {ex.Message}");
            }
        }

        // =======================================================
        // 核心优化区：彻底解决吞列与细线丢失问题的终极排版引擎
        // =======================================================

        private class TextBlock
        {
            public string Text { get; set; }
            public float Left { get; set; }
            public float Top { get; set; }
            public float Right { get; set; }
            public float Bottom { get; set; }
            public float Width => Right - Left;
            public float Height => Bottom - Top;
            public float CenterX => Left + Width / 2f;
            public float CenterY => Top + Height / 2f;

            public TextBlock(PaddleOcrResultRegion r)
            {
                Text = r.Text;
                Left = r.Rect.Center.X - r.Rect.Size.Width / 2f;
                Right = r.Rect.Center.X + r.Rect.Size.Width / 2f;
                Top = r.Rect.Center.Y - r.Rect.Size.Height / 2f;
                Bottom = r.Rect.Center.Y + r.Rect.Size.Height / 2f;
            }
        }

        private void GetGridBoundaries(Mat src, out List<int> colEdges, out List<int> rowEdges)
        {
            using Mat gray = new Mat();
            if (src.Channels() >= 3) Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);
            else src.CopyTo(gray);

            using Mat binary = new Mat();
            Cv2.AdaptiveThreshold(gray, binary, 255, AdaptiveThresholdTypes.GaussianC, ThresholdTypes.BinaryInv, 15, 5);

            using Mat kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(2, 2));
            Cv2.MorphologyEx(binary, binary, MorphTypes.Dilate, kernel);

            using Mat vertical = new Mat();
            int vSize = Math.Max(5, src.Rows / 50);
            using Mat vStruct = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(1, vSize));
            Cv2.MorphologyEx(binary, vertical, MorphTypes.Open, vStruct);

            using Mat horizontal = new Mat();
            int hSize = Math.Max(5, src.Cols / 50);
            using Mat hStruct = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(hSize, 1));
            Cv2.MorphologyEx(binary, horizontal, MorphTypes.Open, hStruct);

            colEdges = GetLinePositions(vertical, true, src.Rows);
            rowEdges = GetLinePositions(horizontal, false, src.Cols);

            if (colEdges.Count == 0 || colEdges[0] > 20) colEdges.Insert(0, 0);
            if (colEdges.Last() < src.Cols - 20) colEdges.Add(src.Cols);
            if (rowEdges.Count == 0 || rowEdges[0] > 20) rowEdges.Insert(0, 0);
            if (rowEdges.Last() < src.Rows - 20) rowEdges.Add(src.Rows);
        }

        private List<int> GetLinePositions(Mat linesMat, bool isVertical, int lengthRef)
        {
            Cv2.FindContours(linesMat, out OpenCvSharp.Point[][] contours, out OpenCvSharp.HierarchyIndex[] hierarchy, RetrievalModes.External, ContourApproximationModes.ApproxSimple);
            List<int> positions = new List<int>();

            foreach (OpenCvSharp.Point[] contour in contours)
            {
                var rect = Cv2.BoundingRect(contour);
                if (isVertical && rect.Height > lengthRef / 30) positions.Add(rect.X + rect.Width / 2);
                else if (!isVertical && rect.Width > lengthRef / 30) positions.Add(rect.Y + rect.Height / 2);
            }

            positions.Sort();
            List<int> merged = new List<int>();
            if (positions.Count > 0)
            {
                int current = positions[0];
                List<int> group = new List<int> { current };
                for (int i = 1; i < positions.Count; i++)
                {
                    if (positions[i] - current < 20)
                        group.Add(positions[i]);
                    else
                    {
                        merged.Add((int)group.Average());
                        group.Clear();
                        group.Add(positions[i]);
                    }
                    current = positions[i];
                }
                if (group.Count > 0) merged.Add((int)group.Average());
            }
            return merged;
        }

        private string[,] ReconstructByPhysicalBorders(Mat src, PaddleOcrResultRegion[] regions)
        {
            var blocks = regions.Select(r => new TextBlock(r)).ToList();
            GetGridBoundaries(src, out List<int> colEdges, out List<int> rowEdges);

            if (colEdges.Count < 3 || rowEdges.Count < 3)
                return FallbackLayout(blocks);

            int rows = rowEdges.Count - 1;
            int cols = colEdges.Count - 1;

            string[,] table = new string[rows, cols];
            var cellBlocks = new Dictionary<(int, int), List<TextBlock>>();

            foreach (var b in blocks)
            {
                int rIdx = -1, cIdx = -1;
                for (int i = 0; i < rows; i++)
                    if (b.CenterY >= rowEdges[i] && b.CenterY <= rowEdges[i + 1]) { rIdx = i; break; }
                for (int j = 0; j < cols; j++)
                    if (b.CenterX >= colEdges[j] && b.CenterX <= colEdges[j + 1]) { cIdx = j; break; }

                if (rIdx == -1) rIdx = b.CenterY < rowEdges[0] ? 0 : rows - 1;
                if (cIdx == -1) cIdx = b.CenterX < colEdges[0] ? 0 : cols - 1;

                var key = (rIdx, cIdx);
                if (!cellBlocks.ContainsKey(key)) cellBlocks[key] = new List<TextBlock>();
                cellBlocks[key].Add(b);
            }

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (cellBlocks.ContainsKey((i, j)))
                    {
                        var cellData = cellBlocks[(i, j)].OrderBy(b => b.CenterY).ThenBy(b => b.Left).ToList();
                        table[i, j] = string.Join(" ", cellData.Select(b => b.Text));
                    }
                }
            }
            return table;
        }

        private string[,] FallbackLayout(List<TextBlock> blocks)
        {
            if (blocks.Count == 0) return new string[0, 0];

            var validWidths = blocks.Select(b => b.Width).OrderBy(w => w).ToList();
            float medianWidth = validWidths.Count > 0 ? validWidths[validWidths.Count / 2] : 0;

            var narrowBlocks = blocks.Where(b => b.Width < medianWidth * 3).OrderBy(b => b.CenterX).ToList();
            if (narrowBlocks.Count == 0) narrowBlocks = blocks.OrderBy(b => b.CenterX).ToList();

            var colCenters = new List<float>();
            foreach (var b in narrowBlocks)
            {
                if (!colCenters.Any(c => Math.Abs(c - b.CenterX) < b.Height * 1.5f))
                {
                    colCenters.Add(b.CenterX);
                }
            }
            colCenters.Sort();
            int cols = colCenters.Count;

            var rowCenters = new List<float>();
            foreach (var b in blocks.OrderBy(b => b.CenterY))
            {
                if (!rowCenters.Any(r => Math.Abs(r - b.CenterY) < b.Height * 0.6f))
                {
                    rowCenters.Add(b.CenterY);
                }
            }
            rowCenters.Sort();
            int rows = rowCenters.Count;

            string[,] table = new string[rows, cols];
            foreach (var b in blocks)
            {
                int rIdx = -1; float minRDiff = float.MaxValue;
                for (int r = 0; r < rows; r++)
                {
                    float diff = Math.Abs(b.CenterY - rowCenters[r]);
                    if (diff < minRDiff) { minRDiff = diff; rIdx = r; }
                }

                int cIdx = -1; float minCDiff = float.MaxValue;
                for (int c = 0; c < cols; c++)
                {
                    float diff = Math.Abs(b.CenterX - colCenters[c]);
                    if (diff < minCDiff) { minCDiff = diff; cIdx = c; }
                }

                if (rIdx != -1 && cIdx != -1)
                {
                    if (string.IsNullOrEmpty(table[rIdx, cIdx])) table[rIdx, cIdx] = b.Text;
                    else table[rIdx, cIdx] += " " + b.Text;
                }
            }
            return table;
        }

        private int GetDisplayWidth(string str)
        {
            if (string.IsNullOrEmpty(str)) return 0;
            int width = 0;
            foreach (char c in str)
            {
                if ((c >= 0x4E00 && c <= 0x9FA5) || (c >= 0xFF01 && c <= 0xFF5E) || c > 255) width += 2;
                else width += 1;
            }
            return width;
        }

        private string PadRightEx(string str, int totalWidth)
        {
            if (str == null) str = "";
            int padding = totalWidth - GetDisplayWidth(str);
            if (padding > 0) return str + new string(' ', padding);
            return str;
        }

        private void Log(string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => Log(message)));
                return;
            }

            LstLog.Items.Insert(0, $"[{DateTime.Now:HH:mm:ss}] {message}");
            if (LstLog.Items.Count > 100) LstLog.Items.RemoveAt(100);
        }

        protected override void OnClosed(EventArgs e)
        {
            _ocrEngine?.Dispose();
            base.OnClosed(e);
        }
    }
}