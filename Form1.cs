using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace tekyaz7
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            string assemblyPath = SelectAssemblyFile();
            if (!string.IsNullOrWhiteSpace(assemblyPath))
            {
                txtFilePath.Text = assemblyPath;
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            string assemblyPath = txtFilePath.Text;

            if (string.IsNullOrWhiteSpace(assemblyPath))
            {
                MessageBox.Show("Lütfen bir SLDASM dosyası seçin.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(assemblyPath))
            {
                MessageBox.Show($"Dosya bulunamadı: {assemblyPath}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                rtbLogs.Clear();
                AddLog("Dönüştürme işlemi başlatılıyor...");

                string outputFolder = CreateOutputFolder(Path.GetDirectoryName(assemblyPath));
                AddLog($"Çıktı klasörü oluşturuldu: {outputFolder}");

                var (uniquePartPaths, totalPartCount) = GetPartPathsFromAssembly(assemblyPath);
                AddLog($"Toplam {uniquePartPaths.Count} farklı parça {totalPartCount} adet bulundu.");

                progressBar.Minimum = 0;
                progressBar.Maximum = uniquePartPaths.Count;
                progressBar.Value = 0;

                HashSet<string> processedParts = new HashSet<string>();

                foreach (string partPath in uniquePartPaths)
                {
                    if (!processedParts.Contains(partPath))
                    {
                        progressBar.Value++;
                        progressBar.Update();
                        AddLog($"İşleniyor: {Path.GetFileName(partPath)}");
                        RunSolidWorksOperations(partPath, outputFolder);
                        processedParts.Add(partPath);
                        Application.DoEvents();
                    }
                }

                AddLog("Dönüştürme işlemi tamamlandı.");
                MessageBox.Show("Dönüştürme işlemi tamamlandı.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AddLog($"Hata oluştu: {ex.Message}");
                MessageBox.Show($"Hata oluştu: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Value = 0;
            }
        }

        private void AddLog(string message)
        {
            if (rtbLogs.InvokeRequired)
            {
                rtbLogs.Invoke(new Action<string>(AddLog), message);
            }
            else
            {
                rtbLogs.AppendText(message + System.Environment.NewLine);
                rtbLogs.ScrollToCaret();
            }
        }

        private string SelectAssemblyFile()
        {
            using OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "SolidWorks Assembly Files (*.SLDASM)|*.SLDASM";
            openFileDialog.Title = "SLDASM Dosyası Seçin";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
            return string.Empty;
        }

        private string CreateOutputFolder(string? baseFolder)
        {
            string outputFolder = Path.Combine(baseFolder ?? "", "ÖNDER ÇALIŞMASI");
            Directory.CreateDirectory(outputFolder);
            return outputFolder;
        }

        private (List<string> uniqueParts, int totalCount) GetPartPathsFromAssembly(string assemblyPath)
        {
            List<string> uniquePartPaths = new List<string>();
            Dictionary<string, int> partCounts = new Dictionary<string, int>();
            SldWorks? swApp = null;
            ModelDoc2? swModel = null;

            try
            {
                Type? swType = Type.GetTypeFromProgID("SldWorks.Application");
                if (swType == null)
                {
                    throw new Exception("SolidWorks uygulaması bulunamadı.");
                }

                swApp = Activator.CreateInstance(swType) as SldWorks;
                if (swApp == null)
                {
                    throw new Exception("SolidWorks uygulaması başlatılamadı.");
                }

                swApp.Visible = false;

                int errors = 0;
                int warnings = 0;
                swModel = swApp.OpenDoc6(assemblyPath, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

                if (swModel == null)
                {
                    throw new Exception($"Montaj açılamadı. Hatalar: {errors}, Uyarılar: {warnings}. Hata açıklaması: {GetErrorDescription(errors)}");
                }

                AssemblyDoc? swAssembly = swModel as AssemblyDoc;
                if (swAssembly == null)
                {
                    throw new Exception("Model bir montaj dosyası değil.");
                }

                object[] components = swAssembly.GetComponents(false);
                foreach (object comp in components)
                {
                    Component2? component = comp as Component2;
                    if (component != null && !component.IsSuppressed())
                    {
                        string partPath = component.GetPathName();
                        if (!string.IsNullOrEmpty(partPath) && File.Exists(partPath) && partPath.EndsWith(".SLDPRT", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!uniquePartPaths.Contains(partPath))
                            {
                                uniquePartPaths.Add(partPath);
                            }
                            if (partCounts.ContainsKey(partPath))
                            {
                                partCounts[partPath]++;
                            }
                            else
                            {
                                partCounts[partPath] = 1;
                            }
                        }
                    }
                }
            }
            finally
            {
                if (swModel != null && swApp != null)
                {
                    swApp.CloseDoc(swModel.GetTitle());
                }
                if (swApp != null)
                {
                    swApp.ExitApp();
                }
            }

            int totalCount = partCounts.Values.Sum();
            return (uniquePartPaths, totalCount);
        }

        private void RunSolidWorksOperations(string filePath, string outputFolder)
        {
            SldWorks? swApp = null;
            ModelDoc2? swModel = null;
            try
            {
                Type? swType = Type.GetTypeFromProgID("SldWorks.Application");
                if (swType == null)
                {
                    throw new Exception("SolidWorks uygulaması bulunamadı.");
                }

                swApp = Activator.CreateInstance(swType) as SldWorks;
                if (swApp == null)
                {
                    throw new Exception("SolidWorks uygulaması başlatılamadı.");
                }

                swApp.Visible = true;

                int errors = 0;
                int warnings = 0;
                swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

                if (swModel == null)
                {
                    throw new Exception($"Model açılamadı. Hatalar: {errors}, Uyarılar: {warnings}. Hata açıklaması: {GetErrorDescription(errors)}");
                }

                PartDoc? swPart = swModel as PartDoc;
                if (swPart == null)
                {
                    throw new Exception("Model bir parça dosyası değil.");
                }

                bool isSheetMetal = HasSheetMetalFeature(swPart);

                if (isSheetMetal)
                {
                    ExportSheetMetalToDWG(swPart, filePath, outputFolder);
                }
                else
                {
                    ExportPartToDWG(swPart, swModel, filePath, outputFolder);
                }
            }
            finally
            {
                if (swModel != null && swApp != null)
                {
                    swApp.CloseDoc(swModel.GetTitle());
                }
                // if (swApp != null)
                // {
                //     swApp.ExitApp();
                // }
            }
        }

        private void ExportSheetMetalToDWG(PartDoc swPart, string filePath, string outputFolder)
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string sPathName = Path.Combine(outputFolder, $"{fileName}.dwg");
            double[] dataAlignment = new double[12] { 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1 };
            object varAlignment = dataAlignment;

            int sheetMetalOptions = 1;
            bool result = swPart.ExportToDWG2(sPathName, filePath, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal,
                true, varAlignment, false, false, sheetMetalOptions, null);

            if (!result)
            {
                throw new Exception("Sheet metal parçası kaydedilemedi.");
            }
        }

        private void ExportPartToDWG(PartDoc swPart, ModelDoc2 swModel, string filePath, string outputFolder)
        {
            object[] bodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if (bodies == null || bodies.Length == 0)
            {
                throw new Exception("Parçada katı gövde bulunamadı.");
            }

            IBody2 body = (IBody2)bodies[0];
            double[] boundingBox = body.GetBodyBox();

            if (boundingBox == null || boundingBox.Length != 6)
            {
                throw new Exception("Bounding box alınamadı.");
            }

            double xDimension = Math.Abs(boundingBox[3] - boundingBox[0]);
            double yDimension = Math.Abs(boundingBox[4] - boundingBox[1]);
            double zDimension = Math.Abs(boundingBox[5] - boundingBox[2]);

            string viewName;

            if (xDimension <= yDimension && xDimension <= zDimension)
            {
                viewName = "*Right";
            }
            else if (yDimension <= xDimension && yDimension <= zDimension)
            {
                viewName = "*Top";
            }
            else
            {
                viewName = "*Front";
            }

            swModel.ShowNamedView2(viewName, (int)swStandardViews_e.swFrontView);
            swModel.ViewZoomtofit();

            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string sPathName = Path.Combine(outputFolder, $"{fileName}.dwg");
            double[] dataAlignment = new double[12] { 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0 };
            object varAlignment = dataAlignment;

            string[] dataViews = new string[1] { viewName };
            object varViews = dataViews;

            bool result = swPart.ExportToDWG2(sPathName, filePath, (int)swExportToDWG_e.swExportToDWG_ExportAnnotationViews,
                                              true, varAlignment, false, false, 0, varViews);

            if (!result)
            {
                throw new Exception("Normal parça kaydedilemedi.");
            }
        }

        private bool HasSheetMetalFeature(PartDoc part)
        {
            Feature? feat = part.FirstFeature() as Feature;
            while (feat != null)
            {
                if (feat.GetTypeName2() == "SheetMetal")
                {
                    return true;
                }
                feat = feat.GetNextFeature() as Feature;
            }
            return false;
        }

        private string GetErrorDescription(int errorCode)
        {
            switch (errorCode)
            {
                case 0: return "Hata yok";
                case 1: return "Genel hata";
                case 2: return "Dosya bulunamadı veya erişim reddedildi";
                case 3: return "Geçersiz dosya türü";
                case 4: return "Dosya açılamadı";
                default: return $"Bilinmeyen hata kodu: {errorCode}";
            }
        }
    }
}