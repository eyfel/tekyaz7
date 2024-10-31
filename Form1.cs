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
                MessageBox.Show("Please select a SLDASM file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(assemblyPath))
            {
                MessageBox.Show($"File not found: {assemblyPath}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                rtbLogs.Clear();
                AddLog("Starting conversion process...");
                
                string outputFolder = CreateOutputFolder(Path.GetDirectoryName(assemblyPath));
                AddLog($"Output folder created: {outputFolder}");
                
                var (uniquePartPaths, totalPartCount) = GetPartPathsFromAssembly(assemblyPath);
                AddLog($"Found {uniquePartPaths.Count} unique parts with {totalPartCount} total instances.");
                
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
                        AddLog($"Processing: {Path.GetFileName(partPath)}");
                        RunSolidWorksOperations(partPath, outputFolder);
                        processedParts.Add(partPath);
                        Application.DoEvents();
                    }
                }

                AddLog("Conversion process completed.");
                MessageBox.Show("Conversion process completed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AddLog($"Error occurred: {ex.Message}");
                MessageBox.Show($"Error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                rtbLogs.AppendText(message + Environment.NewLine);
                rtbLogs.ScrollToCaret();
            }
        }

        private string SelectAssemblyFile()
        {
            using OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "SolidWorks Assembly Files (*.SLDASM)|*.SLDASM";
            openFileDialog.Title = "Select SLDASM File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
            return string.Empty;
        }

        private string CreateOutputFolder(string? baseFolder)
        {
            string outputFolder = Path.Combine(baseFolder ?? "", "CONVERTED_DRAWINGS");
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
                    throw new Exception("SolidWorks application not found.");
                }

                swApp = Activator.CreateInstance(swType) as SldWorks;
                if (swApp == null)
                {
                    throw new Exception("Failed to start SolidWorks application.");
                }

                swApp.Visible = false;

                int errors = 0;
                int warnings = 0;
                swModel = swApp.OpenDoc6(assemblyPath, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                
                if (swModel == null)
                {
                    throw new Exception($"Failed to open assembly. Errors: {errors}, Warnings: {warnings}. Error description: {GetErrorDescription(errors)}");
                }

                AssemblyDoc? swAssembly = swModel as AssemblyDoc;
                if (swAssembly == null)
                {
                    throw new Exception("Model is not an assembly file.");
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
                    throw new Exception("SolidWorks application not found.");
                }

                swApp = Activator.CreateInstance(swType) as SldWorks;
                if (swApp == null)
                {
                    throw new Exception("Failed to start SolidWorks application.");
                }

                swApp.Visible = true;

                int errors = 0;
                int warnings = 0;
                swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                
                if (swModel == null)
                {
                    throw new Exception($"Failed to open model. Errors: {errors}, Warnings: {warnings}. Error description: {GetErrorDescription(errors)}");
                }

                PartDoc? swPart = swModel as PartDoc;
                if (swPart == null)
                {
                    throw new Exception("Model is not a part file.");
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
                if (swApp != null)
                {
                    swApp.ExitApp();
                }
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
                throw new Exception("Failed to save sheet metal part.");
            }
        }

        private void ExportPartToDWG(PartDoc swPart, ModelDoc2 swModel, string filePath, string outputFolder)
        {
            object[] bodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if (bodies == null || bodies.Length == 0)
            {
                throw new Exception("No solid bodies found in part.");
            }

            IBody2 body = (IBody2)bodies[0];
            double[] boundingBox = body.GetBodyBox();

            if (boundingBox == null || boundingBox.Length != 6)
            {
                throw new Exception("Failed to get bounding box.");
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
                throw new Exception("Failed to save normal part.");
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
                case 0: return "No error";
                case 1: return "General error";
                case 2: return "File not found or access denied";
                case 3: return "Invalid file type";
                case 4: return "Failed to open file";
                default: return $"Unknown error code: {errorCode}";
            }
        }
    }
}
