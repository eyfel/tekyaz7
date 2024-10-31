# SolidWorks Drawing Converter

A C# application that automatically exports SolidWorks assembly components to DWG files with optimized view orientations based on bounding box analysis.

## 🚀 Features

- **Automated DWG Export**: Processes SLDASM files and extracts all unique components
- **Intelligent View Orientation**: Automatically determines the optimal view orientation for each part using bounding box analysis
- **Sheet Metal Support**: Automatically detects and flattens sheet metal components before export
- **Efficient Processing**: Handles duplicate components by processing each unique part only once
- **User-Friendly Interface**: Windows Forms interface with progress tracking and detailed logging
- **Direct Integration**: Utilizes SolidWorks API for reliable and efficient operation

## 🛠️ Technical Details

The application employs sophisticated algorithms to process SolidWorks assemblies:

1. **Bounding Box Analysis**
   - Calculates each component's dimensions (X, Y, Z)
   - Determines the shortest dimension
   - Automatically orients the view perpendicular to the shortest dimension
   - Ensures optimal visibility of the component's features

2. **Sheet Metal Handling**
   - Automatically detects sheet metal components
   - Applies flattening before export
   - Preserves bend lines and annotations

3. **Duplicate Management**
   - Maintains a registry of processed components
   - Prevents redundant processing of identical parts
   - Optimizes overall processing time

## 📋 Requirements

- Windows Forms development workload
- .NET 8.0 SDK
- x64
- SolidWorks Interop assemblies:
SolidWorks.Interop.sldworks.dll
SolidWorks.Interop.swconst.dll

## 🚦 Getting Started

1. Clone the repository
2. Open the solution in Visual Studio
3. Build the project
4. Run the application
5. Select your SLDASM file using the file picker
6. Click "Convert" to begin the process
7. Find the exported DWG files in the "CONVERTED_DRAWINGS" folder

## 💡 Usage Example

```csharp
// Example code showing how to use the converter programmatically
string assemblyPath = @"C:\Path\To\Your\Assembly.SLDASM";
var converter = new SolidWorksDrawingConverter();
converter.ConvertAssembly(assemblyPath);
```

## ⚙️ Configuration

The application comes with default settings optimized for most use cases. However, you can modify:
- Output directory name
- Export options for sheet metal parts
- View orientation preferences

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
ats
