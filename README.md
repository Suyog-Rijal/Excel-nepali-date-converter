# Date Converter - Excel Add-in

A powerful Microsoft Excel add-in for converting dates between Nepali (Bikram Sambat) and Gregorian (AD) calendar systems. This add-in provides real-time date conversion capabilities directly within Excel, making it easy to work with Nepali dates in your spreadsheets.

## 🌟 Features

- **Real-time Date Conversion**: Automatically converts dates as you type or modify cells
- **Bidirectional Conversion**: Convert from BS (Bikram Sambat) to AD (Gregorian) and vice versa
- **Column-based Processing**: Select specific columns for date conversion
- **Smart Date Range Lock**: Automatically handles date range validation to prevent invalid conversions
- **Excel Integration**: Seamlessly integrates with Excel's interface and functionality
- **Modern UI**: Built with React and Fluent UI components for a professional look

## 📋 Prerequisites

- Microsoft Excel (Desktop version)
- Windows operating system
- Internet connection (for initial setup and date conversion API calls)

## 🚀 Installation

### Option 1: Using the NDC.bat Installer (Recommended)

1. **Download the NDC.bat file**:
   - [Download NDC.bat](NDC.bat) (Right-click and "Save As")
   - Or copy the file from this repository

2. **Run the installer**:
   - Double-click `NDC.bat` to run the installer
   - Choose option `1` to install the add-in
   - Select option `1` again to download from URL (recommended)
   - The installer will automatically download and install the manifest

3. **Restart Excel**:
   - Close and reopen Excel
   - The add-in will appear in the "Insert" tab under "My Add-ins"

### Option 2: Manual Installation

1. **Download the manifest file**:
   - Download `manifest.xml` from this repository
   - Or access it online: [manifest.xml](https://excel-nepali-date-converter.vercel.app/manifest.xml)

2. **Install via Registry**:
   - Open Registry Editor (regedit)
   - Navigate to: `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`
   - Create a new String Value named "Manifest"
   - Set the value to the full path of your manifest.xml file

3. **Restart Excel**:
   - Close and reopen Excel to see the add-in

## 📖 Usage

### Getting Started

1. **Open the Add-in**:
   - In Excel, go to the "Insert" tab
   - Click "My Add-ins" and select "Date-Converter"
   - The taskpane will open on the right side

2. **Configure Settings**:
   - **Select Column**: Choose the column containing dates to convert
   - **Conversion Operation**: Select "BS to AD" or "AD to BS"
   - **Date Range Lock**: Choose "Auto" for automatic range validation

3. **Start Monitoring**:
   - Click the "Start Monitoring" button

### Supported Date Formats
- **Excel Serial Numbers**: Standard Excel date serial numbers
- **MM/DD/YYYY**: Date format with slashes (e.g., 12/25/2023)

### Conversion Examples
- **BS to AD**: Convert Nepali dates to Gregorian dates
- **AD to BS**: Convert Gregorian dates to Nepali dates

## 🛠️ Development

### Project Structure

```
Date-Converter/
├── src/
│   ├── taskpane/
│   │   ├── components/
│   │   │   ├── App.tsx          # Main application component
│   │   │   └── CustomDropdown.tsx # Custom dropdown component
│   │   ├── index.tsx            # Entry point
│   │   └── taskpane.html        # HTML template
│   └── commands/
│       ├── commands.ts          # Command handlers
│       └── commands.html        # Command UI
├── assets/                      # Icons and images
├── manifest.xml                 # Office add-in manifest
├── NDC.bat                      # Windows installer script
└── package.json                 # Project dependencies
```

### Technology Stack

- **Frontend**: React 18, TypeScript
- **UI Framework**: Fluent UI React Components
- **Build Tool**: Webpack
- **Office Integration**: Office.js API
- **Backend API**: fastapi (hosted on Render)

### Development Setup

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd Date-Converter
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Start development server**:
   ```bash
   npm run dev-server
   ```

4. **Build for production**:
   ```bash
   npm run build
   ```

## 🔧 Configuration

### API Endpoints

The add-in uses the following API endpoints for date conversion:
- **BS to AD**: `https://excel-nepali-date-converter-backend.onrender.com/bs-to-ad`
- **AD to BS**: `https://excel-nepali-date-converter-backend.onrender.com/ad-to-bs`

### Registry Settings

The add-in is installed via Windows Registry at:
```
HKCU\Software\Microsoft\Office\16.0\WEF\Developer
```

## 🚨 Troubleshooting

### Common Issues

1. **Add-in not appearing in Excel**:
   - Restart Excel after installation
   - Check if the manifest file path is correct in registry
   - Ensure you have the correct Office version

2. **Date conversion not working**:
   - Check your internet connection
   - Verify the date format is supported
   - Ensure the date is within the valid range

3. **Installation errors**:
   - Run the NDC.bat as Administrator
   - Check Windows permissions
   - Verify the manifest file is accessible

### Uninstallation

To remove the add-in:
1. Run `NDC.bat` and select option `2` (Uninstall Add-in)
2. Or manually delete the registry key: `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`
3. Restart Excel

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📞 Support

For support and questions:
- Create an issue in the repository
- Contact the development team
- Check the troubleshooting section above

## 🔄 Version History

- **v1.0.0**: Initial release with basic date conversion functionality
- **v1.0.1**: Added date range validation and improved error handling

---

**Note**: This add-in requires an active internet connection for date conversion API calls. The backend service is hosted on Render and provides the conversion logic for Nepali calendar dates. 
