# Excel Component Enrichment Tool

## Overview
This tool enhances Excel spreadsheets containing electronic component data by automatically enriching component descriptions and sourcing information using AI-powered analysis. It processes electronic components data and standardizes descriptions according to industry conventions.

## Features
- **Excel File Processing**: 
  - Supports XLSX files with multiple sheets
  - Extracts sheet names and column headers
  - Handles large data volumes efficiently
- **Automated Component Analysis**: 
  - Component type detection
  - Specification standardization
  - Source URL extraction
  - Additional component information search using OpenAI
- **Real-time Processing**:
  - Progress tracking
  - Preview of changes
  - Intermediate results saving
- **Standardized Output Format**:
  - Consistent component descriptions
  - Primary and secondary source URLs
  - Full search results preservation

## Technical Requirements
- Node.js (v18 or higher)
- Dependencies:
  - ExcelJS
  - OpenAI API client
  - dotenv
  - WebSocket support

## Installation
1. Clone the repository:
```bash
git clone https://github.com/your-repo/excel-processing-app.git
cd excel-processing-app
```

2. Install dependencies:
```bash
npm install exceljs openai dotenv ws
```

3. Configure environment variables:
Create `.env` file in the root directory:
```ini
OPENROUTER_API_KEY=your_api_key_here
```

4. Start the application:
```bash
npm start
```

## API Reference

### Main Processing Function
```typescript
processExcelBuffer(
  buffer: Buffer,           // Excel file buffer
  sheetName: string,        // Sheet name
  partNumberColumn: number, // Part number column (1-based)
  descriptionColumn: number,// Description column (1-based)
  callbacks: {
    onProgress: (current: number, total: number) => void,
    onPreview: (before: string, after: string, source: string) => void
  },
  fileId: string           // Unique file identifier
): Promise<Uint8Array>

// Usage Example:
import { processExcelBuffer } from './path-to-module';
import fs from 'fs';

const buffer = fs.readFileSync('components.xlsx');

processExcelBuffer(
  buffer, 
  'Sheet1', 
  1, 
  2,
  {
    onProgress: (current, total) => console.log(`Processed ${current} of ${total} rows`),
    onPreview: (before, after, source) => console.log(`Changed: ${before} → ${after} (source: ${source})`)
  },
  'unique-file-id'
)
.then(enrichedBuffer => {
  fs.writeFileSync('output.xlsx', enrichedBuffer);
  console.log('File processed successfully');
})
.catch(error => {
  console.error('Processing error:', error);
});
```

### Utility Functions
```typescript
// Get list of sheets
getSheetNames(buffer: Buffer): Promise<string[]>

// Get file headers
getFileHeaders(buffer: Buffer): Promise<string[]>

// Get headers from specific sheet
getFileHeadersFromSheet(buffer: Buffer, sheetName: string): Promise<string[]>

// File management
getAvailableFiles(fileId: string): Array<{ name: string, timestamp: string, size: number }>
getIntermediateFile(fileId: string, fileName: string): Buffer | null
```

## Processing Features

### Data Processing Steps
1. **Excel File Reading**: Extract sheet names and column headers
2. **Component Information Search**: Use OpenAI API to find specifications and sources
3. **Description Formatting**: Standardize and format component descriptions
4. **Post-processing**: Additional cleanup and standardization
5. **Results Saving**: Write enriched data to new Excel file with additional columns

### Component Information Processing
1. **Description Formatting:**
   - Automatic component type detection
   - Unit standardization (Ω → R, uF → MF)
   - Industry standard compliance

2. **Source Management:**
   - Search across 20+ electronic component databases
   - URL validation
   - Alternative source selection

3. **Error Handling:**
   - Network failure retry mechanism
   - Partial results preservation
   - Detailed processing logs

### OpenAI Integration
The application uses OpenAI API for component information search and standardization:

1. **Information Search**:
   - Model: perplexity/llama-3.1-sonar-small-128k-online
   - Temperature: 0.1
   - System: Instructions for finding specifications and URLs
   - User: Query with part number and description

2. **Description Formatting**:
   - Model: microsoft/phi-4
   - Temperature: 0.1
   - System: Instructions for formatting search results
   - User: Original description and search results

### Supported Component Types
- Electronic Components (CAP, RES, CONN, etc.)
- Mechanical Parts (SCREW, BOLT, NUT, etc.)
- Chemical Materials (ADHESIVE, EPOXY, etc.)

### Standardization Rules
- Value formats (resistance, capacitance, etc.)
- Package types
- Mounting types (SMT/THT)
- Temperature coefficients
- Voltage/current ranges

## Output Format
The enriched Excel file includes:
1. Original data
2. Enhanced columns:
   - Enriched Description
   - Primary Source
   - Secondary Source
   - Full Search Result

## AI Model Support
The tool supports all models available through the OpenRouter API. By default, it uses:
- Perplexity/llama-3.1-sonar-small-128k for search operations
- Microsoft/phi-4 for formatting
Other compatible models can be configured as needed.

## Performance Considerations
For workloads >1000 rows, recommended setup:
- 2+ GB memory allocation
- WebSocket data transmission
- Batch processing (100 rows per batch)

## Error Handling
The application includes mechanisms for error handling and logging:
- File read/write errors are handled with console output
- API errors are logged with retry mechanism
- Intermediate errors allow continuing processing of remaining data

## API Limitations
- ~3 requests/second
- 1500 tokens per response
- 5 retry attempts on errors

## Contributing
Contributions are welcome! Please:
1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## Support
For issues and feature requests, please:
- Use the GitHub issue tracker: https://github.com/Mavline
- Contact: mavlinex@gmail.com

## Acknowledgments
- OpenRouter AI for API support
- ExcelJS for Excel file processing
- WebSocket for real-time updates
