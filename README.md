# Server QC Data Collection Tool

An automated quality control data collection tool for server systems. This tool gathers hardware specifications, processes QC grades, and generates standardized Excel reports with detailed system information.

## Features

- Automated hardware detection
- Manual verification and override
- Standardized grading system
- Excel report generation
- Detailed component analysis
- User input validation

## Prerequisites

### System Requirements
- Linux-based operating system
- sudo privileges
- Required system tools:
  - dmidecode
  - lspci
  - free

### Python Dependencies
```bash
pip install openpyxl
```

## Component Detection

Automatically detects:
- CPU model and specifications
- RAM type and capacity
- PCIe devices
- System model and serial numbers
- Product numbers
- Storage controllers
- Network adapters

## Grading System

### Cosmetic Grades
- Format: C1-C9
- Validation enforced
- Required field

### Functional Grades
- Format: F1-F9
- Validation enforced
- Required field

## Usage

Run the script with sudo privileges:
```bash
sudo python server_qc_tool.py
```

## Output

Generates Excel file with:
- Hardware specifications
- QC grades
- Technician information
- Test date
- Component details

### File Format
- Named: `Server_QC_Report_YYYYMMDD_HHMMSS.xlsx`
- Formatted cells
- Auto-sized columns
- Standardized layout

## Data Validation

- Grade format checking
- Date format validation
- Required field enforcement
- Manual override options
- Data verification prompts

## Error Handling

- Hardware detection fallbacks
- Permission checking
- Input validation
- File write verification

## Best Practices

1. Run with sudo privileges
2. Verify detected information
3. Use consistent grading standards
4. Complete all fields
5. Save reports in dedicated folder

## Troubleshooting

Common issues:
- Permission denied: Run with sudo
- Missing hardware info: Verify dmidecode access
- Invalid grades: Check format (C1-C9, F1-F9)

## Contributing

Feel free to submit issues and enhancement requests!

## License

[Specify your license here]

## Acknowledgments

- OpenPyXL developers
- Linux system tool maintainers
