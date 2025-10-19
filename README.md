# Envelope Printer

A modern web application for printing addresses on envelopes with CSV/Excel file import functionality.

## Features

- **File Import**: Upload CSV or Excel files containing address data
- **Multiple Formats**: Support for various envelope sizes (Standard #10, A4, Legal, Custom)
- **Print Preview**: See how your envelopes will look before printing
- **Modern UI**: Clean, responsive design with intuitive controls
- **Address Validation**: Automatically validates and formats addresses
- **Batch Printing**: Print all envelopes at once or one at a time
- **Customizable**: Adjust font size, font family, and envelope dimensions

## Getting Started

1. Open `index.html` in your web browser
2. Upload a CSV or Excel file with your addresses
3. Configure envelope settings (size, font, etc.)
4. Preview your envelopes
5. Print!

## File Format Requirements

Your CSV or Excel file should contain columns with address information. The application automatically detects common column headers:

- **Name**: Name, Full Name, Recipient, To
- **Address**: Address, Address1, Street, Address Line 1
- **Address2**: Address2, Suite, Apt, Apartment, Address Line 2
- **City**: City, Town
- **State**: State, Province, Region
- **ZIP**: ZIP, ZipCode, Postal, Postal Code, PostCode
- **Country**: Country

### Sample CSV Format

```csv
Name,Address,City,State,ZIP
"John Doe","123 Main Street","New York","NY","10001"
"Jane Smith","456 Oak Avenue","Los Angeles","CA","90210"
```

## Envelope Sizes

- **Standard (#10)**: 4.125" × 9.5" (most common business envelope)
- **A4**: 8.27" × 11.69"
- **Legal**: 8.5" × 14"
- **Custom**: Specify your own dimensions

## Browser Compatibility

This application works in all modern web browsers:
- Chrome 60+
- Firefox 55+
- Safari 11+
- Edge 79+

## Usage Tips

1. **Return Address**: Click on the return address in the preview to edit it
2. **Navigation**: Use the arrow buttons to preview different addresses
3. **Print Quality**: For best results, use high-quality printer settings
4. **Custom Sizes**: Enter dimensions in inches for custom envelope sizes
5. **Font Selection**: Choose fonts that are installed on your system

## Files Included

- `index.html` - Main application interface
- `styles.css` - Modern CSS styling and print layouts
- `script.js` - JavaScript functionality for file processing and printing
- `sample-addresses.csv` - Example CSV file for testing
- `.github/copilot-instructions.md` - Project documentation for AI assistance

## Development

This is a client-side web application with no server requirements. Simply open the HTML file in a web browser to use.

### Technologies Used

- HTML5 for structure
- CSS3 with modern features (Grid, Flexbox, CSS Variables)
- Vanilla JavaScript (ES6+)
- SheetJS library for Excel file processing

## License

This project is open source and available under the MIT License.