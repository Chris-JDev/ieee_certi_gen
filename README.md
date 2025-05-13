# IEEE Certificate Generator

A Python script to generate personalized certificates for presenting authors listed in an Excel file, using a background template image and a TrueType font.

## Features

* Reads names and titles from an Excel workbook (`Presenting Authors.xlsx`) ([github.com](https://github.com/Chris-JDev/ieee_certi_gen))
* Uses `image.png` as the certificate background template ([github.com](https://github.com/Chris-JDev/ieee_certi_gen))
* Supports custom fonts via `arial.ttf` ([github.com](https://github.com/Chris-JDev/ieee_certi_gen))
* Auto-wraps and centers text on the template
* Outputs individual certificate PNGs into a `Certificates/` folder

## Prerequisites

* **Python** 3.6 or newer
* **Python packages**: `openpyxl`, `Pillow`

Install dependencies:

```bash
pip install openpyxl pillow
```

## Repository Structure

```
├── Presenting Authors.xlsx   # Excel file with header row and columns: Name, Title ([github.com](https://github.com/Chris-JDev/ieee_certi_gen))
├── arial.ttf                 # TrueType font used for rendering text ([github.com](https://github.com/Chris-JDev/ieee_certi_gen))
├── image.png                 # Certificate template image ([github.com](https://github.com/Chris-JDev/ieee_certi_gen))
└── generate_certificates.py  # Main script to create certificates ([github.com](https://github.com/Chris-JDev/ieee_certi_gen))
```

## Usage

1. **Clone the repository**

   ```bash
   git clone https://github.com/Chris-JDev/ieee_certi_gen.git
   cd ieee_certi_gen
   ```
2. **Verify assets**: Ensure `Presenting Authors.xlsx`, `image.png`, and `arial.ttf` are in the project root.
3. **Install dependencies**

   ```bash
   pip install openpyxl pillow
   ```
4. **Run the generator**

   ```bash
   python generate_certificates.py
   ```
5. **Output**: The script will create a `Certificates/` directory containing one PNG file per entry.

## Customization

* **Text positions**: Adjust `name_y_position` and `title_y_start` in `generate_certificates.py` to align text for your template.
* **Font sizes**: Modify the `size` argument in `ImageFont.truetype` calls for larger or smaller text.
* **Wrapping logic**: Tweak `wrap_text` to handle longer titles or different line spacings.

## License

*No license specified. Feel free to add an appropriate license for your use.*
