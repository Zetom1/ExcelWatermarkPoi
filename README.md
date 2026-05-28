# Excel Watermark POI

Utility built with Java 17 and Apache POI to add watermarks to Excel files.

This project was created to solve a limitation found in most available examples and documentation:
many solutions only support watermarks for printing OR visible watermarks inside the Excel sheet, but not both together.

## Features

- Add visible watermark to Excel sheets
- Support for printable watermark behavior
- Apache POI based implementation
- Java 17 compatible
- Easy integration into existing projects

## Technologies

- Java 17
- Apache POI
- Maven

## Use Case

This project is useful for:

- Confidential Excel reports
- Internal company documents
- Draft spreadsheets
- Automatically generated Excel files

## Example

```java
WatermarkService.addWatermark(
    "input.xlsx",
    "output.xlsx",
    "CONFIDENTIAL"
);
```

## Project Structure

```text
src/
 ├── main/
 │    ├── java/
 │    └── resources/
 └── test/
```

## Getting Started

### Clone repository

```bash
git clone https://github.com/Zetom1/ExcelWatermarkPoi.git
```

### Build project

```bash
mvn clean install
```

### Run project

```bash
mvn spring-boot:run
```

## Screenshots

Add before/after images here if possible.

Example:

| Original File | Watermarked File |
|---|---|
| image | image |

## Motivation

While working with Apache POI, most available documentation and examples only covered partial watermark implementations.
This repository was created as a practical solution that combines both visible and printable watermark behavior.

## Future Improvements

- Custom watermark opacity
- Rotation settings
- Multiple watermark styles
- Image watermark support
- CLI support

## License

MIT
