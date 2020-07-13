# officex-tweaker
Simple console application to tweak slightly OfficeX files

## Usage
officex-tweaker 1.0.0
Copyright (C) 2020 officex-tweaker

  -i, --input        Required. Path to OfficeX file to tweak.

  -o, --output       Output path of tweaked OfficeX file. If not exists, will be <input path>.out.<ext>

  -w, --overwrite    (Default: false) Ignore output file path and overwrite input path.

  --help             Display this help screen.

  --version          Display version information.

### Usage Examples:
`officex-tweaker.exe -i c:\temp\example.docx` - will create a tweaked file at c:\temp\example.out.docx

`officex-tweaker.exe -i c:\temp\example.docx -o c:\temp\output\tweaked.docx` - will create a tweaked file at c:\temp\output\tweaked.docx

`officex-tweaker.exe -i c:\temp\example.docx -w` - will tweak the file and overwrite it at c:\temp\example.docx
