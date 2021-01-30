# docsplitter

[![Software License][ico-license]](LICENSE.md)

This small tool, allows you to split a .docx file into multiple files based on headings. Useful for separating a big word file with chapters, into smaller files.

## Motivation

Dealing with docx files is a bit time consuming, and splitting them is pretty annoying. I looked for the libraries online such as python-docx (included in this tool), but they are far from to be ready for practical usage for automatic splitting .docx files.

## Requirements

Python3. If you want to compile for yourself it needs pip and pyinstaller as well.

## Install

### Use precompiled binaries

You can use binaries in dist folder. Those are `dist/docsplitter` for MacOS X or `dist\docsplitter.exe` for Windows 10. You can copy those binaries into your path.

### To compile

0. You can activate venv using `bin/activate`
1. Make sure that you have your requirements in the `requirements.txt` file.
2. You can install the requirements by running this command:

``` bash
$ pip install -r requirements.txt 
```

3. To compile run the following command in the main directory

``` bash
$ pyinstaller --paths=lib --onefile -n docsplitter main.py 
```

Compiled files will appear on the dist folder.

## Usage

When you run `dist/docsplitter` for MacOS X or `dist\docsplitter.exe` for Windows 10, firstly help text will appear.

``` bash
$ ./docsplitter
Usage:  [OPTIONS]

Options:
  -f, --file TEXT      .docx file to split
  -l, --level INTEGER  Heading Level to split the file
  -nn, --noname        If you do not want to have the filename as a prefix for
                       generated docs.

  --help               Show this message and exit
```
### Splitting files

Suppose you have `book with long chapters.docx` file. 


``` bash
$ docsplitter -f  book\ with\ long\ chapters.docx
```

will generate `book with long chapters.zip` in the folder you run the docsplitter. When you extract the zipfile splitted files should appear like `01 - book with long chapters - Chapter Name.docx`

If you don't want the filename as prefix you can use `-nn, --noname` option.

``` bash
$ docsplitter -f  book\ with\ long\ chapters.docx -nn
```

The generated files will not include original filename na will have only headings. `01 - Chapter Name.docx`

If you want to split your word documents using a different heading, you can use   `-l, --level` option. 

## Todo

1. Unit Tests
2. Custom heading selection instead of Heading 1-n levels
3. Tables (Currently tables does not appear in the splitted docs)

## Warning

This library is for experimental purposes. Use at your own risk. Make backups of your files. I take no responisbility in case of any data loss. 

## Testing

Soon.

## Security

If you discover any security related issues, please email mtkocak@gmail.com instead of using the issue tracker.

## Credits

- [Midori Kocak][link-author]

## License

The MIT License (MIT). Please see [License File](LICENSE.md) for more information.

[ico-license]: https://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square
[link-author]: https://github.com/midorikocak
