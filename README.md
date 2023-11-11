# Moniker

A command line tool for Windows for working with Microsoft Word documents.

#### usage for batch mode:
```
  moniker.exe batch -i [file] -o [directory] -s [style] -m [mapping file]
    -i - input file
    -o - output directory
    -s - style name to match for splitting
    -m - font-style mapping file (comma delimited)
```

#### usage for splitting a Word document:
```
  moniker.exe split -i [file] -o [directory] -s [style]
    -i - input file
    -o - output directory
    -s - style name to match for splitting
```

#### usage for replacing a single font with a style in a Word document:
```
  moniker.exe rfs -i [file] -f [font] -s [style]
    -i - input file
    -f - font name
    -s - style name
```

#### usage for replacing multiple fonts with styles in a Word document:
```
  moniker.exe rfs -i [file] -m [mapping file]
    -i - input file
    -m - font-style mapping file (comma delimited)
```

#### usage for processing cross-reference tags in a Word document:
```
  moniker.exe cref -i [file] -tcs [table caption style] -fgs [figure caption style] -xrs [xref style]
    -i - input file
    -tgs - style for table captions
    -fgs - style for figure captions
    -xrs - style for cross-references
```

#### usage for creating a paragraph style in a Word document:
```
  moniker.exe cps -i [file] -s [style] {-sf [style font]} {-ss [style font size]} {-sb} {-si} {-sc [style color]
    -i - input file
    -s - style name
    -sf  style font name (if style doesn't exist) - OPTIONAL
    -ss  style font size (if style doesn't exist) - OPTIONAL
    -sb  style bold (if style doesn't exist) - OPTIONAL
    -si  style italic (if style doesn't exist) - OPTIONAL
    -sc  style font color (if style doesn't exist) - OPTIONAL
```

#### usage for creating a paragraph style in a Word document:
```
  moniker.exe ccs -i [file] -s [style] {-sf [style font]} {-ss [style font size]} {-sb} {-si} {-sc [style color]
    -i - input file
    -s - style name
    -sf  style font name (if style doesn't exist) - OPTIONAL
    -ss  style font size (if style doesn't exist) - OPTIONAL
    -sb  style bold (if style doesn't exist) - OPTIONAL
    -si  style italic (if style doesn't exist) - OPTIONAL
    -sc  style font color (if style doesn't exist) - OPTIONAL
```

#### usage for replacing tables with tags in a Word document:
```
  moniker.exe ttg -i [file]
    -i - input file
```