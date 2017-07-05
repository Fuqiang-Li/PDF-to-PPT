# pdf to pptx using python

Written in PYTHON 2.7
Tested in ubuntu 16 (Linux)

# Usage:

## Make sure you have installed the required packages

```
cd ppt-pdf/
sudo source requirement.sh
```
## Basic usage:
```
python   cli_pdf_to_ppt.py   a.pdf
```
output file will be created as a.pptx in the same location

## More details:
python cli_pdf_to_ppt.py [-h] [-g] [-b]
                         [-t]
                         [-i] [-c CROP [CROP ...]]
                         pdf_file
                         
onvert PDF to pptx

positional arguments:
  pdf_file              path to the pdf file

optional arguments:
  -h, --help            show this help message and exit
  -g, --greyscale       generate greyscale pptx
  -b, --blackwhite      generate black and white pptx
  -t, --threshold       the threshold value [0,255) for converting black and white
  -i, --invert          invert generated pptx color
  -c CROP [CROP ...], --crop CROP [CROP ...]
                        values specifing region(s) in pixel to crop

For example:
```
python cli_pdf_to_ppt.py 10.pdf -b -i -c '328,361,1326,1109' '328,1231,1326,1979' -t 50
```
