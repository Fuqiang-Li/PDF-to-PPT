# pdf to pptx using python
This is an folk from vijayanandrp/PDF-to-PPT.  
Extended with following supports to generate pptx:
- in greyscale
- in black and white
- in inversed color 
- based on cropped regions, which is useful when converting pdf with multi-slides in a single page

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
.........................[-t]  
.........................[-i] [-c CROP [CROP ...]]  
.........................pdf_file  
                         
convert PDF to pptx

positional arguments:  
&nbsp;&nbsp;pdf_file              path to the pdf file

optional arguments:
&nbsp;&nbsp;-h, --help            show this help message and exit  
&nbsp;&nbsp;-g, --greyscale       generate greyscale pptx  
&nbsp;&nbsp;-b, --blackwhite      generate black and white pptx  
&nbsp;&nbsp;-t, --threshold       the threshold value [0,255) for converting black and white  
&nbsp;&nbsp;-i, --invert          invert generated pptx color  
&nbsp;&nbsp;-c CROP [CROP ...], --crop CROP [CROP ...]    values specifing region(s) in pixel to crop  

For example:
```
python cli_pdf_to_ppt.py a.pdf -b -i -c '328,361,1326,1109' '328,1231,1326,1979' -t 50
```
