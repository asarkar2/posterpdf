#!/usr/bin/env python3

## TODO:
# *) Use argparse for parsing cli arguments
# *) [D] Convert to python3
# *) [D] Use papersize module
# *) [D] Option to print supported papers and their sizes
# *) [D] Add more paper sizes - probably not required any more, since users
#       can input their desired size in the form 'width x height unit'.
# *) [W] Modify code such that output is (md_width + margin) 
#       x (md_height + margin). Otherwise the final product is smaller 
#    than desired. Check what the poster command does.
#    [Won't fix]: Final output of postscript file generated via poster 
#    command seems to be also smaller than the mediasize due to margin.
# *) [D] Check the compatibility of the code with PyPDF2 -> Fully compatibile.
# *) [D] Margin is only required for left and bottom part. 
#       Modify code accordingly.
# *) [D] Create a function to choose smaller of two lengths: xscale, yscale
# *) [D] Margin lines: Some lines appear darker than others. 
#       Why? Not always though. Minor bug. 
#       -> Above behaviour vanished in version 0.4
# *) [D] Write HelpText
# *) [D] Check with input pdf version 1.5
# *) [D] Write docstring for all functions
# *) [D] Clean up
# *) [D] Canvas size support for other page sizes 
# *) [D] fit2size subroutine
# *) [D] check_args subroutine

# To make the program compatible with python2
from __future__ import print_function 
import sys
import os
import re
import random
import platform
import ctypes
from math import ceil
from io import BytesIO
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
import papersize
from decimal import Decimal

def check_args(infile,outfile,overwrite):
    """Function to check for required arguments"""
    
    if ( infile == None ):
        print("Input file not specified.")
        sys.exit()

    if ( not os.path.isfile(infile) ):
        print("Input file \"%s\" doesn't exits." % infile)
        sys.exit()

    if ( outfile == None ):
        print("Output pdf file not specified.")
        sys.exit()

    if (re.search("\.pdf$", outfile.lower()) == None ):
        print("Output file is not a pdf file.")
        sys.exit()

    if ( infile.lower() == outfile.lower() ):
        print("Input and output pdf file cannot be same.")
        sys.exit()

    if ( os.path.isfile(outfile) and overwrite == False ):
        print("Output file \"%s\" exits. Use -f option to overwrite" % outfile)
        sys.exit() 


def get_pdf_dim(ifile):
    """Get the width, height and rotation of a pdf file
    """
    
    pdf = PdfFileReader(open(ifile,"rb"))
    
    page1 = pdf.getPage(0)
    [illx, illy] = [ page1.mediaBox.getLowerLeft_x(),
                     page1.mediaBox.getLowerLeft_y() ]
    [iurx, iury] = [ page1.mediaBox.getUpperRight_x(),
                     page1.mediaBox.getUpperRight_y()]
    pdfw = iurx - illx
    pdfh = iury - illy

    # Paper orientation is in landscape mode, if width >= height
    if ( pdfw >= pdfh ):
        pdfr = papersize.LANDSCAPE
    else:
        pdfr = papersize.PORTRAIT
    
    [pdfw, pdfh] = [round(pdfw,2), round(pdfh,2)]
    return(pdfw,pdfh,pdfr)


def grab_units(dim):
    """Grab the units if dimension is of the form: 'xval x yval units'
    """

    if ( re.search("\d+\s*[iI][nN]\s*$",dim.lower()) != None ):
        iunits = 'in'
    elif ( re.search("\d+\s*[cC][mM]\s*$",dim.lower()) != None ):
        iunits = 'cm'
    elif ( re.search("\d+\s*[mM][mM]\s*$",dim.lower()) != None ):
        iunits = 'mm'
    elif ( re.search("\d+\s*[pP][tT]\s*$",dim.lower()) != None ):
        iunits = 'pt'
    else:
        print("Aborting since unable to decipher the unit")
        sys.exit()
    return iunits


def list_supported_papers():

    for paper in sorted(papersize.SIZES):
        print(paper,':',papersize.SIZES[paper])


def get_paper_dim(paper_name, rotate, units = 'pt'):
    """Get paper dimensions. Default unit is pt. If in, cm or mm is
    specified, then length is converted accordingly."""
    
    paper_name = paper_name.lower()

    paper_size = papersize.SIZES[paper_name]

    if ( rotate == papersize.PORTRAIT ):
        pgwd, pght = papersize.parse_papersize(paper_size,'pt')
    else:
        pgwd, pght = papersize.rotate(
            papersize.parse_papersize(paper_size,'pt'),papersize.LANDSCAPE)

    return (pgwd, pght)


def findratio(len1,len2):
    """Find the ratio of two lengths"""

    scl = float(len1) / float(len2)
    scl = round(scl,2)
    return scl 


def choose_shorter_len(len1,len2):
    """Choose shorter length"""

    ## Choose the smaller of the two scales
    if ( len1 <= len2 ):
        short_len = len1 
    else:
        short_len = len2

    return short_len


def get_percent_margin(md_dim,ct_prcnt):
    """Get the margin in pts if mediasize and cut percent is specified."""

    ct_prcnt = float(ct_prcnt)
    (w,h) = (md_dim[0], md_dim[1])
    mrgn = ct_prcnt * float(w) / 100 ;

    return mrgn ;


def get_margin(md_dim, ct_prcnt, usr_mrgn = None):
    """Get the margin in pts depending on whether usr_mrgn is specified
    or not."""

    if ( usr_mrgn != None ):

        ## If margin is defined by user
        if ( re.search("%",usr_mrgn) != None ):
            ct_prcnt = usr_mrgn.replace('%','').strip()
            mrgn = get_percent_margin(md_dim,ct_prcnt)
        else:
            mrgn = papersize.parse_length(usr_mrgn,'pt')
    else:
        ## No margin specified by user
        mrgn = get_percent_margin(md_dim,ct_prcnt) 

    mrgn = round(mrgn,2)
    return float(mrgn)


def draw_margins(pgwidth,pgheight,allx,ally,aurx,aury,mrgn,i,j):
    """Draw the margin lines and text outside the margin lines"""

    font        = 'Times-Roman'
    fontsize    = 15
    xoffset     = 10
    yoffset     = 10
    dashOn      = 5
    dashOff     = 7
    linewidth   = 0.25
    redColor    = 0.68
    greenColor  = 0.68
    blueColor   = 0.68

    packet = BytesIO()
    # move to the beginning of the StringIO buffer
    packet.seek(0)
    can = canvas.Canvas(packet)
    can.setPageSize((pgwidth, pgheight))
    
    msg = "Grid (" + str(i) + "," + str(j) + ")"
    can.setFont(font,fontsize)
    can.setFillColorRGB(redColor,greenColor,blueColor)
    
    can.drawString(float(allx) + float(mrgn) + xoffset, 
                   float(ally) + float(mrgn)/3, msg)

    # arg1 is length of a dash, arg2 is distance between the two dashes.
    can.setDash(dashOn, dashOff)
    can.setLineWidth(linewidth)
    can.setStrokeColorRGB(redColor,greenColor,blueColor)
    can.line(float(allx), float(ally) + float(mrgn), 
             float(aurx), float(ally) + float(mrgn)) # bottom margin
    can.line(float(allx) + float(mrgn), float(ally),
             float(allx) + float(mrgn), float(aury)) # left margin
#     can.line(float(allx), float(aury) - float(mrgn), 
#              float(aurx), float(aury) - float(mrgn)) # top margin
#     can.line(float(aurx) - float(mrgn), float(ally),
#              float(aurx) - float(mrgn), float(aury)) # right margin
    
    can.save()
    txt_pdf = PdfFileReader(packet)

    return txt_pdf


def get_page_dim(iwidth, iheight, irotate, pgsz_cstm = None):
    """Get the paper dimension"""

    if ( pgsz_cstm == None ):
        ## No custom paper supplied. Set them to default.
        [owidth, oheight] = [iwidth, iheight]
    else: 
        ## Paper size is defined by user
        if ( re.search("\d*\s*x\s*\d*",pgsz_cstm) ):
            ## Width and height supplied as 'w u x h u' 
            ## where w is the width, h is the height and, u is the unit.
            [owidth, oheight] = papersize.parse_papersize(pgsz_cstm,'pt')
        else:
            ## Papersize supplied
            [owidth, oheight] = papersize.rotate(
                papersize.parse_papersize(pgsz_cstm, 'pt'), irotate)

    return(owidth, oheight)


def toposter(ifile = None, md_dim = None, pstr_dim = None, 
    mrgn = None, ofile = None):
    """Slice the poster to multipage pdf file"""

    ## Get dimension of input pdf
    [ifile_width, ifile_height, ifile_rotate] = get_pdf_dim(ifile) 

    ## Get dimension of the postersize
    (pstr_width, pstr_height) = (pstr_dim[0], pstr_dim[1])

    ## Width and height of media size paper, based on whether input pdf
    ## is landscape (rotated) or not.
    (md_width, md_height) = (md_dim[0], md_dim[1])

    ## x and y scale for scaling input pdf to poster size
    pxscale = findratio(pstr_width, ifile_width) ;
    pyscale = findratio(pstr_height, ifile_height) ;

    ## Choose the smaller of the two scales
    pscale = choose_shorter_len(pxscale,pyscale)

    ## Page size is actually size of media page and
    ## single side margin (left, and bottom margins)
    pg_width = float(md_width) + mrgn 
    pg_height = float(md_height) + mrgn 

    ## Scale along x and y direction to 
    ## calculate number of pages required 
    ## in x and y direction
    mxscale = findratio(pstr_width, md_width) 
    myscale = findratio(pstr_height, md_height) 

    ## Always round off to ceiling and get the grid size
    imax = int(ceil(mxscale))
    jmax = int(ceil(myscale))

    totpgs = imax * jmax
    print("Total number of pages:", totpgs)

    ## Cut bounding boxes
    cllx = 0 ;
    clly = 0 ;
    curx = float(pstr_width)  / float(mxscale) ;  
    cury = float(pstr_height) / float(myscale) ;

    ## Reset values
    rllx = cllx ;      
    rurx = curx ;

    ## Increment values
    xinc = curx ;
    yinc = cury ;

    ## Create the output pdf
    pdfout  = PdfFileWriter()
    
    print('Creating page(s): ', end='')
    pagenum = 1 ;
    for i in range(1, imax + 1):

        # Reset x
        cllx = rllx ; 
        curx = rurx ;

        for j in range(1, jmax + 1):

            if pagenum == 1:
                print('%s' % (pagenum), end='')
            else:
                print(',%s' % (pagenum), end='')

            ## Offset values with margins
            [mllx, mlly] = [cllx - float(mrgn), clly - float(mrgn)] 
            [murx, mury] = [cllx + float(md_width), clly + float(md_height)] 
            
            ipdf = PdfFileReader(open(ifile,'rb'))
            page_out = ipdf.getPage(0)
            page_out.scale(pscale,pscale)
            page_out.mediaBox.lowerLeft  = (mllx, mlly)
            page_out.mediaBox.upperRight = (murx, mury)

            ## If margin is not zero, draw margins and add text outside
            ## margin
            if (mrgn != 0.00):
                text_pdf = draw_margins(pg_width,pg_height,
                                        mllx,mlly,murx,mury,mrgn,i,j)
                page_out.mergePage(text_pdf.getPage(0))

            pdfout.addPage(page_out) 

            ## x inc
            cllx = cllx + xinc 
            curx = curx + xinc 

            pagenum = pagenum + 1

        ## y inc
        clly = clly + yinc 
        cury = cury + yinc 
        
    print('')

    outputStream = open(ofile, "wb")
    pdfout.write(outputStream)
    outputStream.close()


def file_rename(src, dst):
    # http://stackoverflow.com/questions/13025313/why-os-rename-is-raising-an-exception-in-python-2-7

    if platform.system() == 'Windows':
        # MOVEFILE_REPLACE_EXISTING = 0x1; MOVEFILE_WRITE_THROUGH = 0x8
        ctypes.windll.kernel32.MoveFileExW(src, dst, 0x1)
        os.unlink(src)
    else:
        os.rename(src, dst) 


def fit2size(ifile,md_dim):
    """Try to fit the input pdf file to mediasize"""
    
    ## Get the pdf dimension
    [ifile_width, ifile_height, ifile_rotate] = get_pdf_dim(ifile)
    
    ## Get output paper size
    (ofile_width, ofile_height) = (md_dim[0], md_dim[1])

    ## Find the scale to enlarge / squeeze the input file size 
    ## to output file size
    fxscale = findratio(ofile_width,  ifile_width) 
    fyscale = findratio(ofile_height, ifile_height) 

    ## Choose the smaller of the two scales
    fscale = choose_shorter_len(fxscale,fyscale)
    
    ipdf = PdfFileReader(open(ifile,'rb'))
    pgcount = ipdf.getNumPages()

    tmpfile = str(random.randint(10000,20000)) + '.pdf'
    tmppdf = PdfFileWriter()

    for pg in range(0, pgcount):
        page_out = ipdf.getPage(pg)
        page_out.scale(fscale,fscale)
        tmppdf.addPage(page_out)

    outputStream = open(tmpfile,"wb")
    tmppdf.write(outputStream)
    outputStream.close()
    
    file_rename(tmpfile, ifile)


def helptext(sname,athr,ver,md_dflt,ct_prcnt):

    print("Script to split poster pdf files into multi-page pdf file with")
    print("margins, such that the file can be printed on ordinary printer.")
    print("Author: %s" % athr)
    print("Version: %s" % ver)
    print("Usage: %s [options] input.pdf -o output.pdf" % sname)
    print("")
    print("Options:")
    print("-h|--help          Show this help and exit.")
    print("-m|--media         Specify output media paper size. Default: %s."
            % md_dflt)
    print("-p|--poster        Specify poster size. Default is the size of")
    print("                   the input pdf file.")
    print("-o|--output        Specifiy output pdf file.")
    print("-f|--force         Ovewrite output file if it exists.")
    print("-c|--cut-margin    Specify cut margins either as percentage of the")
    print("                   width of the media paper or in in|mm|cm|pt.")
    print("                   (1 in = 25.4 mm = 2.54 cm = 72 pt)")
    print("                   Default: %s%% of width of media paper." 
            % ct_prcnt)
    print("-l                 List the supported paper and their sizes.")
    


if __name__ == '__main__':

    scriptname = os.path.basename(sys.argv[0])
    author = 'Anjishnu Sarkar'
    version = '0.11'
    cut_percent = 5     # In percentage
    media_default = 'a4'
    media_custom = None
    infile = None
    outfile = None
    poster_custom = None
    user_margin = None
    overwrite = False 
  
    ## Parse through cli arguments
    numargv = len(sys.argv)-1
    nargv = 1
    
    while nargv <= numargv:
      if sys.argv[nargv] == "-h" or sys.argv[nargv] == "--help":
          helptext(scriptname,author,version,media_default,cut_percent)
          sys.exit(0)
  
      elif re.search('\.pdf$',sys.argv[nargv].lower()):
          infile = sys.argv[nargv] 
  
      elif sys.argv[nargv] == "-o" or sys.argv[nargv] == "--output":
          outfile = sys.argv[nargv+1]
          nargv += 1
  
      elif sys.argv[nargv] == "-f" or sys.argv[nargv] == "--force":
          overwrite = True
  
      elif sys.argv[nargv] == "-m" or sys.argv[nargv] == "--media":
          media_custom = sys.argv[nargv+1]
          nargv += 1 
  
      elif sys.argv[nargv] == "-p" or sys.argv[nargv] == "--poster":
          poster_custom = sys.argv[nargv+1]
          nargv += 1 
  
      elif sys.argv[nargv] == "-c" or sys.argv[nargv] == "--cut-margin":
          user_margin = sys.argv[nargv+1]
          nargv += 1 
  
      elif sys.argv[nargv] == "-l":
          list_supported_papers()
          sys.exit()
  
      else:
          print("%s: Unspecified option. Aborting." % (scriptname))
          sys.exit(1)
  
      nargv += 1 
  
    ## Check for required arguments
    check_args(infile,outfile,overwrite)
  
    ## Get input pdf dimensions
    (ifile_width, ifile_height, ifile_rotate) = get_pdf_dim(infile) 
  
    ## Get the poster dimensions
    poster_dim = get_page_dim(ifile_width, ifile_height,
                              ifile_rotate, poster_custom)
    
    ## Get the dimensions of the default media paper
    (media_dflt_width, media_dflt_height) = get_paper_dim(
                                            media_default,ifile_rotate)
  
    ## Get the dimensions of the custom media paper, if any
    media_dim = get_page_dim(media_dflt_width, media_dflt_height,
                             ifile_rotate, media_custom)
  
    ## Get margin
    margin = get_margin(media_dim,cut_percent,user_margin)
    
    ## Crop the input pdf to multipage pdf file of mediasize
    toposter(ifile = infile, 
        md_dim = media_dim,
        pstr_dim = poster_dim,
        mrgn = margin,
        ofile = outfile
    )
    
    ## Resize the file so that the file is actually of mediasize
    print("Resizing the pdf to mediasize")
    fit2size(outfile,media_dim)
  
    (outpdf_width, outpdf_height, outpdf_rotate) = get_pdf_dim(outfile) 
    (media_width, media_height) = (media_dim[0], media_dim[1])
  
    print("Required dimension: %6.2f pt x %6.2f pt" 
            % (media_width, media_height))
    print("Output dimension  : %6.2f pt x %6.2f pt" 
            % (outpdf_width, outpdf_height))

