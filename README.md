# posterpdf
Python script to split poster size pdf files into multi-page pdf file of smaller paper size (Eg: A4) with **margins**, such that the file can be printed on ordinary printer and then later pasted together.

## Example:

    # To slice the a2 poster into a4 paper
    posterpdf.py old.pdf -o new.pdf

    # To change the poster size from a1 to a0
    posterpdf.py -p a0 -m a0 -c 0% old.pdf -o new.pdf

## Usage:

    posterpdf.py [options] input.pdf -o output.pdf

    Options:
    -h|--help          Show this help and exit.
    -p|--poster        Specify poster size. Default is the size of
                       the input pdf file.
    -m|--media         Specify output media paper size. Default: a4.
    -o|--output        Specifiy output pdf file.
    -f|--force         Ovewrite output file if it exists.
    -c|--cut-margin    Specify cut margins either as percentage of the
                       width of the media paper or in in|mm|cm|pt.
                       (1 in = 25.4 mm = 2.54 cm = 72 pt)
                       Default: 5% of width of media paper.
    -l                 List the supported paper and their sizes.
