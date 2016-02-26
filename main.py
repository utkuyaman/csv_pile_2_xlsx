#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
Created on Nov 21, 2015

@author: tuku
"""

import argparse, sys, os
import xlsxwriter
import csv

if __name__ == '__main__':

    parser = argparse.ArgumentParser()

    parser.add_argument("-s", "--source", help="source folder that contains csv files", type=str)
    parser.add_argument("-d", "--destination", help="final file name", type=str)
    parser.add_argument("-f", "--force", help="force final name override", action="store_true")
    parser.add_argument("-v", "--verbose", help="toggle logging output", action="store_true")

    args = parser.parse_args()

    verbose = args.verbose

    if not args.source or not args.destination:
        parser.print_help()
        sys.exit(-1)

    # check if folder exists
    if not os.path.exists(args.source) or not os.path.isdir(args.source):
        print('Source must be a folder')
        parser.print_help()
        sys.exit(-1)

    # check if destination file exists or not
    if os.path.exists(args.destination) and not args.force:
        print('Destination file already exists, try using -f argument or choosing another file name')
        parser.print_help()
        sys.exit(-1)

    csv_files = [f for f in os.listdir(args.source) if os.path.isfile(os.path.join(args.source, f))]

    if verbose:
        print('found {} files'.format(len(csv_files)))

    # Create an new Excel file
    with xlsxwriter.Workbook(args.destination) as workbook:

        for file in csv_files:
            if verbose:
                print('merging file "{}"'.format(file))

            # add a worksheet for each csv file
            worksheet = workbook.add_worksheet(file[:31])

            with open(os.path.join(args.source, file), 'r') as csvfile:
                reader = csv.reader(csvfile, delimiter='\t')

                row_index = 0
                for row in reader:
                    worksheet.write_row(row_index, 0, tuple(row))
                    row_index += 1

            if verbose:
                print('merged {} rows'.format(row_index))

        if verbose:
            print('saving file')
